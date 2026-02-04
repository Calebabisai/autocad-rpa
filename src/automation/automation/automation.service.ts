/* eslint-disable */
import { Injectable } from '@nestjs/common';
import * as XLSX from 'xlsx';
import { AUTOCAD_MAPPINGS } from '../../common/constants/coordinates';
@Injectable()
export class AutomationService {
  processProject(fileBuffer: Buffer): string {
    const workbook = XLSX.read(fileBuffer, { type: 'buffer' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];

    const getVal = (cell: string): string => {
      const cellData = sheet[cell];
      return cellData && cellData.v !== undefined
        ? String(cellData.v).trim()
        : '';
    };

    // Extraemos los valores de las celdas que mencionaste
    const tipoConexion = getVal('M57'); // Celda para TESLA POWERWALL 3
    const moveDiagram = getVal('R59'); // Celda que dice YES o NO

    // Preparamos el objeto de datos para la generación del script
    const proyecto = {
      cliente: getVal('C2'),
      esTeslaPW3: tipoConexion.toUpperCase() === 'TESLA POWERWALL 3',
      debeMoverse: moveDiagram.toUpperCase() === 'YES',
    };

    return this.generateScrContent(proyecto);
  }

  /* eslint-disable */
  private generateScrContent(data: any): string {
    let scr = `_REGEN\n`;
    scr += `; --- Procesando Proyecto: ${data.cliente} ---\n`;

    // Solo si es Powerwall 3 Y la celda R59 dice YES
    if (data.esTeslaPW3 && data.debeMoverse) {
      const source = AUTOCAD_MAPPINGS.templates.tesla_backup_riser;
      const dest = AUTOCAD_MAPPINGS.windows.pv3;

      // MARGEN DE ERROR: Creamos un cuadro de 40x40 unidades alrededor del centro
      const margen = 200;
      const x1 = source.x - margen;
      const y1 = source.y - margen;
      const x2 = source.x + margen;
      const y2 = source.y + margen;

      // Calculamos el desplazamiento (Delta) si quieres usar @
      const dx = dest.x - source.x;
      const dy = dest.y - source.y;

      scr += `; Moviendo diagrama de Tesla a PV3...\n`;

      // 1. Zoom al objeto para asegurar que AutoCAD lo "vea"
      scr += `_ZOOM _W ${x1},${y1} ${x2},${y2}\n`;

      // 2. Comando MOVE
      scr += `_MOVE\n`;

      // 3. _C inicia una selección de "Crossing" (Cuadro Verde)
      scr += `_C\n`;

      // 4. Coordenada Esquina Inferior Izquierda del cuadro
      scr += `${x1},${y1}\n`;

      // 5. Coordenada Esquina Superior Derecha del cuadro
      scr += `${x2},${y2}\n`;

      // 6. Enter vacío para "Terminar Selección"
      scr += `\n`;

      // 7. Punto Base (Origen) -> Punto Destino
      scr += `${source.x},${source.y}\n${dest.x},${dest.y}\n`;
    } else {
      scr += `; Condicion no cumplida.\n`;
    }

    scr += `_ZOOM _E\n`;
    // Restauramos los Snaps (opcional, el 4133 es un valor común)
    scr += `_OSMODE 4133\n`;
    return scr;
  }
}
