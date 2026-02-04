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

    // LÓGICA IMPORTANTE:
    // Solo si es Powerwall 3 Y la celda R59 dice YES
    if (data.esTeslaPW3 && data.debeMoverse) {
      const source = AUTOCAD_MAPPINGS.templates.tesla_backup_riser;
      const dest = AUTOCAD_MAPPINGS.windows.pv3;

      // Calculamos el desplazamiento (Delta) si quieres usar @
      const dx = dest.x - source.x;
      const dy = dest.y - source.y;

      scr += `; Moviendo diagrama de Tesla a PV3...\n`;
      scr += `_MOVE ${source.x},${source.y}\n\n${source.x},${source.y}\n${dest.x},${dest.y}\n`;
    } else {
      scr += `; Condicion Tesla PW3 o R59=YES no cumplida. No se movio el diagrama.\n`;
    }

    scr += `_ZOOM _E\n`;
    return scr;
  }
}
