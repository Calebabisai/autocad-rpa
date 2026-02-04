import {
  Controller,
  Post,
  UploadedFile,
  UseInterceptors,
  Res,
  BadRequestException,
} from '@nestjs/common';
import { FileInterceptor } from '@nestjs/platform-express';
import { AutomationService } from './automation.service';
import type { Response } from 'express';

// Esta línea es clave para que reconozca el tipo de Multer sin errores de metadata
import 'multer';

@Controller('automation')
export class AutomationController {
  constructor(private readonly automationService: AutomationService) {}

  @Post('upload-excel')
  @UseInterceptors(FileInterceptor('file'))
  // Cambiamos el tipo a Express.Multer.File directamente
  uploadExcel(@UploadedFile() file: Express.Multer.File, @Res() res: Response) {
    if (!file || !file.buffer) {
      throw new BadRequestException('Archivo no válido');
    }

    const scrContent = this.automationService.processProject(file.buffer);

    res.setHeader('Content-Type', 'text/plain');
    res.setHeader(
      'Content-Disposition',
      'attachment; filename=automatizacion.scr',
    );

    return res.send(scrContent);
  }
}
