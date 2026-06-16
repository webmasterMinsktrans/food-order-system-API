import { Inject, Injectable, NotAcceptableException } from '@nestjs/common';
import { MenuFileParser } from './menu-parser.interface';
import * as path from 'path';

@Injectable()
export class MenuParserService {
  constructor(@Inject('PARSERS') private readonly parsers: MenuFileParser[]) {}

   async parseMenuFile(file: Express.Multer.File) {
    const extName = path.extname(file.originalname).toLowerCase(); 
    
    for (const parser of this.parsers) {
      if (parser.getParsedExtensions().toLowerCase().includes(extName)) {
        return parser.parseFile(file.buffer);
      }
    }
    throw new NotAcceptableException(
      'Файл указанного формата не поддерживается',
    );
  }
}
