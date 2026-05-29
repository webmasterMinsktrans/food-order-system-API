import { Module } from '@nestjs/common';
import { MenuParserService } from './menu-parser.service';
import { DocxMenuParser } from './menu-parser-implementations/docx-menu-parser-new';
import { MenuFileParser } from './menu-parser.interface';

@Module({
  providers: [
    MenuParserService,
    {
      provide: 'PARSERS',
      useFactory: () => {
        return [new DocxMenuParser()] as MenuFileParser[];
      },
    },
  ],
  exports: [MenuParserService],
})
export class MenuParserModule {}
