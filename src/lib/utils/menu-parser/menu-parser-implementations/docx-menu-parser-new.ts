import {
  menuDeclaration,
  menuPositionDeclaration,
  MenuFileParser,
} from '../menu-parser.interface';
import { BadRequestException, NotAcceptableException } from '@nestjs/common';
import * as WordExtractor from 'word-extractor';
import * as XLSX from 'xlsx';
import * as dayjs from 'dayjs';
import * as CustomParseFormat from 'dayjs/plugin/customParseFormat';
import * as timezone from 'dayjs/plugin/timezone';
import 'dayjs/locale/ru';
import * as utc from 'dayjs/plugin/utc';

export class DocxMenuParser extends MenuFileParser {
  constructor() {
    super();
    dayjs.extend(CustomParseFormat);
    dayjs.extend(utc);
    dayjs.extend(timezone);
  }

  parseFile(filePath: string);
  parseFile(buffer: Buffer);
  async parseFile(buffer: unknown): Promise<menuDeclaration> {
    if (!(buffer instanceof Buffer)) {
      throw new NotAcceptableException('Некорректный формат файла меню');
    }

    // 1. Пытаемся обработать файл как Excel (.xls / .xlsx)
    try {
      // cellNF и cellText заставляют библиотеку считывать маски форматов отображения
      const workbook = XLSX.read(buffer, { 
        type: 'buffer',
        cellNF: true,
        cellText: true
      });
      if (workbook && workbook.SheetNames && workbook.SheetNames.length > 0) {
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        if (firstSheet && firstSheet['!ref']) {
          console.log('Start Excel');
          return this.parseExcel(workbook);
        }
      }
    } catch (e) {
      // Игнорируем ошибки парсинга Excel и идем к Word
    }

    // 2. Если не Excel, обрабатываем как оригинальный Word-файл
    const extractor = new WordExtractor();
    const extracted = await extractor.extract(buffer);
    if (!extracted) throw new BadRequestException();
    try {
      console.log('Start');
      return this.parseDocument(extracted);
    } catch (err) {
      throw new NotAcceptableException(
        err.message ?? 'Некорректный формат файла меню',
      );
    }
  }

  async parseExcel(workbook: XLSX.WorkBook): Promise<menuDeclaration> {
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    
    const rows = XLSX.utils.sheet_to_json<string[]>(worksheet, { 
      header: 1, 
      defval: '',
      raw: false 
    });

    let menuDate = dayjs().tz('Europe/Minsk');
    let dateFound = false;

    const searchRowsCount = Math.min(rows.length, 7);
    for (let i = 0; i < searchRowsCount; i++) {
      for (const cellValue of rows[i]) {
        const val = String(cellValue || '').trim();
        if (val && val.includes('на') && val.includes('г.')) {
          const dateStr = val.slice(val.indexOf('на') + 2, val.lastIndexOf('г.')).trim();
          const parsed = dayjs(dateStr, 'D MMMM YYYY', 'ru').tz('Europe/Minsk');
          if (parsed.isValid()) {
            menuDate = parsed;
            dateFound = true;
            break;
          }
        }
      }
      if (dateFound) break;
    }

    if (!dateFound) throw new BadRequestException('В Excel-файле не найдена строка с датой меню');

    let tableStartRowIndex = -1;
    for (let i = 0; i < rows.length; i++) {
      if (rows[i].some(cell => String(cell || '').toUpperCase().includes('ККАЛ'))) {
        tableStartRowIndex = i;
        break;
      }
    }

    if (tableStartRowIndex === -1) throw new BadRequestException('В Excel-файле не найдена колонка "ККАЛ"');

    const result: menuPositionDeclaration[] = [];
    let currentCategory = 'Общее';

    for (let i = tableStartRowIndex + 1; i < rows.length; i++) {
      const row = rows[i];
      if (!row || row.length === 0) continue;

      if (row.every(cell => String(cell || '').trim() === '')) {
        console.log('Достигнута пустая строка. Парсинг Excel завершен.');
        break;
      }

      // Базовые значения по ТЗ (когда диета присутствует)
      let dietText = String(row[1] || '').trim();     
      let nameText = String(row[2] || '').trim();     
      let quantityText = String(row[9] || '').trim(); 
      let priceText = String(row[11] || '').trim();   
      let kcalText = String(row[14] || '').trim();    

      // Проверка на строку состава блюда (слэши)
      if (dietText.startsWith('/') || nameText.startsWith('/')) {
        const rawComposition = dietText.startsWith('/') ? dietText : nameText;
        const cleanDescription = rawComposition.replace(/\//g, '').trim();
        
        if (result.length > 0) {
          result[result.length - 1].dishDescription.description = cleanDescription.slice(0, 255);
        }
        continue; 
      }

      // КОРРЕКЦИЯ СДВИГА: Если диеты нет, вся строка съезжает влево на 1 ячейку
      if (!nameText && dietText && !dietText.startsWith('/') && String(row[10] || '').trim() !== '') {
        nameText = String(row[1] || '').trim();
        dietText = ''; // Диеты нет
        quantityText = String(row[8] || '').trim();
        priceText = String(row[10] || '').trim();
        kcalText = String(row[13] || '').trim();
      }

      // Проверка на строку новой категории
      if (nameText && !quantityText && !priceText) {
        currentCategory = nameText;
        continue;
      }

      if (!nameText && !dietText && !quantityText) continue;

      let price = 0;
      let usedPriceStr = priceText;

      // Если в целевой ячейке всё еще пусто, ищем цену по всей строке динамически
      if (!usedPriceStr.trim()) {
        const foundCell = row.find(cell => {
          const c = String(cell || '').trim().toLowerCase();
          return (c.includes('р') && /\d/.test(c)) || (!isNaN(+c.replace(',', '.')) && +c > 0);
        });
        if (foundCell) {
          usedPriceStr = String(foundCell).trim();
        }
      }

      const cleanPriceStr = usedPriceStr.replace(/\s+/g, '').toLowerCase();
      const rublesMatch = cleanPriceStr.match(/(\d+)\s*р/);
      const kopecksMatch = cleanPriceStr.match(/(\d+)\s*к/);

      // ИСПРАВЛЕНО: Добавлены индексы [1] для корректного извлечения чисел из регулярного выражения
      if (rublesMatch || kopecksMatch) {
        const rubles = rublesMatch ? +rublesMatch[1] : 0;
        const kopecks = kopecksMatch ? +kopecksMatch[1] : 0;
        price = (rubles * 100) + kopecks;
      } else {
        const normalizedPrice = cleanPriceStr.replace(',', '.').replace('-', '.');
        if (!isNaN(+normalizedPrice) && +normalizedPrice > 0) {
          price = Math.round(+normalizedPrice * 100);
        }
      }

      if (isNaN(price) || price === 0) {
        const rowJson = JSON.stringify(row);
        throw new BadRequestException(
          `Ошибка обработки цены для: "${nameText || dietText}". Вся строка: ${rowJson}`
        );
      }

      const finalName = dietText ? `${nameText} (Диета ${dietText})` : nameText;

      result.push({
        price,
        discount: 0,
        dishDescription: {
          name: finalName.slice(0, 150),
          description: '', 
          quantity: quantityText.slice(0, 45),
          calorieContent: kcalText ? kcalText.slice(0, 45) : null,
          bestBeforeDate: null,
          carbohydrates: null,
          externalProducer: null,
          proteins: null,
          fats: null,
          categoryName: currentCategory.slice(0, 100),
        },
      });
    }

    return {
      name: null,
      relevantFrom: menuDate.set('h', 8).toDate(),
      expire: menuDate.set('h', 9).set('m', 30).toDate(),
      providingCanteenName: '',
      menuPositions: result,
    };
  }

  parseDocument(document: WordExtractor.Document) {
    const documentLines = document
      .getBody()
      .split(/\n|\t/)
      .map(item => item.trim())
      .filter((item) => !!item);

    if (!documentLines[1]) throw new BadRequestException('Строка даты не найдена');

    const dateStr = documentLines[1].slice(
      documentLines[1].indexOf('на') + 2,
      documentLines[1].lastIndexOf('г.')
    ).trim();

    const menuDate = dayjs(dateStr, 'D MMMM YYYY', 'ru').tz('Europe/Minsk');

    if (!menuDate.isValid()) {
      throw new BadRequestException(`Не удалось распознать дату из строки: "${dateStr}"`);
    }

    const startIndex = documentLines.findIndex((item) => item.toUpperCase().includes('ККАЛ'));
    if (startIndex === -1) {
      throw new BadRequestException('Некорректный формат файла. Не найдена колонка ККАЛ.');
    }

    const result: menuPositionDeclaration[] = [];
    let currentCategory: string = undefined;

    for (let i = startIndex + 1; i < documentLines.length; i++) {
      const currentStr = documentLines[i];

      if (currentStr.startsWith('/')) continue;

      let priceIndex = -1;
      let hasDiet = false;

      const isPotentialDiet = /^[0-9,\s]+$/.test(currentStr);

      if (isPotentialDiet && documentLines[i + 3] && (documentLines[i + 3].includes('р.') || documentLines[i + 3].includes('к.'))) {
        priceIndex = i + 3;
        hasDiet = true;
      }
      else if (documentLines[i + 2] && (documentLines[i + 2].includes('р.') || documentLines[i + 2].includes('к.'))) {
        priceIndex = i + 2;
        hasDiet = false;
      }

      if (priceIndex === -1) {
        currentCategory = currentStr;
        continue;
      }

      const diet = hasDiet ? currentStr : null;
      const name = hasDiet ? documentLines[i + 1] : currentStr;
      const quantity = documentLines[priceIndex - 1];
      const priceStr = documentLines[priceIndex];
      const kcalStr = documentLines[priceIndex + 1];

      const descriptionLine = documentLines[priceIndex + 2];
      const description = descriptionLine && descriptionLine.startsWith('/')
        ? descriptionLine.replace(/\//g, '').trim()
        : '';

      const rubles = +(priceStr.match(/(\d+)р/)?.[1] ?? 0);
      const kopecks = +(priceStr.match(/(\d+)к/)?.[1] ?? 0);
      const price = (rubles * 100) + kopecks;

      if (isNaN(price) || price === 0) {
        throw new BadRequestException(`Ошибка обработки цены блюда ${name}`);
      }

      const finalName = diet ? `${name} (Диета ${diet})` : name;

      result.push({
        price,
        discount: 0,
        dishDescription: {
          name: finalName,
          description,
          quantity,
          calorieContent: kcalStr ? kcalStr.trim() : null,
          bestBeforeDate: null,
          carbohydrates: null,
          externalProducer: null,
          proteins: null,
          fats: null,
          categoryName: currentCategory,
        },
      });

      i = priceIndex + (description ? 2 : 1);
    }

    return {
      name: null,
      relevantFrom: menuDate.set('h', 8).toDate(),
      expire: menuDate.set('h', 9).set('m', 30).toDate(),
      providingCanteenName: '',
      menuPositions: result,
    };
  }

  getParsedExtensions(): string {
    return '.docx,.xlsx,.xls';
  }
}
