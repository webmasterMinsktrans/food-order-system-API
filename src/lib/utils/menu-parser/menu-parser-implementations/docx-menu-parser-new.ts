import {
  menuDeclaration,
  menuPositionDeclaration,
  MenuFileParser,
} from '../menu-parser.interface';
import { BadRequestException, NotAcceptableException } from '@nestjs/common';
import * as WordExtractor from 'word-extractor';
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
    const extractor = new WordExtractor();
    if (buffer instanceof Buffer || typeof buffer === 'string') {
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
    throw new NotAcceptableException('Некорректный формат файла меню');
  }
  parseDocument(document: WordExtractor.Document) {
    // 1. Извлекаем и чистим строки
    const documentLines = document
      .getBody()
      .split(/\n|\t/)
      .map(item => item.trim())
      .filter((item) => !!item);

    // 2. Оригинальный поиск даты по жесткому индексу строки
    if (!documentLines[1]) throw new BadRequestException('Строка даты не найдена');

    const dateStr = documentLines[1].slice(
      documentLines[1].indexOf('на') + 2,
      documentLines[1].lastIndexOf('г.')
    ).trim();

    const menuDate = dayjs(dateStr, 'D MMMM YYYY', 'ru').tz('Europe/Minsk');

    if (!menuDate.isValid()) {
      throw new BadRequestException(`Не удалось распознать дату из строки: "${dateStr}"`);
    }

    // 3. Ищем начало таблицы блюд
    const startIndex = documentLines.findIndex((item) => item.toUpperCase().includes('ККАЛ'));
    if (startIndex === -1) {
      throw new BadRequestException('Некорректный формат файла. Не найдена колонка ККАЛ.');
    }

    const result: menuPositionDeclaration[] = [];
    let currentCategory: string = undefined;

    // 4. Обход позиций меню
    for (let i = startIndex + 1; i < documentLines.length; i++) {
      const currentStr = documentLines[i];

      if (currentStr.startsWith('/')) continue;

      let priceIndex = -1;
      let hasDiet = false;

      // Диета — это ВСЕГДА только цифры (например, "5" или "7,8")
      const isPotentialDiet = /^[0-9,\s]+$/.test(currentStr);

      if (isPotentialDiet && documentLines[i + 3] && (documentLines[i + 3].includes('р.') || documentLines[i + 3].includes('к.'))) {
        priceIndex = i + 3;
        hasDiet = true;
      }
      else if (documentLines[i + 2] && (documentLines[i + 2].includes('р.') || documentLines[i + 2].includes('к.'))) {
        priceIndex = i + 2;
        hasDiet = false;
      }

      // Если структура не распознана как блюдо с диетой или без — это категория
      if (priceIndex === -1) {
        currentCategory = currentStr;
        continue;
      }

      // Разбираем структуру блока
      const diet = hasDiet ? currentStr : null;
      const name = hasDiet ? documentLines[i + 1] : currentStr;
      const quantity = documentLines[priceIndex - 1];
      const priceStr = documentLines[priceIndex];
      const kcalStr = documentLines[priceIndex + 1];

      const descriptionLine = documentLines[priceIndex + 2];
      const description = descriptionLine && descriptionLine.startsWith('/')
        ? descriptionLine.replace(/\//g, '').trim()
        : '';

      // Парсим цену
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
    return '.docx';
  }
}