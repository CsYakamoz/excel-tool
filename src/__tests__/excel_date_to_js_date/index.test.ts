import { join } from 'path';
import { readFile } from 'xlsx';
import { getCellValue, excelDate2JsDate } from '../../index';

function pathRelativeCurrDir(path: string) {
    return join(__dirname, path);
}

describe('Financial Sample.xlsx', () => {
    const workBook = readFile(
        pathRelativeCurrDir('../sample_xlsx/Financial Sample.xlsx')
    );
    const workSheet = workBook.Sheets[workBook.SheetNames[0]];

    test('M7', () => {
        const value = getCellValue('M', 7, workSheet) as number;
        expect(typeof value).toBe('number');
        expect(excelDate2JsDate(value)).toEqual(new Date('2014-12-01'));
    });

    test('M139', () => {
        const value = getCellValue('M', 139, workSheet) as number;
        expect(typeof value).toBe('number');
        expect(excelDate2JsDate(value)).toEqual(new Date('2014-03-01'));
    });

    test('M303', () => {
        const value = getCellValue('M', 303, workSheet) as number;
        expect(typeof value).toBe('number');
        expect(excelDate2JsDate(value)).toEqual(new Date('2013-10-01'));
    });

    test('M566', () => {
        const value = getCellValue('M', 566, workSheet) as number;
        expect(typeof value).toBe('number');
        expect(excelDate2JsDate(value)).toEqual(new Date('2014-11-01'));
    });

    test('M701', () => {
        const value = getCellValue('M', 701, workSheet) as number;
        expect(typeof value).toBe('number');
        expect(excelDate2JsDate(value)).toEqual(new Date('2014-05-01'));
    });
});

