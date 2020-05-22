import { join } from 'path';
import { readFile } from 'xlsx';
import { getCellValue, excelDate2JsDate } from '../../index';

function pathRelativeCurrDir(path: string) {
    return join(__dirname, path);
}

test('file_example_XLSX_10.xlsx', () => {
    const workBook = readFile(
        pathRelativeCurrDir('../sample_xlsx/file_example_XLSX_10.xlsx')
    );
    const workSheet = workBook.Sheets[workBook.SheetNames[0]];

    expect(getCellValue('H', 10, workSheet)).toBe(6548);
    expect(getCellValue('h', 9, workSheet)).toBe(2456);

    expect(getCellValue('C', 2, workSheet)).toBe('Abril');
    expect(getCellValue('c', 4, workSheet)).toBe('Gent');

    expect(getCellValue('G', 5, workSheet)).toBe('15/10/2017');
    expect(getCellValue('g', 6, workSheet)).toBe('16/08/2016');

    expect(getCellValue('G', 11, workSheet)).toBe(undefined);
    expect(getCellValue('i', 10, workSheet)).toBe(undefined);
});

test('file_example_XLSX_50.xlsx', () => {
    const workBook = readFile(
        pathRelativeCurrDir('../sample_xlsx/file_example_XLSX_50.xlsx')
    );
    const workSheet = workBook.Sheets[workBook.SheetNames[0]];

    expect(getCellValue('H', 20, workSheet)).toBe(9654);
    expect(getCellValue('h', 21, workSheet)).toBe(3569);

    expect(getCellValue('C', 22, workSheet)).toBe('Partain');
    expect(getCellValue('b', 21, workSheet)).toBe('Teresa');

    expect(getCellValue('G', 35, workSheet)).toBe('16/08/2016');
    expect(getCellValue('E', 37, workSheet)).toBe('United States');

    expect(getCellValue('G', 52, workSheet)).toBe(undefined);
    expect(getCellValue('i', 10, workSheet)).toBe(undefined);
});

test('file_example_XLSX_100.xlsx', () => {
    const workBook = readFile(
        pathRelativeCurrDir('../sample_xlsx/file_example_XLSX_100.xlsx')
    );
    const workSheet = workBook.Sheets[workBook.SheetNames[0]];

    expect(getCellValue('H', 72, workSheet)).toBe(2564);
    expect(getCellValue('h', 77, workSheet)).toBe(5555);

    expect(getCellValue('C', 88, workSheet)).toBe('Wachtel');
    expect(getCellValue('b', 80, workSheet)).toBe('Garth');

    expect(getCellValue('G', 100, workSheet)).toBe('15/10/2017');
    expect(getCellValue('E', 94, workSheet)).toBe('France');

    expect(getCellValue('G', 1055, workSheet)).toBe(undefined);
    expect(getCellValue('i', 10, workSheet)).toBe(undefined);
});

test('file_example_XLSX_1000.xlsx', () => {
    const workBook = readFile(
        pathRelativeCurrDir('../sample_xlsx/file_example_XLSX_1000.xlsx')
    );
    const workSheet = workBook.Sheets[workBook.SheetNames[0]];

    expect(getCellValue('H', 951, workSheet)).toBe(6125);
    expect(getCellValue('h', 957, workSheet)).toBe(2554);

    expect(getCellValue('C', 969, workSheet)).toBe('Perrine');
    expect(getCellValue('b', 975, workSheet)).toBe('Libbie');

    expect(getCellValue('G', 990, workSheet)).toBe('21/05/2015');
    expect(getCellValue('E', 982, workSheet)).toBe('Great Britain');

    expect(getCellValue('G', 1055, workSheet)).toBe(undefined);
    expect(getCellValue('i', 10, workSheet)).toBe(undefined);
});

test('file_example_XLSX_5000.xlsx', () => {
    const workBook = readFile(
        pathRelativeCurrDir('../sample_xlsx/file_example_XLSX_5000.xlsx')
    );
    const workSheet = workBook.Sheets[workBook.SheetNames[0]];

    expect(getCellValue('D', 4770, workSheet)).toBe('Female');
    expect(getCellValue('h', 4780, workSheet)).toBe(3256);

    expect(getCellValue('C', 4921, workSheet)).toBe('Strawn');
    expect(getCellValue('b', 4939, workSheet)).toBe('Sau');

    expect(getCellValue('G', 4987, workSheet)).toBe('15/10/2017');
    expect(getCellValue('E', 4975, workSheet)).toBe('France');

    expect(getCellValue('G', 5055, workSheet)).toBe(undefined);
    expect(getCellValue('i', 10, workSheet)).toBe(undefined);
});

test('Financial Sample.xlsx', () => {
    const workBook = readFile(
        pathRelativeCurrDir('../sample_xlsx/Financial Sample.xlsx')
    );
    const workSheet = workBook.Sheets[workBook.SheetNames[0]];

    expect(getCellValue('a', 157, workSheet)).toBe('Small Business');
    expect(getCellValue('b', 167, workSheet)).toBe('Canada');

    expect(getCellValue('C', 226, workSheet)).toBe('Carretera');
    expect(getCellValue('d', 310, workSheet)).toBe('Medium');

    expect(getCellValue('E', 369, workSheet)).toBe(1562);
    expect(getCellValue('f', 599, workSheet)).toBe(10);

    expect(getCellValue('G', 611, workSheet)).toBe(7);
    expect(getCellValue('h', 662, workSheet)).toBe(29700);

    expect(getCellValue('I', 684, workSheet)).toBe(900);
    expect(getCellValue('j', 699, workSheet)).toBe(8139.6);

    expect(getCellValue('K', 410, workSheet)).toBe(6963);
    expect(getCellValue('l', 342, workSheet)).toBe(22546.08);

    expect(typeof getCellValue('M', 130, workSheet)).toBe('number');
    expect(
        excelDate2JsDate(getCellValue('M', 130, workSheet) as number)
    ).toEqual(new Date('2014-12-01'));
    expect(getCellValue('n', 33, workSheet)).toBe(1);

    expect(getCellValue('O', 5, workSheet)).toBe('June');
    expect(getCellValue('p', 1, workSheet)).toBe('Year');
});
