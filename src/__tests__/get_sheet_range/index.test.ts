import { join } from 'path';
import { readFile } from 'xlsx';
import { getSheetRange } from '../../index';

function pathRelativeCurrDir(path: string) {
    return join(__dirname, path);
}

test('file_example_XLSX_10.xlsx', () => {
    const workBook = readFile(
        pathRelativeCurrDir('../sample_xlsx/file_example_XLSX_10.xlsx')
    );
    const workSheet = workBook.Sheets[workBook.SheetNames[0]];

    expect(workSheet['!ref']).toBe('A1:H10');

    expect(getSheetRange(workSheet)).toEqual({
        begin: { row: 1, col: 'A' },
        end: { row: 11, col: 'I' },
    });
});

test('file_example_XLSX_50.xlsx', () => {
    const workBook = readFile(
        pathRelativeCurrDir('../sample_xlsx/file_example_XLSX_50.xlsx')
    );
    const workSheet = workBook.Sheets[workBook.SheetNames[0]];

    expect(workSheet['!ref']).toBe('A1:H51');

    expect(getSheetRange(workSheet)).toEqual({
        begin: { row: 1, col: 'A' },
        end: { row: 52, col: 'I' },
    });
});

test('file_example_XLSX_100.xlsx', () => {
    const workBook = readFile(
        pathRelativeCurrDir('../sample_xlsx/file_example_XLSX_100.xlsx')
    );
    const workSheet = workBook.Sheets[workBook.SheetNames[0]];

    expect(workSheet['!ref']).toBe('A1:H101');

    expect(getSheetRange(workSheet)).toEqual({
        begin: { row: 1, col: 'A' },
        end: { row: 102, col: 'I' },
    });
});

test('file_example_XLSX_1000.xlsx', () => {
    const workBook = readFile(
        pathRelativeCurrDir('../sample_xlsx/file_example_XLSX_1000.xlsx')
    );
    const workSheet = workBook.Sheets[workBook.SheetNames[0]];

    expect(workSheet['!ref']).toBe('A1:H1001');

    expect(getSheetRange(workSheet)).toEqual({
        begin: { row: 1, col: 'A' },
        end: { row: 1002, col: 'I' },
    });
});

test('file_example_XLSX_5000.xlsx', () => {
    const workBook = readFile(
        pathRelativeCurrDir('../sample_xlsx/file_example_XLSX_5000.xlsx')
    );
    const workSheet = workBook.Sheets[workBook.SheetNames[0]];

    expect(workSheet['!ref']).toBe('A1:H5001');

    expect(getSheetRange(workSheet)).toEqual({
        begin: { row: 1, col: 'A' },
        end: { row: 5002, col: 'I' },
    });
});

test('Financial Sample.xlsx', () => {
    const workBook = readFile(
        pathRelativeCurrDir('../sample_xlsx/Financial Sample.xlsx')
    );
    const workSheet = workBook.Sheets[workBook.SheetNames[0]];

    expect(workSheet['!ref']).toBe('A1:P701');

    expect(getSheetRange(workSheet)).toEqual({
        begin: { row: 1, col: 'A' },
        end: { row: 702, col: 'Q' },
    });
});

describe('sample.xlsx', () => {
    const workBook = readFile(
        pathRelativeCurrDir('../sample_xlsx/sample.xlsx')
    );

    test('1st sheet', () => {
        const workSheet = workBook.Sheets[workBook.SheetNames[0]];
        expect(getSheetRange(workSheet)).toEqual({
            begin: { row: 2, col: 'B' },
            end: { row: 8, col: 'H' },
        });
    });

    test('2st sheet', () => {
        const workSheet = workBook.Sheets[workBook.SheetNames[1]];
        expect(() => getSheetRange(workSheet)).toThrow(
            'no data in this workSheet'
        );
    });

    test('3st sheet', () => {
        const workSheet = workBook.Sheets[workBook.SheetNames[2]];
        expect(getSheetRange(workSheet)).toEqual({
            begin: { row: 1, col: 'A' },
            end: { row: 2, col: 'B' },
        });
    });
});
