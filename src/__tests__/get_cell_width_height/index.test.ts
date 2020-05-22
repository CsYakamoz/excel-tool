import { join } from 'path';
import { readFile } from 'xlsx';
import { getCellWidthHeight } from '../../index';

function pathRelativeCurrDir(path: string) {
    return join(__dirname, path);
}

describe('sample.xlsx', () => {
    const workBook = readFile(
        pathRelativeCurrDir('../sample_xlsx/sample.xlsx')
    );
    const workSheet = workBook.Sheets[workBook.SheetNames[0]];

    test('empty', () => {
        expect(getCellWidthHeight('A', 1, workSheet)).toEqual({
            width: 1,
            height: 1,
        });
    });

    test('single', () => {
        expect(getCellWidthHeight('B', 2, workSheet)).toEqual({
            width: 1,
            height: 1,
        });
    });

    test('width is 2', () => {
        expect(getCellWidthHeight('C', 3, workSheet)).toEqual({
            width: 2,
            height: 1,
        });

        expect(getCellWidthHeight('C', 4, workSheet)).toEqual({
            width: 1,
            height: 1,
        });
    });

    test('height is 2', () => {
        expect(getCellWidthHeight('E', 4, workSheet)).toEqual({
            width: 1,
            height: 2,
        });

        expect(getCellWidthHeight('E', 5, workSheet)).toEqual({
            width: 1,
            height: 1,
        });
    });

    test('both width and height are 2', () => {
        expect(getCellWidthHeight('F', 6, workSheet)).toEqual({
            width: 2,
            height: 2,
        });

        expect(getCellWidthHeight('G', 6, workSheet)).toEqual({
            width: 1,
            height: 1,
        });

        expect(getCellWidthHeight('F', 7, workSheet)).toEqual({
            width: 1,
            height: 1,
        });

        expect(getCellWidthHeight('G', 7, workSheet)).toEqual({
            width: 1,
            height: 1,
        });
    });
});

describe('error', () => {
    const workBook = readFile(
        pathRelativeCurrDir('../sample_xlsx/sample.xlsx')
    );
    const workSheet = workBook.Sheets[workBook.SheetNames[1]];

    test('empty workSheet', () => {
        expect(() => {
            getCellWidthHeight('', 0, workSheet);
        }).toThrow('no data in this workSheet');
    });
});
