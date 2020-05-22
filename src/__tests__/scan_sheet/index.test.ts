import { join } from 'path';
import { readFile } from 'xlsx';
import { scanSheet, excelDate2JsDate } from '../../index';

function pathRelativeCurrDir(path: string) {
    return join(__dirname, path);
}

describe('file_example_XLSX_10.xlsx', () => {
    const workBook = readFile(
        pathRelativeCurrDir('../sample_xlsx/file_example_XLSX_10.xlsx')
    );
    const workSheet = workBook.Sheets[workBook.SheetNames[0]];

    test('B col', () => {
        expect(
            scanSheet
                .scanColBetRowRange('B', { begin: 2, end: 5 }, workSheet)
                .map((item) => item.content)
        ).toEqual(['Dulce', 'Mara', 'Philip']);
    });
});

describe('file_example_XLSX_50.xlsx', () => {
    const workBook = readFile(
        pathRelativeCurrDir('../sample_xlsx/file_example_XLSX_50.xlsx')
    );
    const workSheet = workBook.Sheets[workBook.SheetNames[0]];

    test('C col', () => {
        expect(
            scanSheet
                .scanColBetRowRange('c', { begin: 33, end: 38 }, workSheet)
                .map((item) => item.content)
        ).toEqual(['Becker', 'Grindle', 'Claywell', 'Borger', 'Hacker']);
    });
});

describe('file_example_XLSX_100.xlsx', () => {
    const workBook = readFile(
        pathRelativeCurrDir('../sample_xlsx/file_example_XLSX_100.xlsx')
    );
    const workSheet = workBook.Sheets[workBook.SheetNames[0]];

    test('D col', () => {
        expect(
            scanSheet
                .scanColBetRowRange('D', { begin: 77, end: 80 }, workSheet)
                .map((item) => item.content)
        ).toEqual(['Female', 'Female', 'Female']);
    });
});

describe('file_example_XLSX_1000.xlsx', () => {
    const workBook = readFile(
        pathRelativeCurrDir('../sample_xlsx/file_example_XLSX_1000.xlsx')
    );
    const workSheet = workBook.Sheets[workBook.SheetNames[0]];

    test('E col', () => {
        expect(
            scanSheet
                .scanColBetRowRange('e', { begin: 977, end: 980 }, workSheet)
                .map((item) => item.content)
        ).toEqual(['Great Britain', 'United States', 'France']);
    });
});

describe('file_example_XLSX_1000.xlsx', () => {
    const workBook = readFile(
        pathRelativeCurrDir('../sample_xlsx/file_example_XLSX_5000.xlsx')
    );
    const workSheet = workBook.Sheets[workBook.SheetNames[0]];

    test('F col', () => {
        expect(
            scanSheet
                .scanColBetRowRange('F', { begin: 2411, end: 2415 }, workSheet)
                .map((item) => item.content)
        ).toEqual([28, 39, 38, 32]);
    });

    test('G col', () => {
        expect(
            scanSheet
                .scanColBetRowRange('g', { begin: 3080, end: 3084 }, workSheet)
                .map((item) => item.content)
        ).toEqual(['16/08/2016', '21/05/2015', '15/10/2017', '16/08/2016']);
    });

    test('G col', () => {
        expect(
            scanSheet
                .scanColBetRowRange('H', { begin: 4621, end: 4623 }, workSheet)
                .map((item) => item.content)
        ).toEqual([3569, 2564]);
    });
});

describe('Financial Sample.xlsx', () => {
    const workBook = readFile(
        pathRelativeCurrDir('../sample_xlsx/Financial Sample.xlsx')
    );
    const workSheet = workBook.Sheets[workBook.SheetNames[0]];

    test('A col', () => {
        expect(
            scanSheet
                .scanColBetRowRange('A', { begin: 1, end: 5 }, workSheet)
                .map((item) => item.content)
        ).toEqual(['Segment', 'Government', 'Government', 'Midmarket']);
    });

    test('B col', () => {
        expect(
            scanSheet
                .scanColBetRowRange('B', { begin: 1, end: 5 }, workSheet)
                .map((item) => item.content)
        ).toEqual(['Country', 'Canada', 'Germany', 'France']);
    });

    test('C col', () => {
        expect(
            scanSheet
                .scanColBetRowRange('C', { begin: 1, end: 5 }, workSheet)
                .map((item) => item.content)
        ).toEqual(['Product', 'Carretera', 'Carretera', 'Carretera']);
    });

    test('D col', () => {
        expect(
            scanSheet
                .scanColBetRowRange('D', { begin: 1, end: 5 }, workSheet)
                .map((item) => item.content)
        ).toEqual(['Discount Band', 'None', 'None', 'None']);
    });

    test('E col', () => {
        expect(
            scanSheet
                .scanColBetRowRange('E', { begin: 1, end: 5 }, workSheet)
                .map((item) => item.content)
        ).toEqual(['Units Sold', 1618.5, 1321, 2178]);
    });

    test('F col', () => {
        expect(
            scanSheet
                .scanColBetRowRange('F', { begin: 1, end: 5 }, workSheet)
                .map((item) => item.content)
        ).toEqual(['Manufacturing Price', 3, 3, 3]);
    });

    test('G col', () => {
        expect(
            scanSheet
                .scanColBetRowRange('G', { begin: 1, end: 5 }, workSheet)
                .map((item) => item.content)
        ).toEqual(['Sale Price', 20, 20, 15]);
    });

    test('H col', () => {
        expect(
            scanSheet
                .scanColBetRowRange('H', { begin: 1, end: 5 }, workSheet)
                .map((item) => item.content)
        ).toEqual(['Gross Sales', 32370, 26420, 32670]);
    });

    test('I col', () => {
        expect(
            scanSheet
                .scanColBetRowRange('I', { begin: 1, end: 5 }, workSheet)
                .map((item) => item.content)
        ).toEqual(['Discounts', 0, 0, 0]);
    });

    test('J col', () => {
        expect(
            scanSheet
                .scanColBetRowRange('J', { begin: 1, end: 5 }, workSheet)
                .map((item) => item.content)
        ).toEqual([' Sales', 32370, 26420, 32670]);
    });

    test('K col', () => {
        expect(
            scanSheet
                .scanColBetRowRange('K', { begin: 1, end: 5 }, workSheet)
                .map((item) => item.content)
        ).toEqual(['COGS', 16185, 13210, 21780]);
    });

    test('L col', () => {
        expect(
            scanSheet
                .scanColBetRowRange('L', { begin: 1, end: 5 }, workSheet)
                .map((item) => item.content)
        ).toEqual(['Profit', 16185, 13210, 10890]);
    });

    test('M col', () => {
        expect(
            scanSheet
                .scanColBetRowRange('M', { begin: 1, end: 5 }, workSheet)
                .map((item) =>
                    typeof item.content === 'string'
                        ? item.content
                        : excelDate2JsDate(item.content as number)
                )
        ).toEqual([
            'Date',
            new Date('2014-01-01'),
            new Date('2014-01-01'),
            new Date('2014-06-01'),
        ]);
    });

    test('N col', () => {
        expect(
            scanSheet
                .scanColBetRowRange('N', { begin: 1, end: 5 }, workSheet)
                .map((item) => item.content)
        ).toEqual(['Month Number', 1, 1, 6]);
    });

    test('O col', () => {
        expect(
            scanSheet
                .scanColBetRowRange('O', { begin: 1, end: 5 }, workSheet)
                .map((item) => item.content)
        ).toEqual(['Month Name', 'January', 'January', 'June']);
    });

    test('P col', () => {
        expect(
            scanSheet
                .scanColBetRowRange('P', { begin: 1, end: 5 }, workSheet)
                .map((item) => item.content)
        ).toEqual(['Year', '2014', '2014', '2014']);
    });
});
