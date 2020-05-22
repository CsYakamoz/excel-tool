import { updateCellAddress } from '../../index';

describe('row update', () => {
    test('cell address row update: 0 + 1 -> 1', () => {
        expect(updateCellAddress.row(0, 1)).toBe(1);
    });

    test('cell address row update: 1 + 26 -> 27', () => {
        expect(updateCellAddress.row(1, 26)).toBe(27);
    });

    test('cell address row update: 27 + 73 -> 100', () => {
        expect(updateCellAddress.row(27, 73)).toBe(100);
    });

    test('cell address row update: 100 + 754 -> 854', () => {
        expect(updateCellAddress.row(100, 754)).toBe(854);
    });

    test('cell address row update: 854 + 3804 -> 4658', () => {
        expect(updateCellAddress.row(854, 3804)).toBe(4658);
    });

    test('cell address row update: 4658 + 9669 -> 14327', () => {
        expect(updateCellAddress.row(4658, 9669)).toBe(14327);
    });

    test('cell address row update: 14327 + 2057 -> 16384', () => {
        expect(updateCellAddress.row(14327, 2057)).toBe(16384);
    });
});

describe('col update', () => {
    test('error', () => {
        expect(() => updateCellAddress.col('123123', 1)).toThrow(
            Error('str should match /^(^$|[a-zA-Z]+)$/')
        );

        expect(() => updateCellAddress.col('1a', 1)).toThrow(
            Error('str should match /^(^$|[a-zA-Z]+)$/')
        );

        expect(() => updateCellAddress.col('1A', 1)).toThrow(
            Error('str should match /^(^$|[a-zA-Z]+)$/')
        );

        expect(() => updateCellAddress.col('a1', 1)).toThrow(
            Error('str should match /^(^$|[a-zA-Z]+)$/')
        );

        expect(() => updateCellAddress.col('A1', 1)).toThrow(
            Error('str should match /^(^$|[a-zA-Z]+)$/')
        );
    });

    test('cell address col update: "" + 1 -> "A"', () => {
        expect(updateCellAddress.col('', 1)).toBe('A');
    });

    test('cell address col update: "A" + 26 -> "AA"', () => {
        expect(updateCellAddress.col('A', 26)).toBe('AA');
        expect(updateCellAddress.col('a', 26)).toBe('AA');
    });

    test('cell address col update: "AA" + 73 -> "CV"', () => {
        expect(updateCellAddress.col('AA', 73)).toBe('CV');
        expect(updateCellAddress.col('aa', 73)).toBe('CV');
    });

    test('cell address col update: "CV" + 754 -> "AFV"', () => {
        expect(updateCellAddress.col('CV', 754)).toBe('AFV');
        expect(updateCellAddress.col('cv', 754)).toBe('AFV');
    });

    test('cell address col update: "AFV" + 3804 -> "FWD"', () => {
        expect(updateCellAddress.col('AFV', 3804)).toBe('FWD');
        expect(updateCellAddress.col('afv', 3804)).toBe('FWD');
    });

    test('cell address col update: "FWD" + 9669 -> "UEA"', () => {
        expect(updateCellAddress.col('FWD', 9669)).toBe('UEA');
        expect(updateCellAddress.col('fwd', 9669)).toBe('UEA');
    });

    test('cell address col update: "UEA" + 2057 -> "XFD"', () => {
        expect(updateCellAddress.col('UEA', 2057)).toBe('XFD');
        expect(updateCellAddress.col('uea', 2057)).toBe('XFD');
    });

    test('cell address col update: "XFD" - 2057 -> "UEA"', () => {
        expect(updateCellAddress.col('XFD', -2057)).toBe('UEA');
        expect(updateCellAddress.col('xfd', -2057)).toBe('UEA');
    });

    test('cell address col update: "UEA" - 9669 -> "FWD"', () => {
        expect(updateCellAddress.col('UEA', -9669)).toBe('FWD');
        expect(updateCellAddress.col('uea', -9669)).toBe('FWD');
    });

    test('cell address col update: "FWD" - 3804 -> "AFV"', () => {
        expect(updateCellAddress.col('FWD', -3804)).toBe('AFV');
        expect(updateCellAddress.col('fwd', -3804)).toBe('AFV');
    });

    test('cell address col update: "AFV" - 754 -> "CV"', () => {
        expect(updateCellAddress.col('AFV', -754)).toBe('CV');
        expect(updateCellAddress.col('afv', -754)).toBe('CV');
    });

    test('cell address col update: "CV" - 73 -> "AA"', () => {
        expect(updateCellAddress.col('CV', -73)).toBe('AA');
        expect(updateCellAddress.col('cv', -73)).toBe('AA');
    });

    test('cell address col update: "AA" - 26 -> "A"', () => {
        expect(updateCellAddress.col('AA', -26)).toBe('A');
        expect(updateCellAddress.col('aa', -26)).toBe('A');
    });

    test('cell address col update: "A" - 1 -> ""', () => {
        expect(updateCellAddress.col('A', -1)).toBe('');
    });
});
