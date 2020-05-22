import { convertExcelTitle } from '../../index';

function getRandomIntInclusive(min: number, max: number) {
    min = Math.ceil(min);
    max = Math.floor(max);
    return Math.floor(Math.random() * (max - min + 1)) + min; //含最大值，含最小值
}

describe('str to number', () => {
    test('"" -> 0', () => {
        expect(convertExcelTitle.strToNumber('')).toBe(0);
    });

    test('"A" -> 1', () => {
        expect(convertExcelTitle.strToNumber('A')).toBe(1);
        expect(convertExcelTitle.strToNumber('a')).toBe(1);
    });

    test('"AA" -> 27', () => {
        expect(convertExcelTitle.strToNumber('AA')).toBe(27);
        expect(convertExcelTitle.strToNumber('aa')).toBe(27);
    });

    test('"CV" -> 100', () => {
        expect(convertExcelTitle.strToNumber('CV')).toBe(100);
        expect(convertExcelTitle.strToNumber('cv')).toBe(100);
    });

    test('"AFV" -> 854', () => {
        expect(convertExcelTitle.strToNumber('AFV')).toBe(854);
        expect(convertExcelTitle.strToNumber('afv')).toBe(854);
    });

    test('"FWD" -> 4658', () => {
        expect(convertExcelTitle.strToNumber('FWD')).toBe(4658);
        expect(convertExcelTitle.strToNumber('fwd')).toBe(4658);
    });

    test('"UEA" -> 14327', () => {
        expect(convertExcelTitle.strToNumber('UEA')).toBe(14327);
        expect(convertExcelTitle.strToNumber('uea')).toBe(14327);
    });

    test('"XFD" -> 16384', () => {
        expect(convertExcelTitle.strToNumber('XFD')).toBe(16384);
        expect(convertExcelTitle.strToNumber('xfd')).toBe(16384);
    });
});

describe('number to str', () => {
    test('0 -> ""', () => {
        expect(convertExcelTitle.numberToStr(0)).toBe('');
    });

    test('1 -> "A"', () => {
        expect(convertExcelTitle.numberToStr(1)).toBe('A');
    });

    test('27 -> "AA"', () => {
        expect(convertExcelTitle.numberToStr(27)).toBe('AA');
    });

    test('100 -> "CV"', () => {
        expect(convertExcelTitle.numberToStr(100)).toBe('CV');
    });

    test('854 -> "AFV"', () => {
        expect(convertExcelTitle.numberToStr(854)).toBe('AFV');
    });

    test('4658 -> "FWD"', () => {
        expect(convertExcelTitle.numberToStr(4658)).toBe('FWD');
    });

    test('14327 -> "UEA"', () => {
        expect(convertExcelTitle.numberToStr(14327)).toBe('UEA');
    });

    test('16384 -> "XFD"', () => {
        expect(convertExcelTitle.numberToStr(16384)).toBe('XFD');
    });
});

describe('rollback', () => {
    test('number -> str -> number', () => {
        for (let i = 0; i < 10; i++) {
            const number = getRandomIntInclusive(0, Number.MAX_SAFE_INTEGER);
            expect(
                convertExcelTitle.strToNumber(
                    convertExcelTitle.numberToStr(number)
                )
            ).toBe(number);
        }
    });

    test('str -> number -> str', () => {
        function getRandomStr() {
            const result = [];
            const length = getRandomIntInclusive(1, 4);
            const COL_LIST = Array.from('ABCDEFGHIJKLMNOPQRSTUVWXYZ');

            while (result.length < length) {
                result.push(
                    COL_LIST[getRandomIntInclusive(0, COL_LIST.length - 1)]
                );
            }

            return result.join('');
        }
        for (let i = 0; i < 10; i++) {
            const str = getRandomStr();
            expect(
                convertExcelTitle.numberToStr(
                    convertExcelTitle.strToNumber(str)
                )
            ).toBe(str);
        }
    });
});

describe('throw error', () => {
    test('number to str: n cannot be negative', () => {
        const msg = 'n cannot be negative';
        expect(() =>
            convertExcelTitle.numberToStr(
                getRandomIntInclusive(Number.MIN_SAFE_INTEGER, -1)
            )
        ).toThrow(msg);

        expect(() =>
            convertExcelTitle.numberToStr(
                getRandomIntInclusive(Number.MIN_SAFE_INTEGER, -1)
            )
        ).toThrow(msg);

        expect(() =>
            convertExcelTitle.numberToStr(
                getRandomIntInclusive(Number.MIN_SAFE_INTEGER, -1)
            )
        ).toThrow(msg);
    });

    test('str to number: str should match /^(^$|[a-zA-Z]+)$/', () => {
        const msg = `str should match ${/^(^$|[a-zA-Z]+)$/}`;

        expect(() => convertExcelTitle.strToNumber('0123456789')).toThrow(msg);

        expect(() => convertExcelTitle.strToNumber('~,./;[]{}()-=_+')).toThrow(
            msg
        );

        expect(() => convertExcelTitle.strToNumber(' ')).toThrow(msg);
    });
});
