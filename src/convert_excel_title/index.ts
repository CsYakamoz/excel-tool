const COL_PATTERN = /^(^$|[a-zA-Z]+)$/;
const COL_LIST = Array.from('ABCDEFGHIJKLMNOPQRSTUVWXYZ');

function numberHelper(n: number, str: string): string {
    if (n === 0) {
        return str;
    }

    const nextN = Math.floor((n - 1) / 26);
    const curCol = COL_LIST[(n - 1) % 26];

    return numberHelper(nextN, curCol + str);
}

export function numberToStr(n: number): string {
    if (n < 0) {
        throw new Error('n cannot be negative');
    }

    return numberHelper(n, '');
}

function strHelper(c: string): number {
    return c.charCodeAt(0) - 65;
}

export function strToNumber(str: string): number {
    if (!COL_PATTERN.test(str)) {
        throw new Error(`str should match ${COL_PATTERN}`);
    }

    return str
        .toUpperCase()
        .split('')
        .reduce((acc, curr) => acc * 26 + strHelper(curr) + 1, 0);
}

export default { numberToStr, strToNumber };
