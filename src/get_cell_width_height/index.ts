/* eslint-disable-next-line */
import { Range, WorkSheet } from 'xlsx';
import { strToNumber } from '../convert_excel_title';

interface CellWidthHeight {
    width: number;
    height: number;
}

export default function cellWidthHeight(
    col: string,
    row: number,
    workSheet: WorkSheet
): CellWidthHeight {
    if (workSheet['!ref'] === undefined) {
        throw new Error('no data in this workSheet');
    }
    if (workSheet['!merges'] === undefined) {
        return defaultValue();
    }

    const numCol = strToNumber(col);
    const range = workSheet['!merges'].find(
        (m) => m.s.r === row - 1 && m.s.c === numCol - 1
    );

    if (range === undefined) {
        return defaultValue();
    }

    return {
        width: calcWidth(range),
        height: calcHeight(range),
    };
}

function defaultValue() {
    return { width: 1, height: 1 };
}

function calcHeight(range: Range) {
    return range.e.r - range.s.r + 1;
}

function calcWidth(range: Range) {
    return range.e.c - range.s.c + 1;
}
