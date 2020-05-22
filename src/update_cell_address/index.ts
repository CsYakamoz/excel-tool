import { numberToStr, strToNumber } from '../convert_excel_title';

export function row(r: number, variation: number): number {
    return r + variation;
}

export function col(c: string, variation: number): string {
    const numCol = strToNumber(c);
    return numberToStr(numCol + variation);
}

export default { row, col };
