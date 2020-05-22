/* eslint-disable-next-line */
import { WorkSheet } from 'xlsx';
import { numberToStr, strToNumber } from '../convert_excel_title';

const singlePattern = /^([A-Z]+)(\d+)$/;
const rangePattern = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/;

interface SheetRange {
    begin: { row: number; col: string };
    end: { row: number; col: string };
}

export default function sheetRange(workSheet: WorkSheet): SheetRange {
    const result = getResult(workSheet);
    if (result === undefined) {
        throw new Error('no data in this workSheet');
    }

    return {
        begin: helper(result.begin.col, result.begin.row, false),
        end: helper(result.end.col, result.end.row, true),
    };
}

function getResult(workSheet: WorkSheet) {
    if (workSheet['!ref'] === undefined) {
        return undefined;
    }

    const ref = workSheet['!ref'];
    if (singlePattern.test(ref)) {
        const match = ref.match(singlePattern) as RegExpMatchArray;
        const address = { col: match[1], row: match[2] };

        return { begin: address, end: address };
    }

    if (rangePattern.test(ref)) {
        const match = ref.match(rangePattern) as RegExpMatchArray;

        return {
            begin: { col: match[1], row: match[2] },
            end: { col: match[3], row: match[4] },
        };
    }

    throw new Error(`unknown workSheet["!ref"]: ${ref}`);
}

function helper(col: string, row: string, isEnd: boolean) {
    const inc = isEnd ? 1 : 0;
    return {
        row: Number(row) + inc,
        col: numberToStr(strToNumber(col) + inc),
    };
}
