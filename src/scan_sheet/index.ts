/* eslint-disable-next-line */
import { WorkSheet } from 'xlsx';
import getCellValue from '../get_cell_value';
import getCellWidthHeight from '../get_cell_width_height';
import { numberToStr, strToNumber } from '../convert_excel_title';

interface Range<T> {
    begin: T;
    end: T;
}

interface ScanResultItem<T> {
    width: number;
    height: number;
    content: string | number | boolean;
    idx: T;
}

export function scanColBetRowRange(
    col: string,
    rowRange: Range<number>,
    workSheet: WorkSheet
): ScanResultItem<number>[] {
    const result = [];

    for (let r = rowRange.begin; r < rowRange.end; r++) {
        const value = getCellValue(col, r, workSheet);

        if (value === undefined) {
            continue;
        }

        const { width, height } = getCellWidthHeight(col, r, workSheet);

        result.push({
            width,
            height,

            content: value,
            idx: r,
        });
    }

    return result;
}

export function scanRowBetColRange(
    row: number,
    colRange: Range<string>,
    workSheet: WorkSheet
): ScanResultItem<string>[] {
    const result = [];
    const numColRange = {
        begin: strToNumber(colRange.begin),
        end: strToNumber(colRange.end),
    };

    for (let c = numColRange.begin; c < numColRange.end; c++) {
        const strCol = numberToStr(c);
        const value = getCellValue(strCol, row, workSheet);

        if (value === undefined) {
            continue;
        }

        const { width, height } = getCellWidthHeight(strCol, row, workSheet);

        result.push({
            width,
            height,

            content: value,
            idx: strCol,
        });
    }

    return result;
}

export default { scanRowBetColRange, scanColBetRowRange };
