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
    content: string | number | boolean | undefined;
    idx: T;
}

interface ScanOption {
    ignoreUndef: boolean;
}

export function scanColBetRowRange(
    col: string,
    rowRange: Range<number>,
    workSheet: WorkSheet,
    option: ScanOption = { ignoreUndef: true }
): ScanResultItem<number>[] {
    const result: ScanResultItem<number>[] = [];

    for (let r = rowRange.begin; r < rowRange.end; ) {
        const value = getCellValue(col, r, workSheet);

        if (value === undefined) {
            if (!option.ignoreUndef) {
                result.push({
                    width: 1,
                    height: 1,
                    content: undefined,
                    idx: r,
                });
            }

            r++;
        } else {
            const { width, height } = getCellWidthHeight(col, r, workSheet);
            result.push({
                width,
                height,

                content: value,
                idx: r,
            });

            r += height;
        }
    }

    return result;
}

export function scanRowBetColRange(
    row: number,
    colRange: Range<string>,
    workSheet: WorkSheet,
    option: ScanOption = { ignoreUndef: true }
): ScanResultItem<string>[] {
    const result: ScanResultItem<string>[] = [];
    const numColRange = {
        begin: strToNumber(colRange.begin),
        end: strToNumber(colRange.end),
    };

    for (let c = numColRange.begin; c < numColRange.end; ) {
        const strCol = numberToStr(c);
        const value = getCellValue(strCol, row, workSheet);

        if (value === undefined) {
            if (!option.ignoreUndef) {
                result.push({
                    width: 1,
                    height: 1,
                    content: undefined,
                    idx: strCol,
                });
            }

            c++;
        } else {
            const { width, height } = getCellWidthHeight(
                strCol,
                row,
                workSheet
            );

            result.push({
                width,
                height,

                content: value,
                idx: strCol,
            });
            c += width;
        }
    }

    return result;
}

export default { scanRowBetColRange, scanColBetRowRange };
