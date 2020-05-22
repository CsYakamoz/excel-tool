/* eslint-disable-next-line */
import { CellObject, WorkSheet } from 'xlsx';

export default function cellValue(
    col: string,
    row: number,
    workSheet: WorkSheet
): string | number | boolean | undefined {
    const point = col.toUpperCase() + row.toString();

    if (workSheet[point] === undefined) {
        return undefined;
    }

    const cell = workSheet[point] as CellObject;

    return cell.v as string | number | boolean;
}
