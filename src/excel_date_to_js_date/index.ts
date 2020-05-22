/**
 * keyword: excel leap year bug
 * reference: https://gist.github.com/christopherscott/2782634
 */
export default function excelDate2JsDate(excelDate: number): Date {
    return new Date((excelDate - (25567 + 2)) * 86400 * 1000);
}
