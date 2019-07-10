import { environment } from '../environments/environment';

/// <reference lib="webworker" />

declare function importScripts(...urls: string[]): void;
declare function postMessage(message: any): any;


// Declare the library object so the script can be compiled without any problem
declare const XLSX: any;

addEventListener('message', ({ data }) => {

    let str = `${data.window.protocol}//${data.window.host}${data.window.path}/scripts/xlsx/xlsx.full.min.js`;

    // production
    if (environment.production) {
        str = str.replace('/index.html', '');
    }

    importScripts(str);

    const hiddenRows = [2, 3, 6, 7, 8, 9, 10, 11, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35];
    const xls = XLSX.read(data.file, { type: 'binary' });

    function clamp_range(range) {
        if (range.e.r >= (1 << 20)) { range.e.r = (1 << 20) - 1; }
        if (range.e.c >= (1 << 14)) { range.e.c = (1 << 14) - 1; }
        return range;
    }

    const crefregex = /(^|[^._A-Z0-9])([$]?)([A-Z]{1,2}|[A-W][A-Z]{2}|X[A-E][A-Z]|XF[A-D])([$]?)([1-9]\d{0,5}|10[0-3]\d{4}|104[0-7]\d{3}|1048[0-4]\d{2}|10485[0-6]\d|104857[0-6])(?![_.\(A-Za-z0-9])/g;

    /*
        deletes `nrows` rows STARTING WITH `startRow`
        - ws         = worksheet object
        - startRow  = starting row (0-indexed) | default 0
        - nrows      = number of rows to delete | default 1
    */

    function delete_rows(ws, startRow, nrows) {
        if (!ws) { throw new Error('operation expects a worksheet'); }
        const dense = Array.isArray(ws);
        if (!nrows) { nrows = 1; }
        if (!startRow) { startRow = 0; }

        /* extract original range */
        const range = XLSX.utils.decode_range(ws['!ref']);
        let R = 0;
        let C = 0;

        const formulaCb = ($0, $1, $2, $3, $4, $5) => {
            let _R = XLSX.utils.decode_row($5);
            const _C = XLSX.utils.decode_col($3);
            if (_R >= startRow) {
                _R -= nrows;
                if (_R < startRow) { return '#REF!'; }
            }
            return $1 + ($2 === '$' ? $2 + $3 : XLSX.utils.encode_col(_C)) + ($4 === '$' ? $4 + $5 : XLSX.utils.encode_row(_R));
        };

        let addr;
        let naddr;
        /* move cells and update formulae */
        if (dense) {
            for (R = startRow + nrows; R <= range.e.r; ++R) {
                if (ws[R]) { ws[R].forEach((cell) => { cell.f = cell.f.replace(crefregex, formulaCb); }); }
                ws[R - nrows] = ws[R];
            }
            ws.length -= nrows;
            for (R = 0; R < startRow; ++R) {
                if (ws[R]) { ws[R].forEach((cell) => { cell.f = cell.f.replace(crefregex, formulaCb); }); }
            }
        } else {
            for (R = startRow + nrows; R <= range.e.r; ++R) {
                for (C = range.s.c; C <= range.e.c; ++C) {
                    addr = XLSX.utils.encode_cell({ r: R, c: C });
                    naddr = XLSX.utils.encode_cell({ r: R - nrows, c: C });
                    if (!ws[addr]) { delete ws[naddr]; continue; }
                    if (ws[addr].f) { ws[addr].f = ws[addr].f.replace(crefregex, formulaCb); }
                    ws[naddr] = ws[addr];
                }
            }
            for (R = range.e.r; R > range.e.r - nrows; --R) {
                for (C = range.s.c; C <= range.e.c; ++C) {
                    addr = XLSX.utils.encode_cell({ r: R, c: C });
                    delete ws[addr];
                }
            }
            for (R = 0; R < startRow; ++R) {
                for (C = range.s.c; C <= range.e.c; ++C) {
                    addr = XLSX.utils.encode_cell({ r: R, c: C });
                    if (ws[addr] && ws[addr].f) { ws[addr].f = ws[addr].f.replace(crefregex, formulaCb); }
                }
            }
        }

        /* write new range */
        range.e.r -= nrows;
        if (range.e.r < range.s.r) { range.e.r = range.s.r; }
        ws['!ref'] = XLSX.utils.encode_range(clamp_range(range));

        /* merge cells */
        if (ws['!merges']) {
            ws['!merges'].forEach((merge, idx) => {
                let mergerange;
                switch (typeof merge) {
                    case 'string': mergerange = XLSX.utils.decode_range(merge); break;
                    case 'object': mergerange = merge; break;
                    default: throw new Error('Unexpected merge ref ' + merge);
                }
                if (mergerange.s.r >= startRow) {
                    mergerange.s.r = Math.max(mergerange.s.r - nrows, startRow);
                }
                if (mergerange.e.r < startRow + nrows) {
                    delete ws['!merges'][idx];
                    return;
                } else if (mergerange.e.r >= startRow) {
                    mergerange.e.r = Math.max(mergerange.e.r - nrows, startRow);
                }
                clamp_range(mergerange);
                ws['!merges'][idx] = mergerange;
            });
        }
        if (ws['!merges']) {
            ws['!merges'] = ws['!merges'].filter((x) => !!x);
        }

        /* rows */
        if (ws['!rows']) {
            ws['!rows'].splice(startRow, nrows);
        }
    }

    const sheet = xls.Sheets[xls.SheetNames[0]];
    sheet['!cols'] = [];

    for (let i = 0; i <= 26; i++) {
        sheet['!cols'][i] = { 'hidden': false, 'width': 40 };
    }

    let scanRange = XLSX.utils.decode_range(sheet['!ref']); // get the range

    for (let R = scanRange.s.r; R <= scanRange.e.r; ++R) {

        // skip first row

        if (R === 0) { continue; }

        /* find the cell object */
        const cellref = XLSX.utils.encode_cell({ c: 1, r: R });

        if (!sheet[cellref]) { continue; } // if cell doesn't exist, move on
        const cell = sheet[cellref];
        if (cell.v !== data.locale) {
            delete_rows(sheet, R, 1);
            R--;
        }
    }


    scanRange = XLSX.utils.decode_range(sheet['!ref']);

    for (let R = scanRange.s.r; R <= scanRange.e.r; ++R) {
        if (R === 0) { continue; }

        let akCell = sheet['AK' + (R + 1)];
        if (!akCell) { akCell = sheet['AK' + (R + 1)] = { t: 's' }; }
        akCell.v = data.val.ak;

        let alCell = sheet['AL' + (R + 1)];
        if (!alCell) { alCell = sheet['AL' + (R + 1)] = { t: 's' }; }
        alCell.v = data.val.al;

        let amCell = sheet['AM' + (R + 1)];
        if (!amCell) { amCell = sheet['AM' + (R + 1)] = { t: 's' }; }
        amCell.v = data.val.am;



    }


    sheet['!ref'] = 'A1:AM' + scanRange.e.r + 1;

    hiddenRows.forEach(i => sheet['!cols'][i] = { hidden: true });



    postMessage(xls);
});




