import { Injectable } from '@angular/core';
import * as XLSX from 'xlsx';

@Injectable({
    providedIn: 'root'
})
export class XlsService {

    inputWindow = {
        host: window.location.host,
        path: window.location.pathname,
        protocol: window.location.protocol
    };

    constructor() {}

    loadXLS(files: FileList, values: {}, loc: string) {
        Array.from(files).forEach(file => {
            const worker = new Worker('../excel.worker.ts', {
                type: 'module'
            });
            const reader: FileReader = new FileReader();
            reader.readAsArrayBuffer(file);
            reader.onloadend = (e: any) => {
                worker.onmessage = ({ data }) => {
                    XLSX.writeFile(data, file.name);
                };
                worker.postMessage({ file: new Uint8Array(e.target.result), locale: loc, val: values, window: this.inputWindow });
            };
        });
    }
}
