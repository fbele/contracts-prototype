import { Injectable } from '@angular/core';
import * as ExcelJS from 'exceljs/dist/exceljs.min.js';
import * as ExcelProper from 'exceljs';
import * as FileSaver from 'file-saver';
/* import * as PDFJS from 'pdfjs-dist/build/pdf.min';
import * as PDFJSWorker from 'pdfjs-dist/build/pdf.worker.min'; */

@Injectable({
    providedIn: 'root'
})

export class ExportExcelService {
    blobType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8';

    constructor() { }

    exportExcelFile() {
        // workbook
        const workbook = new ExcelJS.Workbook();
        workbook.creator = 'Bastian Weick';
        workbook.lastModifiedBy = 'Eva Groselj-Varvodic';
        workbook.created = new Date(2015, 7, 22, 17, 16);
        workbook.modified = new Date();
        // workbook.lastPrinted = new Date(2016, 9, 27);

        // worksheet

        const worksheet = workbook.addWorksheet('Sheet', {
            properties: {
                defaultRowHeight: 15
            },
            pageSetup: {
                paperSize : 9,
                orientation : 'portrait',
                fitToPage : true,
                fitToWidth : 1,
                fitToHeight : 0
            },
            views: [{
                /* x: 0,
                y: 0,
                width: 10000,
                height: 10000,
                firstSheet: 0,
                activeTab: 1,
                visibility: 'visible', */
                zoomScale : 85,
                zoomScaleNormal : 100
            }]
        });

        console.log(worksheet);

        worksheet.columns = [
            { key: 'A', width: 3.71 },
            { key: 'B', width: 5.42 },
            { key: 'C', width: 74.71 },
            { key: 'D', width: 22.71 },
            { key: 'E', width: 16.57 },
            { key: 'F', width: 12.71 },
            { key: 'G', width: 1 },
            { key: 'H', width: 22.71 },
            { key: 'I', width: 1 },
            { key: 'J', width: 50.71 }
        ];

        /* worksheet.addRow(['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h']);
        worksheet.addRow(['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h']);
        worksheet.addRow(['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h']);
        worksheet.getRow(1).height = 300; */

        worksheet.addRow().height = 36;
        worksheet.addRow().height = 48;
        worksheet.addRow().height = 18.75;
        worksheet.addRow().height = 18.75;
        worksheet.addRow().height = 19.50;
        worksheet.addRow().height = 16.50;
        worksheet.addRow().height = 15.75;
        worksheet.addRow().height = 29.25;
        worksheet.addRow().height = 15.75;

        worksheet.mergeCells('A1:J1');
        worksheet.mergeCells('A2:J2');
        worksheet.mergeCells('A3:A5');
        worksheet.mergeCells('B3:B5');
        worksheet.mergeCells('D3:H3');
        worksheet.mergeCells('D4:H4');
        worksheet.mergeCells('D5:H5');
        worksheet.mergeCells('I3:J5');
        worksheet.mergeCells('A6:J6');
        worksheet.mergeCells('A7:J7');
        worksheet.mergeCells('A8:J8');
        worksheet.mergeCells('A9:J9');

        const A1 = worksheet.getCell('A1');
        A1.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFFF6600' }
        };
        A1.font = {
            name: 'Calibri',
            size: 36,
            bold: true
        };
        A1.alignment = {
            vertical: 'middle',
            horizontal: 'center'
        };
        A1.value = 'Customer Order';

        const A2 = worksheet.getCell('A2');
        A2.fill = A1.fill;
        A2.font = A1.font;
        A2.alignment = A1.alignment;
        A2.value = 'Germany';

        // fill A3 with yellow dark trellis and blue behind
        worksheet.getCell('A3').fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFFF6600' }
        };

        const B3 = worksheet.getCell('B3');
        B3.border = {
            top: { style: 'medium' },
            left: { style: 'medium' },
            bottom: { style: 'medium' }
        };
        B3.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFC0C0C0' }
        };

        const C3 = worksheet.getCell('C3');
        C3.border = {
            top: { style: 'medium' }
        };
        C3.font = {
            name: 'Calibri',
            size: 14,
            bold: true
        };
        C3.alignment = {
            vertical: 'bottom'
        };
        C3.value = 'Lieferantennr.: Supplier Number:';

        const C4 = worksheet.getCell('C4');
        C4.font = C3.font;
        C4.alignment = C3.alignment;
        C4.value = 'Lieferantenname: Supplier Name:';

        const C5 = worksheet.getCell('C5');
        C5.border = {
            bottom: { style: 'medium' }
        };
        C5.font = C3.font;
        C5.alignment = C3.alignment;
        C5.value = 'Gültig ab: Valid from:';

        const D3 = worksheet.getCell('D3');
        D3.border = {
            top: { style: 'medium' }
        };
        D3.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFFFFF99' }
        };
        D3.alignment = {
            vertical: 'middle',
            horizontal: 'left'
        };
        D3.font = C3.font;

        const D4 = worksheet.getCell('D4');
        D4.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFEEECE1' }
        };
        D4.font = C3.font;
        D4.alignment = D3.alignment;
        D4.value = 'kein Wert vorhanden';

        const D5 = worksheet.getCell('D5');
        D5.border = {
            bottom: { style: 'medium' }
        };
        D5.fill = D4.fill;
        D5.font = C3.font;
        D5.alignment = D3.alignment;
        D5.value = new Date(2017, 1, 1);

        worksheet.getCell('I3').border = {
            top: { style: 'medium' },
            right: { style: 'medium' },
            bottom: { style: 'medium' }
        };

        const A6 = worksheet.getCell('A6');
        A6.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFFF6600' }
        };
        A6.border = {
            bottom: { style: 'medium' }
        };

        const A7 = worksheet.getCell('A7');
        A7.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFFFFFFF' }
        };
        A7.border = {
            bottom: { style: 'medium' }
        };

        const A8 = worksheet.getCell('A8');
        A8.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FF808080' }
        };
        A8.border = {
            bottom: { style: 'medium' }
        };

        const A9 = worksheet.getCell('A9');
        A9.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFFFFFFF' }
        };
        A9.border = {
            bottom: { style: 'medium' }
        };

        /* console.log(PDFJS);
        console.log(PDFJSWorker);
        PDFJS.GlobalWorkerOptions.workerSrc = './assets/src/pdf.worker.js'; */

        // return workbook.xlsx.writeBuffer();
        workbook.xlsx.writeBuffer().then((data: Uint8Array) => {
            const blob = new Blob([data], { type: this.blobType });

            const fileURL = URL.createObjectURL(blob);
            const w = window.open(fileURL);
            w.print();
            // FileSaver.saveAs(blob, 'test.xlsx');
        });

        /* workbook.xlsx.writeBuffer().then((data: Uint8Array) => {
            // It is necessary to create a new blob object with mime-type explicitly set
            // otherwise only Chrome works like it should
            console.log(data);
            const newBlob = new Blob([data], {type: 'application/pdf'})

            // IE doesn't allow using a blob object directly as link href
            // instead it is necessary to use msSaveOrOpenBlob
            if (window.navigator && window.navigator.msSaveOrOpenBlob) {
                window.navigator.msSaveOrOpenBlob(newBlob);
                return;
            }

            // For other browsers: 
            // Create a link pointing to the ObjectURL containing the blob.
            const newData = window.URL.createObjectURL(newBlob);
            const link = document.createElement('a');
            link.href = newData;
            link.download = 'file.pdf';
            link.click();
            setTimeout(() => {
                // For Firefox it is necessary to delay revoking the ObjectURL
                window.URL.revokeObjectURL(newData);
            }, 100);
        }); */

        /* workbook.xlsx.writeBuffer().then((data: Uint8Array) => {
            const blob = new Blob([data], { type: this.blobType });
            FileSaver.saveAs(blob, 'test.xlsx');
            console.log(data);
            const len = data.byteLength;
            const chArray = new Array(len);
            for (let i = 0; i < len; i++) {
                chArray[i] = String.fromCharCode(data[i]);
            }

            const loadingTask = PDFJS.getDocument({data});
            // console.log(loadingTask);

            loadingTask.promise.then((pdf) => {
                console.log(pdf);
            });
        }); */

        /* workbook.xlsx.writeBuffer().then((data: Uint8Array) => {
            const len = data.byteLength;
            const chArray = new Array(len);
            for (let i = 0; i < len; i++) {
                chArray[i] = String.fromCharCode(data[i]);
            }
            console.log(window.btoa(chArray.join('')));

            const base64String = btoa(String.fromCharCode.apply(null, data));
            console.log(base64String);
            const blob = new Blob([data], { type: this.blobType });
            FileSaver.saveAs(blob, 'test.xlsx');
        }); */
    }
}
