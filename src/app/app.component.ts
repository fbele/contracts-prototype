import { Component } from '@angular/core';
import { ExportExcelService } from './services/export-excel.service';
import { PdfExportService } from './services/pdf-export.service';

@Component({
    selector: 'app-root',
    templateUrl: './app.component.html',
    styleUrls: ['./app.component.scss']
})
export class AppComponent {
    public title = 'OBI - Aufträge Prototype';
    // public pdfSrc = './assets/files/dummy.pdf';
    public pdfSrc: Uint8Array;

    public listItems = [{
        cell1 : 'Lorem ipsum dolor sit amet.',
        cell2 : 'Duis rutrum vel quam at sodales.',
        cell3 : 'Quisque porttitor elementum lorem eu laoreet.',
        cell4 : 'Vestibulum quis vehicula mauris, sed blandit lorem. In ultricies efficitur turpis, eget porttitor ipsum blandit vitae. Nunc ipsum nisi, pretium vitae accumsan sed, bibendum a mi. Mauris a condimentum orci, et porta ligula.',
    }];

    constructor(private _exportExcelService: ExportExcelService,
                private _pdfExportService: PdfExportService) { }

    public addRow() {
        const listItemClone = Object.create(this.listItems[0]);
        this.listItems.push(listItemClone);
    }

    public exportExcel() {
        console.log('export');
        this._exportExcelService.exportExcelFile();
        /* this._exportExcelService.exportExcelFile().then((data: Uint8Array) => {
            console.log(this.pdfSrc);
            this.pdfSrc = data;
            console.log(this.pdfSrc);
        }); */
    }

    public exportPDF() {
        console.log('export');
        this._pdfExportService.exportPDFFile();
    }
}
