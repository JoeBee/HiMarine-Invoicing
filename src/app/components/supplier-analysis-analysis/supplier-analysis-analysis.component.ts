import { Component, OnInit, OnDestroy } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { SupplierAnalysisService, ExcelRowData, SupplierAnalysisFileInfo, SupplierAnalysisFileSet } from '../../services/supplier-analysis.service';
import { LoggingService } from '../../services/logging.service';
import { Subscription } from 'rxjs';
import * as ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import JSZip from 'jszip';
import * as XLSX from 'xlsx';

interface AnalysisSet {
    id: number;
    invoiceFiles: SupplierAnalysisFileInfo[];
    supplierQuotationFiles: SupplierAnalysisFileInfo[];
    invoiceData: ExcelRowData[];
    supplierQuotationData: ExcelRowData[][];
    invoiceHeaders: string[];
    supplierQuotationHeaders: string[][];
    rowCountsMatch: boolean;
    exportFileName: string;
    invoiceLabel: string;
    tableExpanded: boolean;
}

@Component({
    selector: 'app-supplier-analysis-analysis',
    standalone: true,
    imports: [CommonModule, FormsModule],
    templateUrl: './supplier-analysis-analysis.component.html',
    styleUrls: ['./supplier-analysis-analysis.component.scss']
})
export class SupplierAnalysisAnalysisComponent implements OnInit, OnDestroy {
    analysisSets: AnalysisSet[] = [];
    isLoading = false;
    exportFileName: string = '';
    private filesSubscription?: Subscription;

    constructor(
        private supplierAnalysisService: SupplierAnalysisService,
        private loggingService: LoggingService
    ) { }

    async ngOnInit(): Promise<void> {
        this.updateExportFileName();
        await this.loadData();

        this.filesSubscription = this.supplierAnalysisService.fileSets$.subscribe(async () => {
            await this.loadData();
        });
    }

    ngOnDestroy(): void {
        if (this.filesSubscription) {
            this.filesSubscription.unsubscribe();
        }
    }

    async loadData(): Promise<void> {
        this.isLoading = true;
        this.analysisSets = [];

        const fileSets = this.supplierAnalysisService.getFileSets();

        for (const fileSet of fileSets) {
            if (fileSet.files.length === 0) continue;

            const invoiceFiles = fileSet.files.filter(f => f.category === 'Invoice');
            const supplierQuotationFiles = fileSet.files.filter(f => f.category === 'Supplier Quotations');

            const allFiles = [...invoiceFiles, ...supplierQuotationFiles];
            const allRowCounts = allFiles.map(f => f.rowCount);
            const rowCountsMatch = allRowCounts.length > 0 && allRowCounts.every(count => count === allRowCounts[0]);

            const analysisSet: AnalysisSet = {
                id: fileSet.id,
                invoiceFiles,
                supplierQuotationFiles,
                invoiceData: [],
                supplierQuotationData: [],
                invoiceHeaders: [],
                supplierQuotationHeaders: [],
                rowCountsMatch,
                exportFileName: '',
                invoiceLabel: '',
                tableExpanded: false
            };

            if (rowCountsMatch && allFiles.length > 0) {
                try {
                    if (invoiceFiles.length > 0) {
                        const invoiceResult = await this.supplierAnalysisService.extractDataFromFile(invoiceFiles[0]);
                        analysisSet.invoiceData = invoiceResult.rows;
                        analysisSet.invoiceHeaders = invoiceResult.headers.map(header =>
                            header === 'Provisions' ? 'Invoice' : header
                        );
                    }

                    for (const file of supplierQuotationFiles) {
                        const result = await this.supplierAnalysisService.extractDataFromFile(file);
                        const filteredHeaders = result.headers.filter(header => {
                            const headerLower = header.toLowerCase().trim();
                            return headerLower === 'description' || headerLower === 'remark' ||
                                headerLower === 'price' || headerLower === 'total' ||
                                headerLower.includes('description') || headerLower.includes('remark') ||
                                headerLower.includes('price') || headerLower.includes('total');
                        });
                        analysisSet.supplierQuotationHeaders.push(filteredHeaders);

                        const filteredRows = result.rows.map(row => {
                            const filteredRow: ExcelRowData = {};
                            filteredHeaders.forEach(header => {
                                if (row[header] !== undefined) {
                                    filteredRow[header] = row[header];
                                }
                            });
                            return filteredRow;
                        });
                        analysisSet.supplierQuotationData.push(filteredRows);
                    }

                    if (invoiceFiles.length > 0) {
                        const nameWithoutExt = invoiceFiles[0].fileName.replace(/\.(xlsx|xls|xlsm)$/i, '');
                        const words = nameWithoutExt.trim().split(/\s+/);

                        let label = nameWithoutExt;
                        if (words.length > 0) {
                            label = words.length >= 2 ? `${words[0]} ${words[1]}` : words[0];
                        }

                        analysisSet.invoiceLabel = label;
                        analysisSet.exportFileName = `_OUTPUT ${nameWithoutExt}`;
                    } else {
                        analysisSet.exportFileName = `_OUTPUT Invoice Set ${fileSet.id}`;
                    }

                    this.analysisSets.push(analysisSet);
                    if (this.analysisSets.length === 1) {
                        analysisSet.tableExpanded = true;
                    }

                } catch (error) {
                    this.loggingService.logError(error as Error, 'data_extraction_error', 'SupplierAnalysisAnalysisComponent', { setId: fileSet.id });
                }
            } else if (allFiles.length > 0) {
                this.analysisSets.push(analysisSet);
            }
        }

        this.updateExportFileName();
        this.isLoading = false;
    }

    updateExportFileName(): void {
        const now = new Date();
        const year = now.getFullYear();
        const month = String(now.getMonth() + 1).padStart(2, '0');
        const day = String(now.getDate()).padStart(2, '0');
        this.exportFileName = `_Invoice ${year}${month}${day}`;
    }

    toggleTable(set: AnalysisSet): void {
        set.tableExpanded = !set.tableExpanded;
    }

    getCellValue(row: ExcelRowData, header: string): any {
        const dataKey = header === 'Invoice' ? 'Provisions' : header;
        const value = row[dataKey] !== undefined ? row[dataKey] : '';
        if (this.isPriceOrTotalColumn(header)) {
            return this.formatToTwoDecimals(value);
        }
        return value;
    }

    getSupplierQuotationValue(set: AnalysisSet, fileIndex: number, rowIndex: number, header: string): any {
        if (fileIndex < set.supplierQuotationData.length &&
            rowIndex < set.supplierQuotationData[fileIndex].length) {
            const row = set.supplierQuotationData[fileIndex][rowIndex];
            const value = row[header] !== undefined ? row[header] : '';
            if (this.isPriceOrTotalColumn(header)) {
                return this.formatToTwoDecimals(value);
            }
            return value;
        }
        return '';
    }

    private isPriceOrTotalColumn(header: string): boolean {
        const headerLower = header.toLowerCase().trim();
        return headerLower === 'price' || headerLower === 'total' ||
            headerLower.includes('price') || headerLower.includes('total');
    }

    private formatToTwoDecimals(value: any): string {
        if (value === null || value === undefined || value === '') {
            return '';
        }
        const numValue = Number(value);
        if (isNaN(numValue)) {
            return String(value);
        }
        return numValue.toFixed(2);
    }

    getSupplierQuotationHeaderLength(set: AnalysisSet, index: number): number {
        return set.supplierQuotationHeaders[index]?.length || 1;
    }

    isRightAlignedHeader(header: string): boolean {
        const headerLower = header.toLowerCase().trim();
        return headerLower === 'qty' || headerLower === 'price' || headerLower === 'total' ||
            headerLower.includes('price') || headerLower.includes('total');
    }

    shouldHighlightCell(set: AnalysisSet, fileIndex: number, rowIndex: number, header: string): boolean {
        if (rowIndex >= set.invoiceData.length) {
            return false;
        }

        const headerLower = header.toLowerCase().trim();
        const isDescriptionOrRemark = headerLower === 'description' || headerLower === 'remark' ||
            headerLower.includes('description') || headerLower.includes('remark');

        if (!isDescriptionOrRemark) {
            return false;
        }

        const supplierValue = this.getSupplierQuotationValue(set, fileIndex, rowIndex, header);
        const invoiceRow = set.invoiceData[rowIndex];

        let invoiceHeader = header;
        for (const invHeader of set.invoiceHeaders) {
            const invHeaderLower = invHeader.toLowerCase().trim();
            if (invHeaderLower === headerLower ||
                (headerLower.includes('description') && invHeaderLower.includes('description')) ||
                (headerLower.includes('remark') && invHeaderLower.includes('remark'))) {
                invoiceHeader = invHeader;
                break;
            }
        }

        const invoiceValue = this.getCellValue(invoiceRow, invoiceHeader);

        return this.valuesDiffer(supplierValue, invoiceValue);
    }

    private normalizeValue(value: any): string {
        if (value === null || value === undefined) {
            return '';
        }
        return String(value).trim();
    }

    private valuesDiffer(value1: any, value2: any): boolean {
        const normalized1 = this.normalizeValue(value1);
        const normalized2 = this.normalizeValue(value2);
        return normalized1 !== normalized2;
    }

    getRecordCount(set: AnalysisSet): number {
        return set.invoiceData.length;
    }

    getFileCount(set: AnalysisSet): number {
        return set.invoiceFiles.length + set.supplierQuotationFiles.length;
    }

    private async extractHeaderInfo(fileInfo: SupplierAnalysisFileInfo): Promise<{ rows: any[][], styles: Map<string, Partial<ExcelJS.Style>> }> {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();

            reader.onload = async (e: any) => {
                try {
                    const arrayBuffer = e.target.result;
                    const workbook = new ExcelJS.Workbook();
                    await workbook.xlsx.load(arrayBuffer);
                    const worksheet = workbook.getWorksheet(1);

                    if (!worksheet) {
                        reject(new Error('Worksheet not found'));
                        return;
                    }

                    const topLeftCellRef = XLSX.utils.decode_cell(fileInfo.topLeftCell);
                    const headerRowCount = topLeftCellRef.r;

                    const headerRows: any[][] = [];
                    const stylesMap = new Map<string, Partial<ExcelJS.Style>>();

                    for (let excelRowNum = 1; excelRowNum <= headerRowCount; excelRowNum++) {
                        const row = worksheet.getRow(excelRowNum);
                        const rowData: any[] = [];
                        let maxCol = 0;

                        row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
                            const colIndex = colNumber - 1;
                            rowData[colIndex] = cell.value;
                            maxCol = Math.max(maxCol, colNumber);
                            const styleKey = `${excelRowNum - 1}-${colIndex}`;
                            const style: Partial<ExcelJS.Style> = {};
                            if (cell.font) style.font = { ...cell.font };
                            if (cell.alignment) style.alignment = { ...cell.alignment };
                            if (cell.border) style.border = { ...cell.border };
                            if (cell.fill) style.fill = { ...cell.fill };
                            if (cell.numFmt) style.numFmt = cell.numFmt;
                            if (cell.protection) style.protection = { ...cell.protection };
                            if (Object.keys(style).length > 0) stylesMap.set(styleKey, style);
                        });

                        for (let col = 0; col < maxCol; col++) {
                            if (rowData[col] === undefined) rowData[col] = '';
                        }
                        headerRows.push(rowData);
                    }
                    resolve({ rows: headerRows, styles: stylesMap });
                } catch (error) {
                    reject(error);
                }
            };
            reader.onerror = () => reject(new Error('File reading error'));
            reader.readAsArrayBuffer(fileInfo.file);
        });
    }

    private isDescriptionOrRemarkColumn(header: string): boolean {
        const headerLower = header.toLowerCase().trim();
        return headerLower === 'description' || headerLower === 'remark' ||
            headerLower.includes('description') || headerLower.includes('remark');
    }

    private hasColumnDifferences(set: AnalysisSet, fileIndex: number, header: string): boolean {
        if (!this.isDescriptionOrRemarkColumn(header)) return true;

        let invoiceHeader = header;
        const headerLower = header.toLowerCase().trim();
        for (const invHeader of set.invoiceHeaders) {
            const invHeaderLower = invHeader.toLowerCase().trim();
            if (invHeaderLower === headerLower ||
                (headerLower.includes('description') && invHeaderLower.includes('description')) ||
                (headerLower.includes('remark') && invHeaderLower.includes('remark'))) {
                invoiceHeader = invHeader;
                break;
            }
        }

        for (let rowIndex = 0; rowIndex < set.invoiceData.length; rowIndex++) {
            const supplierValue = this.getSupplierQuotationValue(set, fileIndex, rowIndex, header);
            const invoiceValue = this.getCellValue(set.invoiceData[rowIndex], invoiceHeader);
            if (this.valuesDiffer(supplierValue, invoiceValue)) return true;
        }
        return false;
    }

    private getFilteredHeadersForExport(set: AnalysisSet, fileIndex: number, headers: string[]): string[] {
        return headers.filter(header => {
            if (!this.isDescriptionOrRemarkColumn(header)) return true;
            return this.hasColumnDifferences(set, fileIndex, header);
        });
    }

    async exportToExcel(): Promise<void> {
        const exportSets = this.analysisSets.filter(s => s.rowCountsMatch && s.invoiceData.length > 0);

        if (exportSets.length === 0) {
            alert('No data available to export.');
            return;
        }

        this.loggingService.logButtonClick('export_to_excel', 'SupplierAnalysisAnalysisComponent', {
            fileName: this.exportFileName,
            setCount: exportSets.length
        });

        try {
            const workbook = new ExcelJS.Workbook();
            const worksheet = workbook.addWorksheet('Sheet1');
            worksheet.properties.showGridLines = false;
            worksheet.views = [{ showGridLines: false }];

            try {
                const fetchImage = async (path: string): Promise<{ buffer: ArrayBuffer; width: number; height: number }> => {
                    const response = await fetch(path);
                    const buffer = await response.arrayBuffer();
                    const blob = new Blob([buffer], { type: 'image/png' });
                    const url = URL.createObjectURL(blob);
                    const img = new Image();
                    await new Promise<void>((resolve, reject) => {
                        img.onload = () => { URL.revokeObjectURL(url); resolve(); };
                        img.onerror = () => { URL.revokeObjectURL(url); reject(new Error(`Failed to load image at ${path}`)); };
                        img.src = url;
                    });
                    return { buffer, width: img.naturalWidth, height: img.naturalHeight };
                };
                const topImage = await fetchImage('assets/images/HIMarineTopImage_sm.png');
                const topImageId = workbook.addImage({ buffer: topImage.buffer, extension: 'png' });
                worksheet.addImage(topImageId, { tl: { col: 0.75, row: 0.5 }, ext: { width: topImage.width, height: topImage.height } });
            } catch (error) {
                console.warn('Failed to load top image for workbook export:', error);
            }

            let currentRow = 1;

            if (exportSets.length > 0 && exportSets[0].invoiceFiles.length > 0) {
                const headerInfo = await this.extractHeaderInfo(exportSets[0].invoiceFiles[0]);
                for (let rowIndex = 0; rowIndex < headerInfo.rows.length; rowIndex++) {
                    const rowData = headerInfo.rows[rowIndex];
                    const headerRow = worksheet.getRow(currentRow);
                    for (let col = 0; col < rowData.length; col++) {
                        const cell = headerRow.getCell(col + 1);
                        const cellValue = rowData[col];
                        cell.value = (cellValue === '') ? null : cellValue;
                        const styleKey = `${rowIndex}-${col}`;
                        const originalStyle = headerInfo.styles.get(styleKey);
                        if (originalStyle) {
                            if (originalStyle.font) cell.font = { ...originalStyle.font, name: 'Cambria', size: 11 };
                            else cell.font = { name: 'Cambria', size: 11 };
                            if (originalStyle.alignment) cell.alignment = { ...originalStyle.alignment, wrapText: false };
                            else cell.alignment = { wrapText: false };
                            if (originalStyle.border) cell.border = originalStyle.border;
                            if (originalStyle.fill) cell.fill = originalStyle.fill;
                            if (originalStyle.numFmt) cell.numFmt = originalStyle.numFmt;
                            if (originalStyle.protection) cell.protection = originalStyle.protection;
                        } else {
                            cell.font = { name: 'Cambria', size: 11 };
                            cell.alignment = { wrapText: false };
                        }
                    }
                    currentRow++;
                }
                if (headerInfo.rows.length > 0) currentRow++;
            }

            const allSpacingColumns: number[] = [];

            for (const set of exportSets) {
                const invoiceHeadersLimited = set.invoiceHeaders.slice(0, 7);
                const filteredSupplierQuotationHeaders: string[][] = [];
                for (let fileIndex = 0; fileIndex < set.supplierQuotationHeaders.length; fileIndex++) {
                    filteredSupplierQuotationHeaders.push(
                        this.getFilteredHeadersForExport(set, fileIndex, set.supplierQuotationHeaders[fileIndex])
                    );
                }

                const sectionRow = worksheet.getRow(currentRow);
                const sectionCell = sectionRow.getCell(1);
                sectionCell.value = set.invoiceLabel || '';
                sectionCell.font = { name: 'Cambria', size: 22 };
                currentRow++;

                const headerRow1 = worksheet.getRow(currentRow);
                let col = 1;
                for (let i = 0; i < invoiceHeadersLimited.length; i++) {
                    const cell = headerRow1.getCell(col + i);
                    if (i === 0) cell.value = 'INVOICE';
                    cell.font = { name: 'Cambria', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
                    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF808080' } };
                }
                col += invoiceHeadersLimited.length;
                col += 2; // Spacing

                for (let i = 0; i < set.supplierQuotationFiles.length; i++) {
                    const headerLength = filteredSupplierQuotationHeaders[i].length;
                    for (let j = 0; j < headerLength; j++) {
                        const cell = headerRow1.getCell(col + j);
                        const headerName = filteredSupplierQuotationHeaders[i][j];
                        const headerLower = headerName.toLowerCase().trim();
                        if (j === 0) cell.value = set.supplierQuotationFiles[i].fileName;
                        cell.font = { name: 'Cambria', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
                        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4472C4' } }; // Blue instead of Orange
                        if (set.supplierQuotationFiles[i].discount !== undefined &&
                            set.supplierQuotationFiles[i].discount !== 0 &&
                            (headerLower === 'price' || headerLower.includes('price'))) {
                            const discountPercent = Math.round(set.supplierQuotationFiles[i].discount * 100);
                            if (j !== 0) {
                                cell.value = `Discount: ${discountPercent}%`;
                                cell.alignment = { horizontal: 'center' };
                            }
                        }
                    }
                    col += headerLength;
                    if (i < set.supplierQuotationFiles.length - 1) col += 2;
                }
                currentRow++;

                const headerRow2 = worksheet.getRow(currentRow);
                col = 1;
                for (const header of invoiceHeadersLimited) {
                    const cell = headerRow2.getCell(col);
                    cell.value = header;
                    cell.font = { name: 'Cambria', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
                    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF808080' } };
                    const headerLower = header.toLowerCase().trim();
                    if (headerLower.includes('price') || headerLower.includes('total')) {
                        cell.alignment = { horizontal: 'center', vertical: 'middle' };
                    }
                    col++;
                }
                col += 2;

                for (let fileIndex = 0; fileIndex < filteredSupplierQuotationHeaders.length; fileIndex++) {
                    for (const header of filteredSupplierQuotationHeaders[fileIndex]) {
                        const cell = headerRow2.getCell(col);
                        cell.value = header;
                        cell.font = { name: 'Cambria', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
                        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4472C4' } }; // Blue instead of Orange
                        const headerLower = header.toLowerCase().trim();
                        if (headerLower.includes('price') || headerLower.includes('total')) {
                            cell.alignment = { horizontal: 'center', vertical: 'middle' };
                        }
                        col++;
                    }
                    if (fileIndex < filteredSupplierQuotationHeaders.length - 1) col += 2;
                }
                currentRow++;

                const priceHighlightStats = new Map<number, { count: number, sum: number }>();
                const priceNonBestStats = new Map<number, { count: number, sum: number }>();

                for (let rowIndex = 0; rowIndex < set.invoiceData.length; rowIndex++) {
                    const dataRow = worksheet.getRow(currentRow);
                    col = 1;
                    const rowPriceValues: { col: number, value: number }[] = [];

                    for (const header of invoiceHeadersLimited) {
                        const cell = dataRow.getCell(col);
                        const value = this.getCellValue(set.invoiceData[rowIndex], header);
                        cell.value = value;
                        cell.font = { name: 'Cambria', size: 11 };

                        // Dark gray border for Invoice data
                        cell.border = {
                            top: { style: 'thin', color: { argb: 'FF404040' } },
                            left: { style: 'thin', color: { argb: 'FF404040' } },
                            bottom: { style: 'thin', color: { argb: 'FF404040' } },
                            right: { style: 'thin', color: { argb: 'FF404040' } }
                        };

                        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF5F5F5' } };
                        const headerLower = header.toLowerCase().trim();
                        if (headerLower.includes('price') || headerLower.includes('total')) {
                            if (value !== '' && value !== null) {
                                const numValue = Number(value);
                                if (!isNaN(numValue)) {
                                    cell.value = numValue;
                                    cell.numFmt = '$#,##0.00';
                                }
                            }
                        }
                        col++;
                    }
                    col += 2;

                    for (let fileIndex = 0; fileIndex < filteredSupplierQuotationHeaders.length; fileIndex++) {
                        for (const header of filteredSupplierQuotationHeaders[fileIndex]) {
                            const cell = dataRow.getCell(col);
                            const value = this.getSupplierQuotationValue(set, fileIndex, rowIndex, header);
                            cell.value = value;
                            cell.font = { name: 'Cambria', size: 11 };

                            // Dark gray border for Supplier Quotation data
                            cell.border = {
                                top: { style: 'thin', color: { argb: 'FF404040' } },
                                left: { style: 'thin', color: { argb: 'FF404040' } },
                                bottom: { style: 'thin', color: { argb: 'FF404040' } },
                                right: { style: 'thin', color: { argb: 'FF404040' } }
                            };

                            const headerLower = header.toLowerCase().trim();

                            // Orange background for Remark data only if cell contains data
                            if (headerLower.includes('remark') && value !== '' && value !== null && value !== undefined) {
                                cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFA500' } };
                            } else {
                                cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF5F5F5' } };
                            }

                            if (headerLower.includes('price') || headerLower.includes('total')) {
                                if (value !== '' && value !== null) {
                                    const numValue = Number(value);
                                    if (!isNaN(numValue)) {
                                        cell.value = numValue;
                                        cell.numFmt = '$#,##0.00';
                                        if (headerLower.includes('price')) rowPriceValues.push({ col, value: numValue });
                                    }
                                }
                            }
                            col++;
                        }
                        if (fileIndex < filteredSupplierQuotationHeaders.length - 1) col += 2;
                    }

                    if (rowPriceValues.length > 0) {
                        const minPrice = Math.min(...rowPriceValues.map(p => p.value));
                        const minEntries = rowPriceValues.filter(p => p.value === minPrice);
                        for (const entry of rowPriceValues) {
                            const cell = dataRow.getCell(entry.col);
                            if (entry.value === minPrice) {
                                const stats = priceHighlightStats.get(entry.col) || { count: 0, sum: 0 };
                                stats.count++;
                                stats.sum += entry.value;
                                priceHighlightStats.set(entry.col, stats);
                                cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: minEntries.length > 1 ? 'FFC6EFCE' : 'FFFFFF00' } };
                                cell.font = { name: 'Cambria', size: 11, bold: true };
                            } else {
                                const stats = priceNonBestStats.get(entry.col) || { count: 0, sum: 0 };
                                stats.count++;
                                stats.sum += entry.value;
                                priceNonBestStats.set(entry.col, stats);
                            }
                        }
                    }
                    currentRow++;
                }

                const totalRow = currentRow + 3;
                const totalRowObj = worksheet.getRow(totalRow);
                let checkCol = 1;

                for (const header of invoiceHeadersLimited) {
                    const hLower = header.toLowerCase().trim();
                    if (hLower.includes('price')) {
                        const labelCell = totalRowObj.getCell(checkCol);
                        labelCell.value = 'TOTAL SUM';
                        labelCell.font = { name: 'Cambria', size: 11, bold: true };
                        labelCell.alignment = { horizontal: 'right' };
                    }
                    if (hLower.includes('total')) {
                        let totalSum = 0;
                        for (let r = 0; r < set.invoiceData.length; r++) {
                            const val = this.getCellValue(set.invoiceData[r], header);
                            if (val) totalSum += Number(val) || 0;
                        }
                        const sumCell = totalRowObj.getCell(checkCol);
                        sumCell.value = totalSum;
                        sumCell.numFmt = '$#,##0.00';
                        sumCell.font = { name: 'Cambria', size: 11, bold: true };
                    }
                    checkCol++;
                }
                checkCol += 2;

                for (let fileIndex = 0; fileIndex < filteredSupplierQuotationHeaders.length; fileIndex++) {
                    for (const header of filteredSupplierQuotationHeaders[fileIndex]) {
                        const hLower = header.toLowerCase().trim();
                        if (hLower.includes('price')) {
                            const labelCell = totalRowObj.getCell(checkCol);
                            labelCell.value = 'TOTAL SUM';
                            labelCell.font = { name: 'Cambria', size: 11, bold: true };
                            labelCell.alignment = { horizontal: 'right' };

                            const nonBestRow = totalRow + 2;
                            const countRow = totalRow + 3;
                            const nonBestRowObj = worksheet.getRow(nonBestRow);
                            const countRowObj = worksheet.getRow(countRow);

                            const nonBest = priceNonBestStats.get(checkCol);
                            if (nonBest) {
                                const c = nonBestRowObj.getCell(checkCol);
                                c.value = `${nonBest.count} items`;
                                c.font = { name: 'Cambria', size: 11, bold: true };
                                c.alignment = { horizontal: 'right' };
                                const s = nonBestRowObj.getCell(checkCol + 1);
                                s.value = nonBest.sum;
                                s.numFmt = '$#,##0.00';
                                s.font = { name: 'Cambria', size: 11, bold: true };
                            }
                            const best = priceHighlightStats.get(checkCol);
                            if (best) {
                                const c = countRowObj.getCell(checkCol);
                                c.value = `${best.count} items`;
                                c.font = { name: 'Cambria', size: 11, bold: true };
                                c.alignment = { horizontal: 'right' };
                                c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } };
                                const s = countRowObj.getCell(checkCol + 1);
                                s.value = best.sum;
                                s.numFmt = '$#,##0.00';
                                s.font = { name: 'Cambria', size: 11, bold: true };
                                s.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } };
                            }
                        }
                        if (hLower.includes('total')) {
                            let totalSum = 0;
                            for (let r = 0; r < set.invoiceData.length; r++) {
                                const val = this.getSupplierQuotationValue(set, fileIndex, r, header);
                                if (val) totalSum += Number(val) || 0;
                            }
                            const sumCell = totalRowObj.getCell(checkCol);
                            sumCell.value = totalSum;
                            sumCell.numFmt = '$#,##0.00';
                            sumCell.font = { name: 'Cambria', size: 11, bold: true };
                        }
                        checkCol++;
                    }
                    if (fileIndex < filteredSupplierQuotationHeaders.length - 1) checkCol += 2;
                }

                allSpacingColumns.push(8, 9);
                let spacerCol = 9;
                for (let i = 0; i < filteredSupplierQuotationHeaders.length - 1; i++) {
                    spacerCol += filteredSupplierQuotationHeaders[i].length;
                    allSpacingColumns.push(spacerCol + 1, spacerCol + 2);
                    spacerCol += 2;
                }
                currentRow += 25;
            }

            worksheet.getColumn(1).width = 56 / 7;
            worksheet.getColumn(2).width = 231 / 7;
            worksheet.getColumn(3).width = 120 / 7;
            worksheet.getColumn(4).width = 51 / 7;
            worksheet.getColumn(5).width = 87 / 7;
            worksheet.getColumn(6).width = 146 / 7;
            worksheet.getColumn(7).width = 99 / 7;

            const columnMaxWidths: Map<number, number> = new Map();
            const remarkColumns: Set<number> = new Set();
            const priceAndTotalColumns: Set<number> = new Set();

            // First pass: identify remark and price/total columns
            worksheet.eachRow((row, rowNumber) => {
                row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
                    if (colNumber >= 8 && !allSpacingColumns.includes(colNumber)) {
                        // Check if cell has blue header background (Supplier Quotation section)
                        if (cell.fill && (cell.fill as ExcelJS.FillPattern).fgColor?.argb === 'FF4472C4') {
                            const headerName = cell.value ? String(cell.value).toLowerCase() : '';
                            if (headerName.includes('remark')) {
                                remarkColumns.add(colNumber);
                            }
                            if (headerName.includes('price') || headerName.includes('total')) {
                                priceAndTotalColumns.add(colNumber);
                            }
                        }
                        // Check if cell has gray header background (Invoice section)
                        if (cell.fill && (cell.fill as ExcelJS.FillPattern).fgColor?.argb === 'FF808080') {
                            const headerName = cell.value ? String(cell.value).toLowerCase() : '';
                            if (headerName.includes('price') || headerName.includes('total')) {
                                priceAndTotalColumns.add(colNumber);
                            }
                        }

                        // Calculate content length for autofit
                        let contentLength = 0;
                        if (cell.value !== null && cell.value !== undefined) {
                            contentLength = String(cell.value).length;
                            if (cell.numFmt && (cell.numFmt.includes('$') || cell.numFmt.includes('#'))) contentLength += 3;
                            if (rowNumber <= 3) contentLength = Math.max(contentLength * 1.2, contentLength + 3);
                        }
                        const currentMax = columnMaxWidths.get(colNumber) || 0;
                        columnMaxWidths.set(colNumber, Math.max(currentMax, contentLength));
                    }
                });
            });

            // Second pass: apply column widths
            columnMaxWidths.forEach((maxWidth, colNumber) => {
                if (!allSpacingColumns.includes(colNumber)) {
                    if (colNumber >= 8) {
                        if (remarkColumns.has(colNumber)) {
                            // Remark columns: 25 pixels (approximately 3.57 Excel units)
                            worksheet.getColumn(colNumber).width = 25 / 7;
                        } else if (priceAndTotalColumns.has(colNumber)) {
                            // Price and Total columns: autofit
                            const calculatedWidth = Math.max(maxWidth + 3, 10);
                            worksheet.getColumn(colNumber).width = calculatedWidth;
                        } else {
                            // Other columns: autofit
                            const calculatedWidth = Math.max(maxWidth + 3, 10);
                            worksheet.getColumn(colNumber).width = calculatedWidth;
                        }
                    }
                }
            });

            const spacingColumnWidth = 11 / 7;
            allSpacingColumns.forEach(c => worksheet.getColumn(c).width = spacingColumnWidth);

            const buffer = await workbook.xlsx.writeBuffer();
            const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
            let fileName = this.exportFileName.trim();
            if (!fileName.endsWith('.xlsx')) fileName += '.xlsx';
            saveAs(blob, fileName);

            this.loggingService.logExport('excel_exported', { fileName, fileSize: blob.size }, 'SupplierAnalysisAnalysisComponent');
        } catch (error) {
            this.loggingService.logError(error as Error, 'excel_export_error', 'SupplierAnalysisAnalysisComponent');
            alert('An error occurred while exporting to Excel.');
        }
    }
}
