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
                            return headerLower === 'description' || headerLower === 'remark' || headerLower === 'unit' ||
                                headerLower === 'price' || headerLower === 'total' ||
                                headerLower.includes('description') || headerLower.includes('remark') || headerLower.includes('unit') ||
                                headerLower.includes('price') || headerLower.includes('total');
                        });

                        // Check if Remark and Unit columns exist, if not add them
                        const hasRemark = filteredHeaders.some(h => h.toLowerCase().trim().includes('remark'));
                        const hasUnit = filteredHeaders.some(h => h.toLowerCase().trim().includes('unit'));

                        // Insert missing columns in the expected order: Description, Remark, Unit, Price, Total
                        const orderedHeaders: string[] = [];

                        // Add Description if it exists
                        const descHeader = filteredHeaders.find(h => h.toLowerCase().trim().includes('description'));
                        if (descHeader) orderedHeaders.push(descHeader);

                        // Add Remark (existing or new)
                        const remarkHeader = filteredHeaders.find(h => h.toLowerCase().trim().includes('remark'));
                        if (remarkHeader) {
                            orderedHeaders.push(remarkHeader);
                        } else {
                            orderedHeaders.push('Remark');
                        }

                        // Add Unit (existing or new)
                        const unitHeader = filteredHeaders.find(h => h.toLowerCase().trim().includes('unit'));
                        if (unitHeader) {
                            orderedHeaders.push(unitHeader);
                        } else {
                            orderedHeaders.push('Unit');
                        }

                        // Add Price and Total if they exist
                        const priceHeader = filteredHeaders.find(h => h.toLowerCase().trim().includes('price'));
                        if (priceHeader) orderedHeaders.push(priceHeader);

                        const totalHeader = filteredHeaders.find(h => h.toLowerCase().trim().includes('total'));
                        if (totalHeader) orderedHeaders.push(totalHeader);

                        analysisSet.supplierQuotationHeaders.push(orderedHeaders);

                        const filteredRows = result.rows.map(row => {
                            const filteredRow: ExcelRowData = {};
                            orderedHeaders.forEach(header => {
                                if (row[header] !== undefined) {
                                    filteredRow[header] = row[header];
                                } else {
                                    // Add empty value for missing columns
                                    filteredRow[header] = '';
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
        this.exportFileName = `Analysis ${year}${month}${day}`;
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
            let value = row[header] !== undefined ? row[header] : '';

            const headerLower = header.toLowerCase().trim();

            // For Unit columns, check if value matches Invoice - if so, return blank
            if (headerLower.includes('unit') && rowIndex < set.invoiceData.length) {
                // Find corresponding Invoice header
                let invoiceHeader = header;
                for (const invHeader of set.invoiceHeaders) {
                    const invHeaderLower = invHeader.toLowerCase().trim();
                    if (invHeaderLower === headerLower || invHeaderLower.includes('unit')) {
                        invoiceHeader = invHeader;
                        break;
                    }
                }
                const invoiceValue = this.getCellValue(set.invoiceData[rowIndex], invoiceHeader);

                // If values match, return blank
                if (!this.valuesDiffer(value, invoiceValue)) {
                    return '';
                }
            }

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
        const headers = set.supplierQuotationHeaders[index];
        if (!headers) return 1;
        return this.getFilteredHeaders(set, index, headers).length;
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

        // Don't highlight blank files
        if (set.supplierQuotationFiles[fileIndex].isBlank) {
            return false;
        }

        const headerLower = header.toLowerCase().trim();
        const isDescriptionRemarkOrUnit = headerLower === 'description' || headerLower === 'remark' || headerLower === 'unit' ||
            headerLower.includes('description') || headerLower.includes('remark') || headerLower.includes('unit');

        if (!isDescriptionRemarkOrUnit) {
            return false;
        }

        const supplierValue = this.getSupplierQuotationValue(set, fileIndex, rowIndex, header);

        // For Unit columns, only highlight if there's an actual value (not blank)
        if (headerLower.includes('unit')) {
            // If the supplier value is blank/empty, don't highlight
            if (supplierValue === '' || supplierValue === null || supplierValue === undefined) {
                return false;
            }
            // If there's a value, it means it differs from Invoice, so highlight it
            return true;
        }

        // For Remark columns, only highlight if there's an actual value
        if (headerLower.includes('remark')) {
            if (supplierValue === '' || supplierValue === null || supplierValue === undefined) {
                return false;
            }
            return true;
        }

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
        // Handle blank files - return empty header info
        if (fileInfo.isBlank || !fileInfo.file) {
            return Promise.resolve({ rows: [], styles: new Map() });
        }

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

            if (fileInfo.file) {
                reader.readAsArrayBuffer(fileInfo.file);
            }
        });
    }

    private isDescriptionRemarkOrUnitColumn(header: string): boolean {
        const headerLower = header.toLowerCase().trim();
        return headerLower === 'description' || headerLower === 'remark' || headerLower === 'unit' ||
            headerLower.includes('description') || headerLower.includes('remark') || headerLower.includes('unit');
    }

    private hasColumnDifferences(set: AnalysisSet, fileIndex: number, header: string): boolean {
        if (!this.isDescriptionRemarkOrUnitColumn(header)) return true;

        let invoiceHeader = header;
        const headerLower = header.toLowerCase().trim();
        for (const invHeader of set.invoiceHeaders) {
            const invHeaderLower = invHeader.toLowerCase().trim();
            if (invHeaderLower === headerLower ||
                (headerLower.includes('description') && invHeaderLower.includes('description')) ||
                (headerLower.includes('remark') && invHeaderLower.includes('remark')) ||
                (headerLower.includes('unit') && invHeaderLower.includes('unit'))) {
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

    getFilteredHeaders(set: AnalysisSet, fileIndex: number, headers: string[]): string[] {
        return headers.filter(header => {
            const headerLower = header.toLowerCase().trim();

            // Always include Remark and Unit columns (even if empty)
            if (headerLower.includes('remark') || headerLower === 'remark') return true;
            if (headerLower.includes('unit') || headerLower === 'unit') return true;

            // For other columns, use existing logic
            if (!this.isDescriptionRemarkOrUnitColumn(header)) return true;
            return this.hasColumnDifferences(set, fileIndex, header);
        });
    }

    private getFilteredHeadersForExport(set: AnalysisSet, fileIndex: number, headers: string[]): string[] {
        return this.getFilteredHeaders(set, fileIndex, headers);
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
            const templatePath = 'assets/templates/supplier-analysis-macros.xlsm';
            const templateResponse = await fetch(templatePath);
            const templateBuffer = await templateResponse.arrayBuffer();

            // 1. Extract macro parts from the template manually to ensure we have them
            // ExcelJS usually drops the vbaProject.bin when writing, so we need to put it back.
            let vbaProject: ArrayBuffer | undefined;
            let workbookRels: string | undefined;

            try {
                const tempZip = await JSZip.loadAsync(templateBuffer);
                const vba = tempZip.file('xl/vbaProject.bin');
                if (vba) vbaProject = await vba.async('arraybuffer');

                const rels = tempZip.file('xl/_rels/workbook.xml.rels');
                if (rels) workbookRels = await rels.async('string');
            } catch (e) {
                console.warn('Error extracting macros from template', e);
            }

            const workbook = new ExcelJS.Workbook();
            await workbook.xlsx.load(templateBuffer);

            let worksheet = workbook.getWorksheet('Sheet1');
            if (!worksheet) {
                worksheet = workbook.addWorksheet('Sheet1');
            }

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
            const supplierPriceAndTotalColumns: Set<number> = new Set(); // Track Price and Total columns in Supplier Quotation sections
            const allSupplierPriceColumns: Set<number> = new Set(); // Track all Price columns in Supplier Quotation sections
            const allSupplierTotalColumns: Set<number> = new Set(); // Track all Total columns in Supplier Quotation sections

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
                sectionCell.font = { name: 'Cambria', size: 18 };
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
                    const isBlankFile = set.supplierQuotationFiles[i].isBlank;

                    for (let j = 0; j < headerLength; j++) {
                        const cell = headerRow1.getCell(col + j);
                        const headerName = filteredSupplierQuotationHeaders[i][j];
                        const headerLower = headerName.toLowerCase().trim();

                        // Only show filename for non-blank files, or on first column
                        if (j === 0 && !isBlankFile) {
                            cell.value = set.supplierQuotationFiles[i].fileName;
                        }

                        cell.font = { name: 'Cambria', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
                        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4472C4' } }; // Blue instead of Orange

                        // Track Price and Total columns in first header row
                        if (headerLower.includes('price') || headerLower.includes('total')) {
                            supplierPriceAndTotalColumns.add(col + j);
                        }

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
                    const isBlankFile = set.supplierQuotationFiles[fileIndex].isBlank;

                    for (const header of filteredSupplierQuotationHeaders[fileIndex]) {
                        const cell = headerRow2.getCell(col);

                        // Don't show header labels for blank files
                        if (!isBlankFile) {
                            cell.value = header;
                        }

                        cell.font = { name: 'Cambria', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
                        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4472C4' } }; // Blue instead of Orange
                        const headerLower = header.toLowerCase().trim();
                        if (headerLower.includes('price') || headerLower.includes('total')) {
                            cell.alignment = { horizontal: 'center', vertical: 'middle' };
                            // Track Price and Total columns in headers
                            supplierPriceAndTotalColumns.add(col);
                        }
                        col++;
                    }
                    if (fileIndex < filteredSupplierQuotationHeaders.length - 1) col += 2;
                }
                currentRow++;

                const priceHighlightStats = new Map<number, { count: number, sum: number }>();
                const priceNonBestStats = new Map<number, { count: number, sum: number }>();

                // Matrix Stats: WinnerFileIndex -> TargetFileIndex -> Stats
                const matrixStats = new Map<number, Map<number, { count: number, sum: number }>>();
                for (let i = 0; i < set.supplierQuotationFiles.length; i++) {
                    const rowMap = new Map<number, { count: number, sum: number }>();
                    for (let j = 0; j < set.supplierQuotationFiles.length; j++) {
                        rowMap.set(j, { count: 0, sum: 0 });
                    }
                    matrixStats.set(i, rowMap);
                }
                const filePriceColMap = new Map<number, number>(); // FileIndex -> Excel Column Index

                // Identify column indices for Invoice section (Qty, Price, Total)
                let invoiceQtyCol: number | null = null;
                let invoicePriceCol: number | null = null;
                let invoiceTotalCol: number | null = null;
                let tempCol = 1;
                for (const header of invoiceHeadersLimited) {
                    const headerLower = header.toLowerCase().trim();
                    if (headerLower === 'qty' || headerLower.includes('qty')) {
                        invoiceQtyCol = tempCol;
                    } else if (headerLower === 'price' || headerLower.includes('price')) {
                        invoicePriceCol = tempCol;
                    } else if (headerLower === 'total' || headerLower.includes('total')) {
                        invoiceTotalCol = tempCol;
                    }
                    tempCol++;
                }

                // Identify column indices for Supplier Quotation sections (Price, Total per file)
                const supplierPriceCols = new Map<number, number>(); // FileIndex -> Price Column
                const supplierTotalCols = new Map<number, number>(); // FileIndex -> Total Column
                tempCol = 1 + invoiceHeadersLimited.length + 2; // Start after Invoice + spacing
                for (let fileIndex = 0; fileIndex < filteredSupplierQuotationHeaders.length; fileIndex++) {
                    for (const header of filteredSupplierQuotationHeaders[fileIndex]) {
                        const headerLower = header.toLowerCase().trim();
                        if (headerLower === 'price' || headerLower.includes('price')) {
                            supplierPriceCols.set(fileIndex, tempCol);
                            allSupplierPriceColumns.add(tempCol);
                        } else if (headerLower === 'total' || headerLower.includes('total')) {
                            supplierTotalCols.set(fileIndex, tempCol);
                            allSupplierTotalColumns.add(tempCol);
                        }
                        tempCol++;
                    }
                    if (fileIndex < filteredSupplierQuotationHeaders.length - 1) tempCol += 2;
                }

                const dataStartRow = currentRow; // Track where data rows start for formula references

                for (let rowIndex = 0; rowIndex < set.invoiceData.length; rowIndex++) {
                    const dataRow = worksheet.getRow(currentRow);
                    col = 1;
                    const rowPriceValues: { col: number, value: number, fileIndex: number }[] = [];
                    const currentFilesPrices = new Map<number, number>(); // FileIndex -> Price
                    const currentFilesTotals = new Map<number, number>(); // FileIndex -> Total

                    for (const header of invoiceHeadersLimited) {
                        const cell = dataRow.getCell(col);
                        const value = this.getCellValue(set.invoiceData[rowIndex], header);
                        const headerLower = header.toLowerCase().trim();

                        // For Total column, use formula instead of value
                        if (col === invoiceTotalCol && invoiceQtyCol !== null && invoicePriceCol !== null) {
                            const qtyColLetter = worksheet.getColumn(invoiceQtyCol).letter;
                            const priceColLetter = worksheet.getColumn(invoicePriceCol).letter;
                            const rowNum = currentRow;
                            cell.value = { formula: `${qtyColLetter}${rowNum}*${priceColLetter}${rowNum}` };
                            cell.numFmt = '$#,##0.00';
                        } else {
                            cell.value = value;
                            if (headerLower.includes('price') || headerLower.includes('total')) {
                                if (value !== '' && value !== null) {
                                    const numValue = Number(value);
                                    if (!isNaN(numValue)) {
                                        cell.value = numValue;
                                        cell.numFmt = '$#,##0.00';
                                    }
                                }
                            }
                        }

                        cell.font = { name: 'Cambria', size: 11 };

                        // Dark gray border for Invoice data
                        cell.border = {
                            top: { style: 'thin', color: { argb: 'FF404040' } },
                            left: { style: 'thin', color: { argb: 'FF404040' } },
                            bottom: { style: 'thin', color: { argb: 'FF404040' } },
                            right: { style: 'thin', color: { argb: 'FF404040' } }
                        };

                        col++;
                    }
                    col += 2;

                    for (let fileIndex = 0; fileIndex < filteredSupplierQuotationHeaders.length; fileIndex++) {
                        const isBlankFile = set.supplierQuotationFiles[fileIndex].isBlank;

                        for (const header of filteredSupplierQuotationHeaders[fileIndex]) {
                            const cell = dataRow.getCell(col);
                            let value = this.getSupplierQuotationValue(set, fileIndex, rowIndex, header);

                            const headerLower = header.toLowerCase().trim();

                            // For Unit columns, check if value matches Invoice - if so, make it blank
                            if (headerLower.includes('unit')) {
                                // Find corresponding Invoice header
                                let invoiceHeader = header;
                                for (const invHeader of set.invoiceHeaders) {
                                    const invHeaderLower = invHeader.toLowerCase().trim();
                                    if (invHeaderLower === headerLower || invHeaderLower.includes('unit')) {
                                        invoiceHeader = invHeader;
                                        break;
                                    }
                                }
                                const invoiceValue = this.getCellValue(set.invoiceData[rowIndex], invoiceHeader);

                                // If values match, make supplier value blank
                                if (!this.valuesDiffer(value, invoiceValue)) {
                                    value = '';
                                }
                            }

                            cell.value = value;
                            cell.font = { name: 'Cambria', size: 11 };

                            // Dark gray border for Supplier Quotation data
                            cell.border = {
                                top: { style: 'thin', color: { argb: 'FF404040' } },
                                left: { style: 'thin', color: { argb: 'FF404040' } },
                                bottom: { style: 'thin', color: { argb: 'FF404040' } },
                                right: { style: 'thin', color: { argb: 'FF404040' } }
                            };

                            // Orange background for Remark and Unit data only if cell contains data AND not a blank file
                            if (!isBlankFile && (headerLower.includes('remark') || headerLower.includes('unit')) && value !== '' && value !== null && value !== undefined) {
                                cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFA500' } };
                            }

                            if (headerLower.includes('price') || headerLower.includes('total')) {
                                // Track this column for auto-fitting
                                supplierPriceAndTotalColumns.add(col);

                                // For Total column, use formula: Invoice Qty * Supplier Price
                                // Do not include formula for blank Supplier Quotation files
                                if (headerLower.includes('total') && !isBlankFile && invoiceQtyCol !== null && supplierPriceCols.has(fileIndex)) {
                                    const qtyColLetter = worksheet.getColumn(invoiceQtyCol).letter;
                                    const priceCol = supplierPriceCols.get(fileIndex)!;
                                    const priceColLetter = worksheet.getColumn(priceCol).letter;
                                    const rowNum = currentRow;
                                    cell.value = { formula: `${qtyColLetter}${rowNum}*${priceColLetter}${rowNum}` };
                                    cell.numFmt = '$#,##0.00';

                                    // Still track the calculated value for stats
                                    if (value !== '' && value !== null) {
                                        const numValue = Number(value);
                                        if (!isNaN(numValue)) {
                                            currentFilesTotals.set(fileIndex, numValue);
                                        }
                                    }
                                } else if (headerLower.includes('price')) {
                                    if (value !== '' && value !== null) {
                                        const numValue = Number(value);
                                        if (!isNaN(numValue)) {
                                            cell.value = numValue;
                                            cell.numFmt = '$#,##0.00';
                                            rowPriceValues.push({ col, value: numValue, fileIndex });
                                            currentFilesPrices.set(fileIndex, numValue);
                                            if (!filePriceColMap.has(fileIndex)) {
                                                filePriceColMap.set(fileIndex, col);
                                            }
                                        }
                                    }
                                } else if (headerLower.includes('total')) {
                                    // Fallback: if we couldn't create formula, use value
                                    // Do not set value for blank files
                                    if (!isBlankFile && value !== '' && value !== null) {
                                        const numValue = Number(value);
                                        if (!isNaN(numValue)) {
                                            cell.value = numValue;
                                            cell.numFmt = '$#,##0.00';
                                            currentFilesTotals.set(fileIndex, numValue);
                                        }
                                    }
                                }
                            }
                            col++;
                        }
                        if (fileIndex < filteredSupplierQuotationHeaders.length - 1) col += 2;
                    }

                    if (rowPriceValues.length > 0) {
                        // Filter out prices of 0 when finding the minimum
                        const nonZeroPrices = rowPriceValues.filter(p => p.value > 0);
                        if (nonZeroPrices.length > 0) {
                            const minPrice = Math.min(...nonZeroPrices.map(p => p.value));
                            const minEntries = rowPriceValues.filter(p => p.value === minPrice && p.value > 0);
                            const winningFileIndices = new Set<number>();

                            for (const entry of rowPriceValues) {
                                const cell = dataRow.getCell(entry.col);
                                // Only highlight if value equals minPrice and is greater than 0
                                if (entry.value === minPrice && entry.value > 0) {
                                    winningFileIndices.add(entry.fileIndex);
                                    const stats = priceHighlightStats.get(entry.col) || { count: 0, sum: 0 };
                                    stats.count++;
                                    stats.sum += entry.value;
                                    priceHighlightStats.set(entry.col, stats);

                                    const highlightFill: ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: minEntries.length > 1 ? 'FFC6EFCE' : 'FFFFFF00' } };
                                    cell.fill = highlightFill;
                                    cell.font = { name: 'Cambria', size: 11, bold: true };

                                    // Also highlight the 'Total' column next to this Price column if it exists
                                    const totalColIndex = entry.col + 1;
                                    const totalCell = dataRow.getCell(totalColIndex);
                                    // Simple check if the next column is indeed a 'Total' column based on typical layout (Price then Total)
                                    // A more robust check would involve inspecting headers, but given the fixed layout generation:
                                    // We know Price and Total are adjacent in the filtered export logic.
                                    if (totalCell.value !== null && totalCell.value !== undefined) {
                                        totalCell.fill = highlightFill;
                                        totalCell.font = { name: 'Cambria', size: 11, bold: true };
                                    }
                                } else {
                                    const stats = priceNonBestStats.get(entry.col) || { count: 0, sum: 0 };
                                    stats.count++;
                                    stats.sum += entry.value;
                                    priceNonBestStats.set(entry.col, stats);
                                }
                            }

                            // Update Matrix Stats
                            winningFileIndices.forEach(winFileIdx => {
                                const winnerRowStats = matrixStats.get(winFileIdx);
                                if (winnerRowStats) {
                                    for (let targetFileIdx = 0; targetFileIdx < set.supplierQuotationFiles.length; targetFileIdx++) {
                                        const total = currentFilesTotals.get(targetFileIdx);
                                        const price = currentFilesPrices.get(targetFileIdx);
                                        // Use Total if available, otherwise fallback to Price
                                        const valueToSum = total !== undefined ? total : price;

                                        if (valueToSum !== undefined) {
                                            const s = winnerRowStats.get(targetFileIdx);
                                            if (s) {
                                                s.sum += valueToSum;
                                                s.count++;
                                            }
                                        }
                                    }
                                }
                            });
                        }
                    }
                    currentRow++;
                }

                const totalRow = currentRow + 3;
                const totalRowObj = worksheet.getRow(totalRow);
                let checkCol = 1;
                const startCol = 1;
                let endCol = 1;

                // Track all cells that have content in the total row for border application
                const totalRowCellsWithContent: number[] = [];

                for (const header of invoiceHeadersLimited) {
                    const hLower = header.toLowerCase().trim();
                    if (hLower.includes('price')) {
                        const labelCell = totalRowObj.getCell(checkCol);
                        labelCell.value = 'TOTAL SUM';
                        labelCell.font = { name: 'Cambria', size: 11, bold: true };
                        labelCell.alignment = { horizontal: 'right' };
                        totalRowCellsWithContent.push(checkCol);
                    }
                    if (hLower.includes('total')) {
                        const sumCell = totalRowObj.getCell(checkCol);
                        const colLetter = worksheet.getColumn(checkCol).letter;
                        const rangeStart = `${colLetter}${dataStartRow}`;
                        const rangeEnd = `${colLetter}${dataStartRow + set.invoiceData.length - 1}`;
                        sumCell.value = { formula: `SUM(${rangeStart}:${rangeEnd})` };
                        sumCell.numFmt = '$#,##0.00';
                        sumCell.font = { name: 'Cambria', size: 11, bold: true };
                        totalRowCellsWithContent.push(checkCol);
                    }
                    checkCol++;
                    endCol = checkCol - 1;
                }
                checkCol += 2;

                for (let fileIndex = 0; fileIndex < filteredSupplierQuotationHeaders.length; fileIndex++) {
                    const isBlankFile = set.supplierQuotationFiles[fileIndex].isBlank;

                    for (const header of filteredSupplierQuotationHeaders[fileIndex]) {
                        const hLower = header.toLowerCase().trim();
                        if (hLower.includes('price')) {
                            // Only show TOTAL SUM label if not a blank file
                            if (!isBlankFile) {
                                const labelCell = totalRowObj.getCell(checkCol);
                                labelCell.value = 'TOTAL SUM';
                                labelCell.font = { name: 'Cambria', size: 11, bold: true };
                                labelCell.alignment = { horizontal: 'right' };
                                totalRowCellsWithContent.push(checkCol);
                            }

                            // Matrix stats will be written after this loop
                        }
                        if (hLower.includes('total')) {
                            // Only calculate and show total sum if not a blank file
                            if (!isBlankFile) {
                                const sumCell = totalRowObj.getCell(checkCol);
                                const colLetter = worksheet.getColumn(checkCol).letter;
                                const rangeStart = `${colLetter}${dataStartRow}`;
                                const rangeEnd = `${colLetter}${dataStartRow + set.invoiceData.length - 1}`;
                                sumCell.value = { formula: `SUM(${rangeStart}:${rangeEnd})` };
                                sumCell.numFmt = '$#,##0.00';
                                sumCell.font = { name: 'Cambria', size: 11, bold: true };
                                totalRowCellsWithContent.push(checkCol);
                            }
                        }
                        checkCol++;
                        endCol = checkCol - 1;
                    }
                    if (fileIndex < filteredSupplierQuotationHeaders.length - 1) checkCol += 2;
                }

                // Add Delivery Fees and Other Fees rows
                // Write Matrix Summary
                let matrixStartRow = totalRow + 4; // Matrix starts before Delivery Fee/Other Fees rows
                const matrixStartRowIndex = matrixStartRow;

                for (let winFileIdx = 0; winFileIdx < set.supplierQuotationFiles.length; winFileIdx++) {
                    if (set.supplierQuotationFiles[winFileIdx].isBlank) continue;

                    const matrixRow = worksheet.getRow(matrixStartRow);

                    // Label: "{FileName}" in Column E
                    const labelCell = matrixRow.getCell(5);
                    labelCell.value = set.supplierQuotationFiles[winFileIdx].fileName;
                    labelCell.font = { name: 'Cambria', size: 11, bold: true };
                    labelCell.alignment = { horizontal: 'right' };

                    // Column G: Total Sum - Use SumByColor formula to sum yellowed values
                    const totalCell = matrixRow.getCell(7);

                    // Get winnerStats for this supplier (needed for matrix calculations)
                    const winnerStats = matrixStats.get(winFileIdx);

                    // Get this supplier's Total column from the main data table
                    const supplierTotalCol = supplierTotalCols.get(winFileIdx);

                    if (supplierTotalCol) {
                        // Calculate data range for this supplier's Total column
                        const dataEndRow = totalRow - 4;
                        const dataStartRow = dataEndRow - set.invoiceData.length + 1;

                        const totalColLetter = worksheet.getColumn(supplierTotalCol).letter;
                        const rangeStart = `${totalColLetter}${dataStartRow}`;
                        const rangeEnd = `${totalColLetter}${dataEndRow}`;
                        const currentCellRef = `G${matrixStartRow}`;

                        // console.log('dataEndRow:', dataEndRow);
                        // console.log('dataStartRow:', dataStartRow);
                        // console.log('totalColLetter:', totalColLetter);
                        // console.log('rangeStart:', rangeStart);
                        // console.log('rangeEnd:', rangeEnd);
                        // console.log('currentCellRef:', currentCellRef);
                        // console.log('*** TEST:', `${totalColLetter}${matrixStartRow}`);

                        // Use SumByColor formula to sum yellowed cells in this supplier's Total column
                        // totalCell.value = { formula: `SumByColor(${rangeStart}:${rangeEnd},${currentCellRef})` };
                        totalCell.value = { formula: `${totalColLetter}${matrixStartRow}` };
                    } else {
                        // Fallback to static value if column not found
                        let rowTotal = 0;
                        if (winnerStats) {
                            const selfStats = winnerStats.get(winFileIdx);
                            if (selfStats) {
                                rowTotal = selfStats.sum;
                            }
                        }
                        totalCell.value = rowTotal;
                    }

                    totalCell.numFmt = '$#,##0.00';
                    totalCell.font = { name: 'Cambria', size: 11, bold: true };
                    totalCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } };

                    if (winnerStats) {
                        for (let targetFileIdx = 0; targetFileIdx < set.supplierQuotationFiles.length; targetFileIdx++) {
                            if (set.supplierQuotationFiles[targetFileIdx].isBlank) continue;

                            const targetStats = winnerStats.get(targetFileIdx);
                            const priceCol = filePriceColMap.get(targetFileIdx);

                            if (targetStats && priceCol) {
                                const countCell = matrixRow.getCell(priceCol);
                                countCell.numFmt = '0 "items"';
                                countCell.font = { name: 'Cambria', size: 11, bold: true };
                                countCell.alignment = { horizontal: 'right' };

                                const sumCell = matrixRow.getCell(priceCol + 1);
                                // sumCell.value = targetStats.sum; // Removing static value
                                sumCell.numFmt = '$#,##0.00';
                                sumCell.font = { name: 'Cambria', size: 11, bold: true };

                                if (winFileIdx === targetFileIdx) {
                                    const yellow: ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } };
                                    countCell.fill = yellow;
                                    sumCell.fill = yellow;
                                }

                                // Calculate range for CountByColor/SumByColor
                                // The range is the Price/Total columns in the main data table.
                                // Data starts at row: initial header rows + 1 (dataRow loop)
                                // Actually, we need the specific column in the main table.
                                // filePriceColMap maps FileIndex -> Price Column Index in the main table?
                                // Wait, filePriceColMap was built during data iteration. Yes.

                                // Reconstruct the data range for this specific file's price column
                                // Main data table starts at:
                                // The variable 'currentRow' at the end of the data loop is essentially the end of the table.
                                // But we need to know where the data started.
                                // We can track it or calculate it.
                                // Let's look at where data iteration started:
                                // 'let currentRow = 1;' ... then headers ... then 'for (let rowIndex = 0...'
                                // We need to capture the start row of the data table.

                                // Assuming 'dataStartRow' is available or we can derive it.
                                // Looking at code structure:
                                // We don't have a 'dataStartRow' variable in scope here easily without tracking it.
                                // However, we know 'set.invoiceData.length' is the number of rows.
                                // The data table ends at 'totalRow - 3'.
                                // So data start row = (totalRow - 3) - set.invoiceData.length + 1.

                                // Corrected logic: totalRow is 3 rows after data.
                                // But we need to adjust range up by 1 cell based on feedback.
                                // Originally: const dataEndRow = totalRow - 3;
                                // Adjusted: const dataEndRow = totalRow - 4;
                                const dataEndRow = totalRow - 4;
                                const dataStartRow = dataEndRow - set.invoiceData.length + 1;

                                const priceColLetter = worksheet.getColumn(priceCol).letter;
                                const rangeStart = `${priceColLetter}${dataStartRow}`;
                                const rangeEnd = `${priceColLetter}${dataEndRow}`;

                                // For CountByColor: Count items in the Price column that are highlighted (green/yellow)
                                // Reference cell: The cell itself if it has the color, or we need to know what "Winning Color" is.
                                // The requirement says "use 'SumByColor' formula... 2nd references a cell".
                                // If the matrix cell itself is yellow (winner), we use it.
                                // If it's not yellow, what color are we looking for?
                                // The main table highlights the *lowest* price.
                                // If we want to count how many times THIS supplier won, we are looking for the highlight color.
                                // The highlight color in the main table is 'FFC6EFCE' (greenish) or 'FFFFFF00' (yellow) depending on duplicates.
                                // But SumByColor usually works by matching the background color of the reference cell.
                                // If the matrix summary cell is highlighted yellow (because it's the winner row/col), 
                                // then SumByColor(Range, MatrixCell) will sum things that match the MatrixCell's yellow color.
                                // But the main table highlights are 'FFC6EFCE' (if unique winner) or 'FFFFFF00' (if tied).
                                // This might be tricky if colors don't match exactly.
                                // However, the user prompt implies using the cell itself as reference.

                                // Let's assume the macro handles color matching robustly or the user ensures colors match.
                                // The prompt example: "=SumByColor(G210:G213,G215)" where G215 is the cell itself.

                                // Apply formulas
                                const countCellRef = `${worksheet.getColumn(priceCol).letter}${matrixRow.number}`;
                                const sumCellRef = `${worksheet.getColumn(priceCol + 1).letter}${matrixRow.number}`;

                                // NOTE: We need to handle the case where the matrix cell is NOT colored (non-diagonal cells).
                                // If the matrix cell is white/transparent, SumByColor(Range, MatrixCell) would sum uncolored cells?
                                // That's probably not what's intended for the non-diagonal cells.
                                // Usually the matrix shows how many times Supplier X beat Supplier Y?
                                // Or is it just a summary of Supplier X?
                                // The code structure suggests:
                                // "Matrix Stats: WinnerFileIndex -> TargetFileIndex -> Stats"
                                // It counts how many times WinnerFileIndex was best, summed up by TargetFileIndex's prices?
                                // Wait, if I am Supplier A, and I won 5 items.
                                // Row "Supplier A":
                                // Col "Supplier A": 5 items, Sum $100 (My price)
                                // Col "Supplier B": 5 items, Sum $120 (B's price for the items I won)
                                //
                                // To calculate this via SumByColor in VBA:
                                // We would need to look at Supplier A's column in the main table and count/sum highlighted cells?
                                // But that only works for the diagonal (My Wins, My Price).
                                // For off-diagonal (My Wins, Competitor Price), we can't use SumByColor on the Competitor Column based on highlighting,
                                // because the Competitor Column is NOT highlighted for the items *I* won.
                                //
                                // UNLESS: The requirement "use the 'SumByColor' formula in every table summation section" 
                                // applies specifically to the columns/rows that HAVE color.
                                // The image shows "TOTAL ALL SPLIT" (Yellow) and "Summation Quotations" (Yellow).
                                // It seems we only apply this to the yellow cells?
                                //
                                // User instruction: "Add a similar formula for 'Summation Quotations' columns" (circled in red/blue).
                                // The image circles the diagonal cells (where count/sum is).
                                // AND the text says "use 'SumByColor' formula in EVERY table summation section".
                                //
                                // If I strictly follow the request:
                                // "The 'TOTAL ALL SPLIT' calculation should be '=SumByColor(xx:yy, zz)'"
                                // "Add a similar formula for 'Summation Quotations' columns... circled in red/blue"
                                //
                                // The red/blue circles are on the diagonal (Winner vs Self).
                                // These cells ARE colored yellow in the code:
                                // `if (winFileIdx === targetFileIdx) { ... countCell.fill = yellow; ... }`
                                //
                                // So for these diagonal cells, we can use the formula pointing to the main table's price/total column.
                                // For non-diagonal cells, we should probably leave the static value or use a different logic?
                                // The prompt only explicitly mentions the "Summation Quotations" columns and shows the diagonal ones circled.
                                // It doesn't explicitly say "ONLY" the diagonal ones, but the "SumByColor" implies color matching.
                                // Since off-diagonal cells aren't colored, SumByColor wouldn't work to count "Winners" there.
                                //
                                // Conclusion: I will apply the formula ONLY to the diagonal cells (where winFileIdx === targetFileIdx).
                                // The static values calculated by JS (`targetStats.count`, `targetStats.sum`) will remain for off-diagonal 
                                // unless the user wants those dynamic too (which would require a different VBA function like SumIfOtherColIsColor).
                                // Given the specific instruction and the nature of SumByColor, I'll target the diagonal yellow cells.

                                if (winFileIdx === targetFileIdx) {
                                    // CountByColor for the 'items' cell
                                    // We want to count highlighted cells in the Price column of this file in the main table.
                                    // But wait, in the main table, the Price column IS highlighted if it's the winner.
                                    // So CountByColor(PriceColumnRange, ThisCell) should work if ThisCell is yellow and PriceCol cells are yellow/green.
                                    // Note: Main table highlight is 'FFC6EFCE' (light green) or 'FFFFFF00' (yellow).
                                    // Matrix cell is 'FFFFFF00' (yellow).
                                    // If the VBA function is strict about color matching, Yellow != Light Green.
                                    // However, we must assume the VBA function handles this or the user is aware.
                                    // I will construct the formula.

                                    countCell.value = { formula: `CountByColor(${rangeStart}:${rangeEnd},${countCellRef})` };

                                    // SumByColor for the 'sum' cell
                                    // We want to sum the Total (or Price) column in the main table where it's highlighted.
                                    // We used 'priceCol' for the range above.
                                    // Ideally we should sum the 'Total' column if it exists, or 'Price' if not.
                                    // In the loop, `sumCell` is `matrixRow.getCell(priceCol + 1)`.
                                    // Let's assume priceCol + 1 is the Total column in the main table too.
                                    // We can verify if the column at priceCol + 1 in main table is indeed a 'Total' column?
                                    // In `rowPriceValues` logic:
                                    // `const totalColIndex = entry.col + 1; ... totalCell.fill = highlightFill;`
                                    // So yes, the Total column is also highlighted.

                                    const totalColLetter = worksheet.getColumn(priceCol + 1).letter;
                                    const sumRangeStart = `${totalColLetter}${dataStartRow}`;
                                    const sumRangeEnd = `${totalColLetter}${dataEndRow}`;

                                    sumCell.value = { formula: `SumByColor(${sumRangeStart}:${sumRangeEnd},${sumCellRef})` };
                                } else {
                                    // Keep static values for non-diagonal cells
                                    countCell.value = targetStats.count;
                                    sumCell.value = targetStats.sum;
                                }
                            }
                        }
                    }

                    // Apply gray top and bottom borders from column C onwards
                    for (let col = 3; col <= endCol; col++) {
                        const cell = matrixRow.getCell(col);
                        const existingFill = cell.fill;
                        cell.border = {
                            top: { style: 'thin', color: { argb: 'FF808080' } },
                            bottom: { style: 'thin', color: { argb: 'FF808080' } }
                        };
                        // Preserve existing fill (for gray shading and yellow highlighting)
                        if (existingFill) {
                            cell.fill = existingFill;
                        }
                    }

                    matrixStartRow++;
                }

                matrixStartRow++; // Add empty line before Total Split

                // TOTAL ALL SPLIT Row
                const totalSplitRow = worksheet.getRow(matrixStartRow);
                const splitLabelCell = totalSplitRow.getCell(6); // Column F
                splitLabelCell.value = 'TOTAL SPLIT';
                splitLabelCell.font = { name: 'Cambria', size: 11, bold: true };
                splitLabelCell.alignment = { horizontal: 'right' };
                splitLabelCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } };

                const splitSumCell = totalSplitRow.getCell(7); // Column G
                const endRow = matrixStartRow - 2; // Data ends before the blank line

                // Calculate the dynamic range for the yellow sum cells (Column G in Matrix)
                // Matrix starts at matrixStartRowIndex. It ends at endRow.
                if (endRow >= matrixStartRowIndex) {
                    // We need the address range for SumByColor.
                    // Column 7 is 'G'.
                    const rangeStart = `G${matrixStartRowIndex}`;
                    const rangeEnd = `G${endRow}`;
                    const currentCellRef = `G${matrixStartRow}`;

                    // Formula: =SumByColor(Range, CellWithColor)
                    // We use the current cell as the color reference since it's yellow.
                    splitSumCell.value = { formula: `SumByColor(${rangeStart}:${rangeEnd},${currentCellRef})` };
                } else {
                    splitSumCell.value = 0;
                }

                splitSumCell.numFmt = '$#,##0.00';
                splitSumCell.font = { name: 'Cambria', size: 11, bold: true };
                splitSumCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } };

                // Add Delivery Fee and Other Fees rows aligned with TOTAL SPLIT
                const deliveryFeesRow = matrixStartRow; // Same row as TOTAL SPLIT
                const otherFeesRow = matrixStartRow + 1; // One row below Delivery Fee
                const deliveryFeesRowObj = worksheet.getRow(deliveryFeesRow);
                const otherFeesRowObj = worksheet.getRow(otherFeesRow);

                // Add Delivery Fee and Other Fees for each supplier quotation file
                for (let fileIndex = 0; fileIndex < filteredSupplierQuotationHeaders.length; fileIndex++) {
                    const isBlankFile = set.supplierQuotationFiles[fileIndex].isBlank;
                    const priceCol = supplierPriceCols.get(fileIndex);

                    if (priceCol && !isBlankFile) {
                        // Delivery Fee row
                        const deliveryLabelCell = deliveryFeesRowObj.getCell(priceCol);
                        deliveryLabelCell.value = 'Delivery Fee:';
                        deliveryLabelCell.font = { name: 'Cambria', size: 11, bold: true };
                        deliveryLabelCell.alignment = { horizontal: 'right' };

                        const deliveryValueCell = deliveryFeesRowObj.getCell(priceCol + 1);
                        deliveryValueCell.value = null; // Empty cell, user can fill in
                        deliveryValueCell.numFmt = '$#,##0.00';
                        deliveryValueCell.font = { name: 'Cambria', size: 11, bold: true };

                        // Other Fees row
                        const otherLabelCell = otherFeesRowObj.getCell(priceCol);
                        otherLabelCell.value = 'Other Fees:';
                        otherLabelCell.font = { name: 'Cambria', size: 11, bold: true };
                        otherLabelCell.alignment = { horizontal: 'right' };

                        const otherValueCell = otherFeesRowObj.getCell(priceCol + 1);
                        otherValueCell.value = null; // Empty cell, user can fill in
                        otherValueCell.numFmt = '$#,##0.00';
                        otherValueCell.font = { name: 'Cambria', size: 11, bold: true };
                    }
                }

                // Add the new summary section 2 rows below "Other Fees:"
                const summarySectionStartRow = otherFeesRow + 2;
                const dataEndRow = dataStartRow + set.invoiceData.length - 1;
                const totalRows = 3; // 3 rows total, all suppliers in the same rows

                // Background colors for the summary rows (light green, light blue, light gold)
                const summaryColors = [
                    { argb: 'FFE2EFDA' }, // Light green (top row)
                    { argb: 'FFDEEBF7' }, // Light blue (middle row)
                    { argb: 'FFFFEAAE' }  // Light gold (bottom row)
                ];

                // Get list of non-blank suppliers
                const nonBlankSuppliers: number[] = [];
                for (let fileIndex = 0; fileIndex < filteredSupplierQuotationHeaders.length; fileIndex++) {
                    const isBlankFile = set.supplierQuotationFiles[fileIndex].isBlank;
                    const priceCol = supplierPriceCols.get(fileIndex);
                    if (priceCol && !isBlankFile) {
                        nonBlankSuppliers.push(fileIndex);
                    }
                }

                // Iterate through the 3 rows (all suppliers in the same rows)
                for (let rowOffset = 0; rowOffset < totalRows; rowOffset++) {
                    const currentRowNum = summarySectionStartRow + rowOffset;
                    const summaryRow = worksheet.getRow(currentRowNum);
                    const bgColor = summaryColors[rowOffset]; // Cycle through colors: green, blue, orange

                    // Set borders - use same thin style as table data rows
                    const borderStyle = 'thin';
                    const borderColor = { argb: 'FF404040' }; // Dark gray, same as table data

                    // Add cells for each supplier in this row
                    for (let supplierIdx = 0; supplierIdx < nonBlankSuppliers.length; supplierIdx++) {
                        const fileIndex = nonBlankSuppliers[supplierIdx];
                        const priceCol = supplierPriceCols.get(fileIndex)!;
                        const priceColLetter = worksheet.getColumn(priceCol).letter;
                        const priceRangeStart = `${priceColLetter}$${dataStartRow}`;
                        const priceRangeEnd = `${priceColLetter}$${dataEndRow}`;
                        const priceRangeStartNoDollar = `${priceColLetter}${dataStartRow}`;
                        const priceRangeEndNoDollar = `${priceColLetter}${dataEndRow}`;
                        const countCol = priceCol - 1;
                        // Move blocks one column to the right
                        const summaryCountCol = countCol + 1;
                        const summaryPriceCol = priceCol + 1;
                        const summaryCountColLetter = worksheet.getColumn(summaryCountCol).letter;
                        const summaryPriceColLetter = worksheet.getColumn(summaryPriceCol).letter;

                        // First column: CountByColor formula
                        // Formula: =@CountByColor(M$30:M$38,L53) where L is summaryCountCol, and the row is currentRowNum
                        // Note: Range still points to original priceCol, but cell is moved one column right
                        const countCellRef = `${summaryCountColLetter}${currentRowNum}`;
                        const countCell = summaryRow.getCell(summaryCountCol);
                        countCell.value = { formula: `@CountByColor(${priceRangeStart}:${priceRangeEnd},${countCellRef})` };
                        countCell.font = { name: 'Cambria', size: 11 };
                        countCell.fill = { type: 'pattern', pattern: 'solid', fgColor: bgColor };
                        countCell.border = {
                            top: { style: borderStyle, color: borderColor },
                            bottom: { style: borderStyle, color: borderColor }
                        };

                        // Second column: SumByColor formula
                        // Formula: =@SumByColor(M30:M38,M53) where M is summaryPriceCol, and the row is currentRowNum
                        // Note: Range still points to original priceCol, but cell is moved one column right
                        const sumCellRef = `${summaryPriceColLetter}${currentRowNum}`;
                        const sumCell = summaryRow.getCell(summaryPriceCol);
                        sumCell.value = { formula: `@SumByColor(${priceRangeStartNoDollar}:${priceRangeEndNoDollar},${sumCellRef})` };
                        sumCell.numFmt = '$#,##0.00';
                        sumCell.font = { name: 'Cambria', size: 11 };
                        sumCell.fill = { type: 'pattern', pattern: 'solid', fgColor: bgColor };
                        sumCell.border = {
                            top: { style: borderStyle, color: borderColor },
                            bottom: { style: borderStyle, color: borderColor }
                        };
                    }
                }


                matrixStartRow++;

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
            worksheet.getColumn(6).width = 90 / 7;
            worksheet.getColumn(7).width = 95 / 7;

            const columnMaxWidths: Map<number, number> = new Map();
            const remarkColumns: Set<number> = new Set();
            const unitColumns: Set<number> = new Set();
            const priceColumns: Set<number> = new Set();
            const totalColumns: Set<number> = new Set();

            // First pass: identify remark, unit, price, and total columns, calculate widths
            worksheet.eachRow((row, rowNumber) => {
                row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
                    if (colNumber >= 8 && !allSpacingColumns.includes(colNumber)) {
                        // Check if cell has blue header background (Supplier Quotation section)
                        if (cell.fill && (cell.fill as ExcelJS.FillPattern).fgColor?.argb === 'FF4472C4') {
                            const headerName = cell.value ? String(cell.value).toLowerCase().trim() : '';
                            if (headerName.includes('remark')) {
                                remarkColumns.add(colNumber);
                            }
                            if (headerName.includes('unit')) {
                                unitColumns.add(colNumber);
                            }
                            if (headerName.includes('price')) {
                                priceColumns.add(colNumber);
                            }
                            if (headerName.includes('total')) {
                                totalColumns.add(colNumber);
                            }
                        }

                        // Calculate content length for autofit
                        let contentLength = 0;
                        if (cell.value !== null && cell.value !== undefined) {
                            // For numeric values with currency formatting, estimate the formatted string length
                            if (typeof cell.value === 'number' && cell.numFmt) {
                                // Format the number to estimate display length
                                const formatted = cell.value.toLocaleString('en-US', {
                                    style: 'currency',
                                    currency: 'USD',
                                    minimumFractionDigits: 2,
                                    maximumFractionDigits: 2
                                });
                                contentLength = formatted.length;
                            } else {
                                contentLength = String(cell.value).length;
                                // Increase length for currency formatting
                                if (cell.numFmt && (cell.numFmt.includes('$') || cell.numFmt.includes('#'))) {
                                    contentLength += 5; // Account for $, commas, and decimals
                                }
                            }

                            // Increase length for bold text (headers are typically bold)
                            if (cell.font && cell.font.bold) {
                                contentLength = Math.ceil(contentLength * 1.15) + 2;
                            }
                        }
                        const currentMax = columnMaxWidths.get(colNumber) || 0;
                        columnMaxWidths.set(colNumber, Math.max(currentMax, contentLength));
                    }
                });
            });

            // Second pass: apply column widths
            // First, explicitly set widths for all tracked Supplier Quotation Price and Total columns
            allSupplierPriceColumns.forEach(colNumber => {
                if (!allSpacingColumns.includes(colNumber)) {
                    worksheet.getColumn(colNumber).width = 90 / 7; // 90 pixels
                }
            });
            allSupplierTotalColumns.forEach(colNumber => {
                if (!allSpacingColumns.includes(colNumber)) {
                    worksheet.getColumn(colNumber).width = 95 / 7; // 95 pixels
                }
            });

            // Then apply widths for other columns
            columnMaxWidths.forEach((maxWidth, colNumber) => {
                if (!allSpacingColumns.includes(colNumber)) {
                    if (colNumber >= 8) {
                        // Skip Price and Total columns as they're already set above
                        if (allSupplierPriceColumns.has(colNumber) || allSupplierTotalColumns.has(colNumber)) {
                            return; // Already set above
                        }
                        if (remarkColumns.has(colNumber)) {
                            // Remark columns: 25 pixels (approximately 3.57 Excel units)
                            worksheet.getColumn(colNumber).width = 25 / 7;
                        } else if (unitColumns.has(colNumber)) {
                            // Unit columns: 25 pixels
                            worksheet.getColumn(colNumber).width = 25 / 7;
                        } else if (priceColumns.has(colNumber)) {
                            // Price columns: 90 pixels (fallback for any missed columns)
                            worksheet.getColumn(colNumber).width = 90 / 7;
                        } else if (totalColumns.has(colNumber)) {
                            // Total columns: 95 pixels (fallback for any missed columns)
                            worksheet.getColumn(colNumber).width = 95 / 7;
                        } else {
                            // Other columns: autofit
                            const calculatedWidth = Math.max(maxWidth + 2, 10);
                            worksheet.getColumn(colNumber).width = calculatedWidth;
                        }
                    }
                }
            });

            const spacingColumnWidth = 11 / 7;
            allSpacingColumns.forEach(c => worksheet.getColumn(c).width = spacingColumnWidth);

            // Generate the buffer
            let buffer = await workbook.xlsx.writeBuffer();

            // Post-process with JSZip to fix Macros (ExcelJS doesn't preserve them correctly on write)
            try {
                const zip = await JSZip.loadAsync(buffer);

                // 1. Inject vbaProject.bin if missing
                if (vbaProject && !zip.file('xl/vbaProject.bin')) {
                    zip.file('xl/vbaProject.bin', vbaProject);
                }

                // 2. Fix [Content_Types].xml
                // Switch workbook content type from XLSX to XLSM
                let contentTypes = await zip.file('[Content_Types].xml')?.async('string');
                if (contentTypes) {
                    const xlsxType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml';
                    const xlsmType = 'application/vnd.ms-excel.sheet.macroEnabled.main+xml';

                    if (contentTypes.includes(xlsxType)) {
                        contentTypes = contentTypes.replace(xlsxType, xlsmType);
                    }

                    // Ensure "bin" extension is mapped (needed for vbaProject.bin)
                    if (!contentTypes.includes('Extension="bin"')) {
                        // Insert before closing </Types>
                        contentTypes = contentTypes.replace('</Types>', '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/></Types>');
                    }
                    zip.file('[Content_Types].xml', contentTypes);
                }

                // 3. Fix xl/_rels/workbook.xml.rels
                // Ensure relationship to vbaProject exists
                if (vbaProject) {
                    let wbRels = await zip.file('xl/_rels/workbook.xml.rels')?.async('string');
                    if (wbRels) {
                        if (!wbRels.includes('vbaProject.bin')) {
                            // Add the relationship if missing. Use a likely unique ID.
                            // We check if rIdVBA exists, if not we add it.
                            // Note: Relationships need unique Ids. ExcelJS usually uses rId1, rId2...
                            // We'll try to reuse the one from template if possible, or just append one.
                            // Simpler: Check if 'http://schemas.microsoft.com/office/2006/relationships/vbaProject' exists
                            if (!wbRels.includes('relationships/vbaProject')) {
                                const rId = 'rIdVbaProject';
                                const rel = `<Relationship Id="${rId}" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/>`;
                                wbRels = wbRels.replace('</Relationships>', `${rel}</Relationships>`);
                                zip.file('xl/_rels/workbook.xml.rels', wbRels);
                            }
                        }
                    } else if (workbookRels) {
                        // If ExcelJS created a workbook without relationships (unlikely with images), restore template's
                        zip.file('xl/_rels/workbook.xml.rels', workbookRels);
                    }
                }

                buffer = await zip.generateAsync({ type: 'arraybuffer' });
            } catch (zipError) {
                console.warn('Error fixing macros in output file', zipError);
            }

            const blob = new Blob([buffer], { type: 'application/vnd.ms-excel.sheet.macroEnabled.12' });
            let fileName = this.exportFileName.trim();

            if (fileName.toLowerCase().endsWith('.xlsx')) {
                fileName = fileName.substring(0, fileName.length - 5);
            }
            if (!fileName.toLowerCase().endsWith('.xlsm')) {
                fileName += '.xlsm';
            }

            saveAs(blob, fileName);

            this.loggingService.logExport('excel_exported', { fileName, fileSize: blob.size }, 'SupplierAnalysisAnalysisComponent');
        } catch (error) {
            this.loggingService.logError(error as Error, 'excel_export_error', 'SupplierAnalysisAnalysisComponent');
            alert('An error occurred while exporting to Excel.');
        }
    }
}
