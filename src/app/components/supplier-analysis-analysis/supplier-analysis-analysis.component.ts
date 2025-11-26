import { Component, OnInit, OnDestroy } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { SupplierAnalysisService, ExcelRowData } from '../../services/supplier-analysis.service';
import { SupplierAnalysisFileInfo } from '../supplier-analysis-inputs/supplier-analysis-inputs.component';
import { LoggingService } from '../../services/logging.service';
import { Subscription } from 'rxjs';
import * as ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import JSZip from 'jszip';
import * as XLSX from 'xlsx';

@Component({
    selector: 'app-supplier-analysis-analysis',
    standalone: true,
    imports: [CommonModule, FormsModule],
    templateUrl: './supplier-analysis-analysis.component.html',
    styleUrls: ['./supplier-analysis-analysis.component.scss']
})
export class SupplierAnalysisAnalysisComponent implements OnInit, OnDestroy {
    invoiceFiles: SupplierAnalysisFileInfo[] = [];
    supplierQuotationFiles: SupplierAnalysisFileInfo[] = [];
    invoiceData: ExcelRowData[] = [];
    supplierQuotationData: ExcelRowData[][] = [];
    invoiceHeaders: string[] = [];
    supplierQuotationHeaders: string[][] = [];
    isLoading = false;
    rowCountsMatch = false;
    exportFileName: string = '';
    invoiceLabel: string = '';
    tableExpanded: boolean = false;

    // Set 2 properties
    invoiceFiles2: SupplierAnalysisFileInfo[] = [];
    supplierQuotationFiles2: SupplierAnalysisFileInfo[] = [];
    invoiceData2: ExcelRowData[] = [];
    supplierQuotationData2: ExcelRowData[][] = [];
    invoiceHeaders2: string[] = [];
    supplierQuotationHeaders2: string[][] = [];
    rowCountsMatch2 = false;
    exportFileName2: string = '';
    invoiceLabel2: string = '';
    tableExpanded2: boolean = false;

    private filesSubscription?: Subscription;
    private files2Subscription?: Subscription;

    constructor(
        private supplierAnalysisService: SupplierAnalysisService,
        private loggingService: LoggingService
    ) { }

    async ngOnInit(): Promise<void> {
        // Set default export filename on initialization
        this.updateExportFileName();

        await this.loadData();
        await this.loadData2();

        // Subscribe to file changes
        this.filesSubscription = this.supplierAnalysisService.files$.subscribe(async () => {
            await this.loadData();
        });

        this.files2Subscription = this.supplierAnalysisService.files2$.subscribe(async () => {
            await this.loadData2();
        });
    }

    ngOnDestroy(): void {
        if (this.filesSubscription) {
            this.filesSubscription.unsubscribe();
        }
        if (this.files2Subscription) {
            this.files2Subscription.unsubscribe();
        }
    }

    async loadData(): Promise<void> {
        this.isLoading = true;

        const allFiles = this.supplierAnalysisService.getFiles();
        this.invoiceFiles = allFiles.filter(f => f.category === 'Invoice');
        this.supplierQuotationFiles = allFiles.filter(f => f.category === 'Supplier Quotations');

        // Check if row counts match
        const allRowCounts = allFiles.map(f => f.rowCount);
        this.rowCountsMatch = allRowCounts.length > 0 &&
            allRowCounts.every(count => count === allRowCounts[0]);

        if (this.rowCountsMatch && allFiles.length > 0) {
            try {
                // Extract invoice data (use first invoice file if multiple)
                if (this.invoiceFiles.length > 0) {
                    const invoiceResult = await this.supplierAnalysisService.extractDataFromFile(this.invoiceFiles[0]);
                    this.invoiceData = invoiceResult.rows;
                    // Replace "Provisions" header with "Invoice"
                    this.invoiceHeaders = invoiceResult.headers.map(header =>
                        header === 'Provisions' ? 'Invoice' : header
                    );
                }

                // Extract supplier quotation data (all files)
                this.supplierQuotationData = [];
                this.supplierQuotationHeaders = [];
                for (const file of this.supplierQuotationFiles) {
                    const result = await this.supplierAnalysisService.extractDataFromFile(file);
                    // Filter headers to only show 'Price' and 'Total' columns
                    const filteredHeaders = result.headers.filter(header => {
                        const headerLower = header.toLowerCase().trim();
                        return headerLower === 'price' || headerLower === 'total' ||
                            headerLower.includes('price') || headerLower.includes('total');
                    });
                    this.supplierQuotationHeaders.push(filteredHeaders);
                    // Filter data rows to only include 'Price' and 'Total' columns
                    const filteredRows = result.rows.map(row => {
                        const filteredRow: ExcelRowData = {};
                        filteredHeaders.forEach(header => {
                            if (row[header] !== undefined) {
                                filteredRow[header] = row[header];
                            }
                        });
                        return filteredRow;
                    });
                    this.supplierQuotationData.push(filteredRows);
                }

                this.loggingService.logDataProcessing('supplier_analysis_data_loaded', {
                    invoiceFiles: this.invoiceFiles.length,
                    supplierQuotationFiles: this.supplierQuotationFiles.length,
                    rowCount: this.invoiceData.length
                }, 'SupplierAnalysisAnalysisComponent');

                // Set default export filename
                this.updateExportFileName();

                // Set default invoice label (last word of invoice filename)
                this.updateInvoiceLabel();
            } catch (error) {
                this.loggingService.logError(error as Error, 'data_extraction_error', 'SupplierAnalysisAnalysisComponent');
            }
        }

        this.isLoading = false;
    }

    async loadData2(): Promise<void> {
        const allFiles = this.supplierAnalysisService.getFiles2();
        this.invoiceFiles2 = allFiles.filter(f => f.category === 'Invoice');
        this.supplierQuotationFiles2 = allFiles.filter(f => f.category === 'Supplier Quotations');

        // Check if row counts match
        const allRowCounts = allFiles.map(f => f.rowCount);
        this.rowCountsMatch2 = allRowCounts.length > 0 &&
            allRowCounts.every(count => count === allRowCounts[0]);

        if (this.rowCountsMatch2 && allFiles.length > 0) {
            try {
                // Extract invoice data (use first invoice file if multiple)
                if (this.invoiceFiles2.length > 0) {
                    const invoiceResult = await this.supplierAnalysisService.extractDataFromFile(this.invoiceFiles2[0]);
                    this.invoiceData2 = invoiceResult.rows;
                    // Replace "Provisions" header with "Invoice"
                    this.invoiceHeaders2 = invoiceResult.headers.map(header =>
                        header === 'Provisions' ? 'Invoice' : header
                    );
                }

                // Extract supplier quotation data (all files)
                this.supplierQuotationData2 = [];
                this.supplierQuotationHeaders2 = [];
                for (const file of this.supplierQuotationFiles2) {
                    const result = await this.supplierAnalysisService.extractDataFromFile(file);
                    // Filter headers to only show 'Price' and 'Total' columns
                    const filteredHeaders = result.headers.filter(header => {
                        const headerLower = header.toLowerCase().trim();
                        return headerLower === 'price' || headerLower === 'total' ||
                            headerLower.includes('price') || headerLower.includes('total');
                    });
                    this.supplierQuotationHeaders2.push(filteredHeaders);
                    // Filter data rows to only include 'Price' and 'Total' columns
                    const filteredRows = result.rows.map(row => {
                        const filteredRow: ExcelRowData = {};
                        filteredHeaders.forEach(header => {
                            if (row[header] !== undefined) {
                                filteredRow[header] = row[header];
                            }
                        });
                        return filteredRow;
                    });
                    this.supplierQuotationData2.push(filteredRows);
                }

                this.loggingService.logDataProcessing('supplier_analysis_data_loaded_set2', {
                    invoiceFiles: this.invoiceFiles2.length,
                    supplierQuotationFiles: this.supplierQuotationFiles2.length,
                    rowCount: this.invoiceData2.length
                }, 'SupplierAnalysisAnalysisComponent');

                // Set default export filename
                this.updateExportFileName2();

                // Set default invoice label (last word of invoice filename)
                this.updateInvoiceLabel2();
            } catch (error) {
                this.loggingService.logError(error as Error, 'data_extraction_error_set2', 'SupplierAnalysisAnalysisComponent');
            }
        }
    }

    updateExportFileName(): void {
        // Default to "_Invoice YYYYMMDD" format
        const now = new Date();
        const year = now.getFullYear();
        const month = String(now.getMonth() + 1).padStart(2, '0');
        const day = String(now.getDate()).padStart(2, '0');
        this.exportFileName = `_Invoice ${year}${month}${day}`;
    }

    updateInvoiceLabel(): void {
        if (this.invoiceFiles.length > 0) {
            const invoiceFileName = this.invoiceFiles[0].fileName;
            // Remove extension if present
            const nameWithoutExt = invoiceFileName.replace(/\.(xlsx|xls|xlsm)$/i, '');
            // Split by spaces and get the last word
            const words = nameWithoutExt.trim().split(/\s+/);
            this.invoiceLabel = words.length > 0 ? words[words.length - 1] : nameWithoutExt;
        } else {
            this.invoiceLabel = '';
        }
    }

    updateExportFileName2(): void {
        if (this.invoiceFiles2.length > 0) {
            const invoiceFileName = this.invoiceFiles2[0].fileName;
            // Remove extension if present
            const nameWithoutExt = invoiceFileName.replace(/\.(xlsx|xls|xlsm)$/i, '');
            this.exportFileName2 = `_OUTPUT ${nameWithoutExt}`;
        } else {
            this.exportFileName2 = '_OUTPUT Invoice';
        }
    }

    updateInvoiceLabel2(): void {
        if (this.invoiceFiles2.length > 0) {
            const invoiceFileName = this.invoiceFiles2[0].fileName;
            // Remove extension if present
            const nameWithoutExt = invoiceFileName.replace(/\.(xlsx|xls|xlsm)$/i, '');
            // Split by spaces and get the last word
            const words = nameWithoutExt.trim().split(/\s+/);
            this.invoiceLabel2 = words.length > 0 ? words[words.length - 1] : nameWithoutExt;
        } else {
            this.invoiceLabel2 = '';
        }
    }

    getRecordCount(): number {
        return this.invoiceData.length;
    }

    getFileCount(): number {
        return this.invoiceFiles.length + this.supplierQuotationFiles.length;
    }

    getAllHeaders(): string[] {
        const allHeaders = new Set<string>(this.invoiceHeaders);
        this.supplierQuotationHeaders.forEach(headers => {
            headers.forEach(header => allHeaders.add(header));
        });
        return Array.from(allHeaders);
    }

    getCellValue(row: ExcelRowData, header: string): any {
        // Map "Invoice" header back to "Provisions" for data access
        const dataKey = header === 'Invoice' ? 'Provisions' : header;
        const value = row[dataKey] !== undefined ? row[dataKey] : '';
        // Format price and total columns to 2 decimal places
        if (this.isPriceOrTotalColumn(header)) {
            return this.formatToTwoDecimals(value);
        }
        return value;
    }

    getSupplierQuotationValue(fileIndex: number, rowIndex: number, header: string): any {
        if (fileIndex < this.supplierQuotationData.length &&
            rowIndex < this.supplierQuotationData[fileIndex].length) {
            const row = this.supplierQuotationData[fileIndex][rowIndex];
            const value = row[header] !== undefined ? row[header] : '';
            // Format price and total columns to 2 decimal places
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

    getSupplierQuotationHeaderLength(index: number): number {
        return this.supplierQuotationHeaders[index]?.length || 1;
    }

    getSupplierQuotationHeaderLength2(index: number): number {
        return this.supplierQuotationHeaders2[index]?.length || 1;
    }

    getCellValue2(row: ExcelRowData, header: string): any {
        // Map "Invoice" header back to "Provisions" for data access
        const dataKey = header === 'Invoice' ? 'Provisions' : header;
        const value = row[dataKey] !== undefined ? row[dataKey] : '';
        // Format price and total columns to 2 decimal places
        if (this.isPriceOrTotalColumn(header)) {
            return this.formatToTwoDecimals(value);
        }
        return value;
    }

    getSupplierQuotationValue2(fileIndex: number, rowIndex: number, header: string): any {
        if (fileIndex < this.supplierQuotationData2.length &&
            rowIndex < this.supplierQuotationData2[fileIndex].length) {
            const row = this.supplierQuotationData2[fileIndex][rowIndex];
            const value = row[header] !== undefined ? row[header] : '';
            // Format price and total columns to 2 decimal places
            if (this.isPriceOrTotalColumn(header)) {
                return this.formatToTwoDecimals(value);
            }
            return value;
        }
        return '';
    }

    getRecordCount2(): number {
        return this.invoiceData2.length;
    }

    getFileCount2(): number {
        return this.invoiceFiles2.length + this.supplierQuotationFiles2.length;
    }

    toggleTable(): void {
        this.tableExpanded = !this.tableExpanded;
    }

    toggleTable2(): void {
        this.tableExpanded2 = !this.tableExpanded2;
    }

    private async extractHeaderInfo(fileInfo: SupplierAnalysisFileInfo): Promise<{ rows: any[][], styles: Map<string, Partial<ExcelJS.Style>> }> {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();

            reader.onload = async (e: any) => {
                try {
                    const arrayBuffer = e.target.result;

                    // Use ExcelJS to read the file to preserve styling
                    const workbook = new ExcelJS.Workbook();
                    await workbook.xlsx.load(arrayBuffer);
                    const worksheet = workbook.getWorksheet(1);

                    if (!worksheet) {
                        reject(new Error('Worksheet not found'));
                        return;
                    }

                    const topLeftCellRef = XLSX.utils.decode_cell(fileInfo.topLeftCell);
                    const headerRowCount = topLeftCellRef.r; // 0-based row count

                    // Extract all rows above the topLeftCell with their styles
                    const headerRows: any[][] = [];
                    const stylesMap = new Map<string, Partial<ExcelJS.Style>>();

                    // ExcelJS uses 1-based row indexing
                    for (let excelRowNum = 1; excelRowNum <= headerRowCount; excelRowNum++) {
                        const row = worksheet.getRow(excelRowNum);
                        const rowData: any[] = [];
                        let maxCol = 0;

                        row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
                            // colNumber is 1-based in ExcelJS
                            const colIndex = colNumber - 1; // Convert to 0-based
                            rowData[colIndex] = cell.value;
                            maxCol = Math.max(maxCol, colNumber);

                            // Create a key for the style map: "row-col" format (0-based)
                            const styleKey = `${excelRowNum - 1}-${colIndex}`;

                            // Extract cell style
                            const style: Partial<ExcelJS.Style> = {};
                            if (cell.font) {
                                style.font = { ...cell.font };
                            }
                            if (cell.alignment) {
                                style.alignment = { ...cell.alignment };
                            }
                            if (cell.border) {
                                style.border = { ...cell.border };
                            }
                            if (cell.fill) {
                                style.fill = { ...cell.fill };
                            }
                            if (cell.numFmt) {
                                style.numFmt = cell.numFmt;
                            }
                            if (cell.protection) {
                                style.protection = { ...cell.protection };
                            }

                            if (Object.keys(style).length > 0) {
                                stylesMap.set(styleKey, style);
                            }
                        });

                        // Fill in empty cells up to maxCol
                        for (let col = 0; col < maxCol; col++) {
                            if (rowData[col] === undefined) {
                                rowData[col] = '';
                            }
                        }

                        headerRows.push(rowData);
                    }

                    resolve({ rows: headerRows, styles: stylesMap });
                } catch (error) {
                    reject(error);
                }
            };

            reader.onerror = () => {
                reject(new Error('File reading error'));
            };

            reader.readAsArrayBuffer(fileInfo.file);
        });
    }

    async exportToExcel(): Promise<void> {
        const hasTable1 = this.rowCountsMatch && this.invoiceData.length > 0;
        const hasTable2 = this.rowCountsMatch2 && this.invoiceData2.length > 0;

        if (!hasTable1 && !hasTable2) {
            alert('No data available to export.');
            return;
        }

        this.loggingService.logButtonClick('export_to_excel', 'SupplierAnalysisAnalysisComponent', {
            fileName: this.exportFileName,
            recordCount: this.getRecordCount() + this.getRecordCount2(),
            fileCount: this.getFileCount() + this.getFileCount2()
        });

        try {
            const workbook = new ExcelJS.Workbook();
            const worksheet = workbook.addWorksheet('Sheet1');

            // Turn off grid lines
            worksheet.properties.showGridLines = false;
            worksheet.views = [{ showGridLines: false }];

            // Add top image
            try {
                const fetchImage = async (path: string): Promise<{ buffer: ArrayBuffer; width: number; height: number }> => {
                    const response = await fetch(path);
                    const buffer = await response.arrayBuffer();
                    const blob = new Blob([buffer], { type: 'image/png' });
                    const url = URL.createObjectURL(blob);
                    const img = new Image();
                    await new Promise<void>((resolve, reject) => {
                        img.onload = () => {
                            URL.revokeObjectURL(url);
                            resolve();
                        };
                        img.onerror = () => {
                            URL.revokeObjectURL(url);
                            reject(new Error(`Failed to load image at ${path}`));
                        };
                        img.src = url;
                    });
                    return { buffer, width: img.naturalWidth, height: img.naturalHeight };
                };

                const topImage = await fetchImage('assets/images/HIMarineTopImage_sm.png');
                const topImageId = workbook.addImage({
                    buffer: topImage.buffer,
                    extension: 'png'
                });
                worksheet.addImage(topImageId, {
                    tl: { col: 0.75, row: 0.5 },
                    ext: { width: topImage.width, height: topImage.height }
                });
            } catch (error) {
                console.warn('Failed to load top image for workbook export:', error);
            }

            let currentRow = 1;

            // Extract and write header information from invoice file (above topLeftCell)
            // Use table 1 invoice file if available, otherwise use table 2 invoice file
            let invoiceFileForHeader: SupplierAnalysisFileInfo | null = null;
            if (hasTable1 && this.invoiceFiles.length > 0) {
                invoiceFileForHeader = this.invoiceFiles[0];
            } else if (hasTable2 && this.invoiceFiles2.length > 0) {
                invoiceFileForHeader = this.invoiceFiles2[0];
            }

            if (invoiceFileForHeader) {
                const headerInfo = await this.extractHeaderInfo(invoiceFileForHeader);

                // Write header information rows with styling
                for (let rowIndex = 0; rowIndex < headerInfo.rows.length; rowIndex++) {
                    const rowData = headerInfo.rows[rowIndex];
                    const headerRow = worksheet.getRow(currentRow);

                    for (let col = 0; col < rowData.length; col++) {
                        const cell = headerRow.getCell(col + 1);
                        cell.value = rowData[col];

                        // Apply styling from original file using "row-col" key format
                        const styleKey = `${rowIndex}-${col}`;
                        const originalStyle = headerInfo.styles.get(styleKey);

                        if (originalStyle) {
                            if (originalStyle.font) {
                                // Override font name and size to Cambria 11, but keep other font properties
                                cell.font = { ...originalStyle.font, name: 'Cambria', size: 11 };
                            } else {
                                cell.font = { name: 'Cambria', size: 11 };
                            }
                            if (originalStyle.alignment) {
                                cell.alignment = originalStyle.alignment;
                            }
                            if (originalStyle.border) {
                                cell.border = originalStyle.border;
                            }
                            if (originalStyle.fill) {
                                cell.fill = originalStyle.fill;
                            }
                            if (originalStyle.numFmt) {
                                cell.numFmt = originalStyle.numFmt;
                            }
                            if (originalStyle.protection) {
                                cell.protection = originalStyle.protection;
                            }
                        } else {
                            // Apply Cambria 11 font if no original style
                            cell.font = { name: 'Cambria', size: 11 };
                        }
                    }
                    currentRow++;
                }

                // Add a blank row after header info
                if (headerInfo.rows.length > 0) {
                    currentRow++;
                }
            }

            // Write first table if it exists
            if (hasTable1) {
                // Limit invoice headers to 7 columns (A-G)
                const invoiceHeadersLimited = this.invoiceHeaders.slice(0, 7);

                // Add Section label in column A just above the datatable
                const sectionRow = worksheet.getRow(currentRow);
                const sectionCell = sectionRow.getCell(1);
                sectionCell.value = this.invoiceLabel || '';
                sectionCell.font = { name: 'Cambria', size: 22 };
                currentRow++;

                // Write header row 1 (file names)
                const headerRow1 = worksheet.getRow(currentRow);
                let col = 1;

                // Invoice header (columns A-G) - dark gray background, white bold font (same as column headers)
                for (let i = 0; i < invoiceHeadersLimited.length; i++) {
                    const invoiceHeaderCell = headerRow1.getCell(col + i);
                    if (i === 0) {
                        invoiceHeaderCell.value = 'INVOICE';
                    }
                    invoiceHeaderCell.font = { name: 'Cambria', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
                    invoiceHeaderCell.fill = {
                        type: 'pattern',
                        pattern: 'solid',
                        fgColor: { argb: 'FF808080' } // Dark gray - same as column headers
                    };
                }
                col += invoiceHeadersLimited.length;

                // Track spacing column positions
                const spacingColumns: number[] = [];

                // Skip 2 columns (H and I) - track for setting width to 11 pixels later
                spacingColumns.push(col + 1); // Column H
                spacingColumns.push(col + 2); // Column I
                col += 2;

                // Track Supplier Quotation column start positions for auto-fit
                const supplierQuotationStartCols: number[] = [];

                // Supplier quotation headers (skip 2 columns between each) - orange background, white bold font (same as column headers)
                for (let i = 0; i < this.supplierQuotationFiles.length; i++) {
                    const headerLength = this.getSupplierQuotationHeaderLength(i);
                    supplierQuotationStartCols.push(col); // Track start column for auto-fit

                    // Apply orange background to all cells in this Supplier Quotation group
                    for (let j = 0; j < headerLength; j++) {
                        const supplierHeaderCell = headerRow1.getCell(col + j);
                        if (j === 0) {
                            supplierHeaderCell.value = this.supplierQuotationFiles[i].fileName;
                        }
                        supplierHeaderCell.font = { name: 'Cambria', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
                        supplierHeaderCell.fill = {
                            type: 'pattern',
                            pattern: 'solid',
                            fgColor: { argb: 'FFFFA500' } // Orange - same as column headers
                        };
                    }
                    col += headerLength;
                    // Skip 2 columns before next supplier quotation (except for the last one)
                    if (i < this.supplierQuotationFiles.length - 1) {
                        // Track spacing column positions to set width to 11 pixels later
                        spacingColumns.push(col + 1);
                        spacingColumns.push(col + 2);
                        col += 2;
                    }
                }

                currentRow++;

                // Write header row 2 (column headers)
                const headerRow2 = worksheet.getRow(currentRow);
                col = 1;

                // Invoice headers (columns A-G) - dark gray background, white bold font
                for (const header of invoiceHeadersLimited) {
                    const cell = headerRow2.getCell(col);
                    cell.value = header;
                    cell.font = { name: 'Cambria', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
                    cell.fill = {
                        type: 'pattern',
                        pattern: 'solid',
                        fgColor: { argb: 'FF808080' } // Dark gray
                    };
                    col++;
                }

                // Skip 2 columns (H and I) - spacing columns
                col += 2;

                // Supplier quotation headers - orange background, white bold font
                for (let fileIndex = 0; fileIndex < this.supplierQuotationHeaders.length; fileIndex++) {
                    for (const header of this.supplierQuotationHeaders[fileIndex]) {
                        const cell = headerRow2.getCell(col);
                        cell.value = header;
                        cell.font = { name: 'Cambria', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
                        cell.fill = {
                            type: 'pattern',
                            pattern: 'solid',
                            fgColor: { argb: 'FFFFA500' } // Orange
                        };
                        col++;
                    }
                    // Skip 2 columns before next supplier quotation (except for the last one)
                    if (fileIndex < this.supplierQuotationHeaders.length - 1) {
                        col += 2;
                    }
                }

                currentRow++;

                // Collect Price values for comparison (to find lowest prices)
                const priceColumnMap: Map<number, { fileIndex: number; headerIndex: number }> = new Map();
                let priceCol = 1;

                // Skip invoice columns (A-G)
                priceCol += invoiceHeadersLimited.length;
                // Skip spacing columns (H-I)
                priceCol += 2;

                // Map Price column positions
                for (let fileIndex = 0; fileIndex < this.supplierQuotationHeaders.length; fileIndex++) {
                    for (let headerIndex = 0; headerIndex < this.supplierQuotationHeaders[fileIndex].length; headerIndex++) {
                        const header = this.supplierQuotationHeaders[fileIndex][headerIndex];
                        const headerLower = header.toLowerCase().trim();
                        if (headerLower === 'price' || headerLower.includes('price')) {
                            priceColumnMap.set(priceCol, { fileIndex, headerIndex });
                        }
                        priceCol++;
                    }
                    // Skip spacing columns between Supplier Quotation groups
                    if (fileIndex < this.supplierQuotationHeaders.length - 1) {
                        priceCol += 2;
                    }
                }

                // Write data rows
                for (let rowIndex = 0; rowIndex < this.invoiceData.length; rowIndex++) {
                    const dataRow = worksheet.getRow(currentRow);
                    col = 1;

                    // Collect Price values for this row to find minimum
                    const priceValues: Array<{ col: number; value: number; fileIndex: number; headerIndex: number }> = [];

                    // Invoice data (columns A-G) - Cambria 11 font
                    for (const header of invoiceHeadersLimited) {
                        const cell = dataRow.getCell(col);
                        const value = this.getCellValue(this.invoiceData[rowIndex], header);
                        cell.value = value;
                        cell.font = { name: 'Cambria', size: 11 };

                        // Format Price and Total columns as currency
                        const headerLower = header.toLowerCase().trim();
                        if (headerLower === 'price' || headerLower === 'total' ||
                            headerLower.includes('price') || headerLower.includes('total')) {
                            if (value !== '' && value !== null && value !== undefined) {
                                const numValue = Number(value);
                                if (!isNaN(numValue)) {
                                    cell.value = numValue;
                                    cell.numFmt = '$#,##0.00';
                                }
                            }
                        }

                        col++;
                    }

                    // Skip 2 columns (H and I) - spacing columns (widths set above)
                    col += 2;

                    // Supplier quotation data - Cambria 11 font
                    for (let fileIndex = 0; fileIndex < this.supplierQuotationHeaders.length; fileIndex++) {
                        for (let headerIndex = 0; headerIndex < this.supplierQuotationHeaders[fileIndex].length; headerIndex++) {
                            const header = this.supplierQuotationHeaders[fileIndex][headerIndex];
                            const cell = dataRow.getCell(col);
                            const value = this.getSupplierQuotationValue(fileIndex, rowIndex, header);
                            cell.value = value;
                            cell.font = { name: 'Cambria', size: 11 };

                            // Format Price and Total columns as currency
                            const headerLower = header.toLowerCase().trim();
                            if (headerLower === 'price' || headerLower === 'total' ||
                                headerLower.includes('price') || headerLower.includes('total')) {
                                if (value !== '' && value !== null && value !== undefined) {
                                    const numValue = Number(value);
                                    if (!isNaN(numValue)) {
                                        cell.value = numValue;
                                        cell.numFmt = '$#,##0.00';

                                        // Collect Price values for comparison
                                        if (headerLower === 'price' || headerLower.includes('price')) {
                                            priceValues.push({ col, value: numValue, fileIndex, headerIndex });
                                        }
                                    }
                                }
                            }

                            col++;
                        }
                        // Skip 2 columns before next supplier quotation (except for the last one)
                        if (fileIndex < this.supplierQuotationHeaders.length - 1) {
                            col += 2;
                        }
                    }

                    // Find minimum Price value and highlight
                    if (priceValues.length > 0) {
                        const minPrice = Math.min(...priceValues.map(p => p.value));
                        const minPriceEntries = priceValues.filter(p => p.value === minPrice);

                        // Highlight Price cells
                        for (const priceEntry of priceValues) {
                            const priceCell = dataRow.getCell(priceEntry.col);

                            if (priceEntry.value === minPrice) {
                                if (minPriceEntries.length > 1) {
                                    // Tie - purple background, bold font
                                    priceCell.fill = {
                                        type: 'pattern',
                                        pattern: 'solid',
                                        fgColor: { argb: 'FF800080' } // Purple
                                    };
                                    priceCell.font = { name: 'Cambria', size: 11, bold: true };
                                } else {
                                    // Lowest (no tie) - yellow background, bold font
                                    priceCell.fill = {
                                        type: 'pattern',
                                        pattern: 'solid',
                                        fgColor: { argb: 'FFFFFF00' } // Yellow
                                    };
                                    priceCell.font = { name: 'Cambria', size: 11, bold: true };
                                }
                            }
                        }
                    }

                    currentRow++;
                }

                // Add totals two rows below the data
                const totalRow = currentRow + 2;
                const totalRowObj = worksheet.getRow(totalRow);

                // Find Price and Total column positions for Invoice
                let priceColInvoice = -1;
                let totalColInvoice = -1;
                col = 1;
                for (const header of invoiceHeadersLimited) {
                    const headerLower = header.toLowerCase().trim();
                    if (headerLower === 'price' || headerLower.includes('price')) {
                        priceColInvoice = col;
                    }
                    if (headerLower === 'total' || headerLower.includes('total')) {
                        totalColInvoice = col;
                    }
                    col++;
                }

                // Add "TOTAL USD" below Price column
                if (priceColInvoice > 0) {
                    const totalLabelCell = totalRowObj.getCell(priceColInvoice);
                    totalLabelCell.value = 'TOTAL USD';
                    totalLabelCell.font = { name: 'Cambria', size: 11, bold: true };
                }

                // Calculate sum of all Totals and add below Total column
                if (totalColInvoice > 0) {
                    let totalSum = 0;
                    const firstDataRow = currentRow - this.invoiceData.length;
                    for (let rowIndex = 0; rowIndex < this.invoiceData.length; rowIndex++) {
                        const dataRow = worksheet.getRow(firstDataRow + rowIndex);
                        const totalCell = dataRow.getCell(totalColInvoice);
                        if (totalCell.value !== null && totalCell.value !== undefined && totalCell.value !== '') {
                            const numValue = Number(totalCell.value);
                            if (!isNaN(numValue)) {
                                totalSum += numValue;
                            }
                        }
                    }
                    const totalSumCell = totalRowObj.getCell(totalColInvoice);
                    totalSumCell.value = totalSum;
                    totalSumCell.numFmt = '$#,##0.00';
                    totalSumCell.font = { name: 'Cambria', size: 11, bold: true };
                }

                // Find Price and Total column positions for Supplier Quotations
                col = 1;
                col += invoiceHeadersLimited.length; // Skip invoice columns
                col += 2; // Skip spacing columns H-I

                for (let fileIndex = 0; fileIndex < this.supplierQuotationHeaders.length; fileIndex++) {
                    for (let headerIndex = 0; headerIndex < this.supplierQuotationHeaders[fileIndex].length; headerIndex++) {
                        const header = this.supplierQuotationHeaders[fileIndex][headerIndex];
                        const headerLower = header.toLowerCase().trim();

                        if (headerLower === 'price' || headerLower.includes('price')) {
                            // Add "TOTAL USD" below Price column
                            const totalLabelCell = totalRowObj.getCell(col);
                            totalLabelCell.value = 'TOTAL USD';
                            totalLabelCell.font = { name: 'Cambria', size: 11, bold: true };
                        }

                        if (headerLower === 'total' || headerLower.includes('total')) {
                            // Calculate sum of all Totals for this Supplier Quotation column
                            let totalSum = 0;
                            const firstDataRow = currentRow - this.invoiceData.length;
                            for (let rowIndex = 0; rowIndex < this.invoiceData.length; rowIndex++) {
                                const dataRow = worksheet.getRow(firstDataRow + rowIndex);
                                const totalCell = dataRow.getCell(col);
                                if (totalCell.value !== null && totalCell.value !== undefined && totalCell.value !== '') {
                                    const numValue = Number(totalCell.value);
                                    if (!isNaN(numValue)) {
                                        totalSum += numValue;
                                    }
                                }
                            }
                            const totalSumCell = totalRowObj.getCell(col);
                            totalSumCell.value = totalSum;
                            totalSumCell.numFmt = '$#,##0.00';
                            totalSumCell.font = { name: 'Cambria', size: 11, bold: true };
                        }

                        col++;
                    }
                    // Skip 2 columns before next supplier quotation (except for the last one)
                    if (fileIndex < this.supplierQuotationHeaders.length - 1) {
                        col += 2;
                    }
                }
            }

            // Add 25 empty rows between tables
            currentRow += 25;

            // Write second table if it exists
            if (hasTable2) {
                // Limit invoice headers to 7 columns (A-G)
                const invoiceHeaders2Limited = this.invoiceHeaders2.slice(0, 7);

                // Add Section label in column A just above the datatable
                const sectionRow2 = worksheet.getRow(currentRow);
                const sectionCell2 = sectionRow2.getCell(1);
                sectionCell2.value = this.invoiceLabel2 || '';
                sectionCell2.font = { name: 'Cambria', size: 22 };
                currentRow++;

                // Write header row 1 (file names)
                const headerRow1 = worksheet.getRow(currentRow);
                let col = 1;

                // Invoice header (columns A-G) - dark gray background, white bold font (same as column headers)
                for (let i = 0; i < invoiceHeaders2Limited.length; i++) {
                    const invoiceHeaderCell2 = headerRow1.getCell(col + i);
                    if (i === 0) {
                        invoiceHeaderCell2.value = 'INVOICE';
                    }
                    invoiceHeaderCell2.font = { name: 'Cambria', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
                    invoiceHeaderCell2.fill = {
                        type: 'pattern',
                        pattern: 'solid',
                        fgColor: { argb: 'FF808080' } // Dark gray - same as column headers
                    };
                }
                col += invoiceHeaders2Limited.length;

                // Track spacing column positions for table 2
                const spacingColumns2: number[] = [];

                // Skip 2 columns (H and I) - track for setting width to 11 pixels later
                spacingColumns2.push(col + 1); // Column H
                spacingColumns2.push(col + 2); // Column I
                col += 2;

                // Track Supplier Quotation column start positions for auto-fit
                const supplierQuotationStartCols2: number[] = [];

                // Supplier quotation headers (skip 2 columns between each) - orange background, white bold font (same as column headers)
                for (let i = 0; i < this.supplierQuotationFiles2.length; i++) {
                    const headerLength = this.getSupplierQuotationHeaderLength2(i);
                    supplierQuotationStartCols2.push(col); // Track start column for auto-fit

                    // Apply orange background to all cells in this Supplier Quotation group
                    for (let j = 0; j < headerLength; j++) {
                        const supplierHeaderCell2 = headerRow1.getCell(col + j);
                        if (j === 0) {
                            supplierHeaderCell2.value = this.supplierQuotationFiles2[i].fileName;
                        }
                        supplierHeaderCell2.font = { name: 'Cambria', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
                        supplierHeaderCell2.fill = {
                            type: 'pattern',
                            pattern: 'solid',
                            fgColor: { argb: 'FFFFA500' } // Orange - same as column headers
                        };
                    }
                    col += headerLength;
                    // Skip 2 columns before next supplier quotation (except for the last one)
                    if (i < this.supplierQuotationFiles2.length - 1) {
                        // Track spacing column positions to set width to 11 pixels later
                        spacingColumns2.push(col + 1);
                        spacingColumns2.push(col + 2);
                        col += 2;
                    }
                }

                currentRow++;

                // Write header row 2 (column headers)
                const headerRow2 = worksheet.getRow(currentRow);
                col = 1;

                // Invoice headers (columns A-G) - dark gray background, white bold font
                for (const header of invoiceHeaders2Limited) {
                    const cell = headerRow2.getCell(col);
                    cell.value = header;
                    cell.font = { name: 'Cambria', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
                    cell.fill = {
                        type: 'pattern',
                        pattern: 'solid',
                        fgColor: { argb: 'FF808080' } // Dark gray
                    };
                    col++;
                }

                // Skip 2 columns
                col += 2;

                // Supplier quotation headers - orange background, white bold font
                for (let fileIndex = 0; fileIndex < this.supplierQuotationHeaders2.length; fileIndex++) {
                    for (const header of this.supplierQuotationHeaders2[fileIndex]) {
                        const cell = headerRow2.getCell(col);
                        cell.value = header;
                        cell.font = { name: 'Cambria', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
                        cell.fill = {
                            type: 'pattern',
                            pattern: 'solid',
                            fgColor: { argb: 'FFFFA500' } // Orange
                        };
                        col++;
                    }
                    // Skip 2 columns before next supplier quotation (except for the last one)
                    if (fileIndex < this.supplierQuotationHeaders2.length - 1) {
                        col += 2;
                    }
                }

                currentRow++;

                // Collect Price values for comparison (to find lowest prices) for table 2
                const priceColumnMap2: Map<number, { fileIndex: number; headerIndex: number }> = new Map();
                let priceCol2 = 1;

                // Skip invoice columns (A-G)
                priceCol2 += invoiceHeaders2Limited.length;
                // Skip spacing columns (H-I)
                priceCol2 += 2;

                // Map Price column positions for table 2
                for (let fileIndex = 0; fileIndex < this.supplierQuotationHeaders2.length; fileIndex++) {
                    for (let headerIndex = 0; headerIndex < this.supplierQuotationHeaders2[fileIndex].length; headerIndex++) {
                        const header = this.supplierQuotationHeaders2[fileIndex][headerIndex];
                        const headerLower = header.toLowerCase().trim();
                        if (headerLower === 'price' || headerLower.includes('price')) {
                            priceColumnMap2.set(priceCol2, { fileIndex, headerIndex });
                        }
                        priceCol2++;
                    }
                    // Skip spacing columns between Supplier Quotation groups
                    if (fileIndex < this.supplierQuotationHeaders2.length - 1) {
                        priceCol2 += 2;
                    }
                }

                // Write data rows
                for (let rowIndex = 0; rowIndex < this.invoiceData2.length; rowIndex++) {
                    const dataRow = worksheet.getRow(currentRow);
                    col = 1;

                    // Collect Price values for this row to find minimum
                    const priceValues2: Array<{ col: number; value: number; fileIndex: number; headerIndex: number }> = [];

                    // Invoice data (columns A-G) - Cambria 11 font
                    for (const header of invoiceHeaders2Limited) {
                        const cell = dataRow.getCell(col);
                        const value = this.getCellValue2(this.invoiceData2[rowIndex], header);
                        cell.value = value;
                        cell.font = { name: 'Cambria', size: 11 };

                        // Format Price and Total columns as currency
                        const headerLower = header.toLowerCase().trim();
                        if (headerLower === 'price' || headerLower === 'total' ||
                            headerLower.includes('price') || headerLower.includes('total')) {
                            if (value !== '' && value !== null && value !== undefined) {
                                const numValue = Number(value);
                                if (!isNaN(numValue)) {
                                    cell.value = numValue;
                                    cell.numFmt = '$#,##0.00';
                                }
                            }
                        }

                        col++;
                    }

                    // Skip 2 columns (H and I) - spacing columns (widths set above)
                    col += 2;

                    // Supplier quotation data - Cambria 11 font
                    for (let fileIndex = 0; fileIndex < this.supplierQuotationHeaders2.length; fileIndex++) {
                        for (let headerIndex = 0; headerIndex < this.supplierQuotationHeaders2[fileIndex].length; headerIndex++) {
                            const header = this.supplierQuotationHeaders2[fileIndex][headerIndex];
                            const cell = dataRow.getCell(col);
                            const value = this.getSupplierQuotationValue2(fileIndex, rowIndex, header);
                            cell.value = value;
                            cell.font = { name: 'Cambria', size: 11 };

                            // Format Price and Total columns as currency
                            const headerLower = header.toLowerCase().trim();
                            if (headerLower === 'price' || headerLower === 'total' ||
                                headerLower.includes('price') || headerLower.includes('total')) {
                                if (value !== '' && value !== null && value !== undefined) {
                                    const numValue = Number(value);
                                    if (!isNaN(numValue)) {
                                        cell.value = numValue;
                                        cell.numFmt = '$#,##0.00';

                                        // Collect Price values for comparison
                                        if (headerLower === 'price' || headerLower.includes('price')) {
                                            priceValues2.push({ col, value: numValue, fileIndex, headerIndex });
                                        }
                                    }
                                }
                            }

                            col++;
                        }
                        // Skip 2 columns before next supplier quotation (except for the last one)
                        if (fileIndex < this.supplierQuotationHeaders2.length - 1) {
                            col += 2;
                        }
                    }

                    // Find minimum Price value and highlight for table 2
                    if (priceValues2.length > 0) {
                        const minPrice2 = Math.min(...priceValues2.map(p => p.value));
                        const minPriceEntries2 = priceValues2.filter(p => p.value === minPrice2);

                        // Highlight Price cells
                        for (const priceEntry of priceValues2) {
                            const priceCell = dataRow.getCell(priceEntry.col);

                            if (priceEntry.value === minPrice2) {
                                if (minPriceEntries2.length > 1) {
                                    // Tie - purple background, bold font
                                    priceCell.fill = {
                                        type: 'pattern',
                                        pattern: 'solid',
                                        fgColor: { argb: 'FF800080' } // Purple
                                    };
                                    priceCell.font = { name: 'Cambria', size: 11, bold: true };
                                } else {
                                    // Lowest (no tie) - yellow background, bold font
                                    priceCell.fill = {
                                        type: 'pattern',
                                        pattern: 'solid',
                                        fgColor: { argb: 'FFFFFF00' } // Yellow
                                    };
                                    priceCell.font = { name: 'Cambria', size: 11, bold: true };
                                }
                            }
                        }
                    }

                    currentRow++;
                }

                // Add totals two rows below the data for table 2
                const totalRow2 = currentRow + 2;
                const totalRowObj2 = worksheet.getRow(totalRow2);

                // Find Price and Total column positions for Invoice
                let priceColInvoice2 = -1;
                let totalColInvoice2 = -1;
                col = 1;
                for (const header of invoiceHeaders2Limited) {
                    const headerLower = header.toLowerCase().trim();
                    if (headerLower === 'price' || headerLower.includes('price')) {
                        priceColInvoice2 = col;
                    }
                    if (headerLower === 'total' || headerLower.includes('total')) {
                        totalColInvoice2 = col;
                    }
                    col++;
                }

                // Add "TOTAL USD" below Price column
                if (priceColInvoice2 > 0) {
                    const totalLabelCell = totalRowObj2.getCell(priceColInvoice2);
                    totalLabelCell.value = 'TOTAL USD';
                    totalLabelCell.font = { name: 'Cambria', size: 11, bold: true };
                }

                // Calculate sum of all Totals and add below Total column
                if (totalColInvoice2 > 0) {
                    let totalSum = 0;
                    const firstDataRow2 = currentRow - this.invoiceData2.length;
                    for (let rowIndex = 0; rowIndex < this.invoiceData2.length; rowIndex++) {
                        const dataRow = worksheet.getRow(firstDataRow2 + rowIndex);
                        const totalCell = dataRow.getCell(totalColInvoice2);
                        if (totalCell.value !== null && totalCell.value !== undefined && totalCell.value !== '') {
                            const numValue = Number(totalCell.value);
                            if (!isNaN(numValue)) {
                                totalSum += numValue;
                            }
                        }
                    }
                    const totalSumCell = totalRowObj2.getCell(totalColInvoice2);
                    totalSumCell.value = totalSum;
                    totalSumCell.numFmt = '$#,##0.00';
                    totalSumCell.font = { name: 'Cambria', size: 11, bold: true };
                }

                // Find Price and Total column positions for Supplier Quotations
                col = 1;
                col += invoiceHeaders2Limited.length; // Skip invoice columns
                col += 2; // Skip spacing columns H-I

                for (let fileIndex = 0; fileIndex < this.supplierQuotationHeaders2.length; fileIndex++) {
                    for (let headerIndex = 0; headerIndex < this.supplierQuotationHeaders2[fileIndex].length; headerIndex++) {
                        const header = this.supplierQuotationHeaders2[fileIndex][headerIndex];
                        const headerLower = header.toLowerCase().trim();

                        if (headerLower === 'price' || headerLower.includes('price')) {
                            // Add "TOTAL USD" below Price column
                            const totalLabelCell = totalRowObj2.getCell(col);
                            totalLabelCell.value = 'TOTAL USD';
                            totalLabelCell.font = { name: 'Cambria', size: 11, bold: true };
                        }

                        if (headerLower === 'total' || headerLower.includes('total')) {
                            // Calculate sum of all Totals for this Supplier Quotation column
                            let totalSum = 0;
                            const firstDataRow2 = currentRow - this.invoiceData2.length;
                            for (let rowIndex = 0; rowIndex < this.invoiceData2.length; rowIndex++) {
                                const dataRow = worksheet.getRow(firstDataRow2 + rowIndex);
                                const totalCell = dataRow.getCell(col);
                                if (totalCell.value !== null && totalCell.value !== undefined && totalCell.value !== '') {
                                    const numValue = Number(totalCell.value);
                                    if (!isNaN(numValue)) {
                                        totalSum += numValue;
                                    }
                                }
                            }
                            const totalSumCell = totalRowObj2.getCell(col);
                            totalSumCell.value = totalSum;
                            totalSumCell.numFmt = '$#,##0.00';
                            totalSumCell.font = { name: 'Cambria', size: 11, bold: true };
                        }

                        col++;
                    }
                    // Skip 2 columns before next supplier quotation (except for the last one)
                    if (fileIndex < this.supplierQuotationHeaders2.length - 1) {
                        col += 2;
                    }
                }
            }

            // Set specific column widths for columns A-G
            worksheet.getColumn(1).width = 56 / 7;  // Column A: 56 pixels
            worksheet.getColumn(2).width = 231 / 7; // Column B: 231 pixels
            worksheet.getColumn(3).width = 120 / 7;  // Column C: 120 pixels
            worksheet.getColumn(4).width = 51 / 7;   // Column D: 51 pixels
            worksheet.getColumn(5).width = 87 / 7;   // Column E: 87 pixels
            worksheet.getColumn(6).width = 146 / 7;  // Column F: 146 pixels
            worksheet.getColumn(7).width = 99 / 7;   // Column G: 99 pixels

            // Set spacing columns to 11 pixels (columns BETWEEN Supplier Quotations)
            // Collect all spacing column positions from both tables
            const allSpacingColumns: number[] = [];

            // Table 1 spacing columns: H, I (columns 8, 9), and between Supplier Quotation groups
            allSpacingColumns.push(8, 9); // Columns H and I
            // Add spacing columns between Supplier Quotation groups for table 1
            // Invoice columns are A-G (7 columns), so spacing starts at column 8 (H) and 9 (I)
            let table1Col = 7; // Start at column G (index 7, 1-based = column 7)
            table1Col += 7; // Move past invoice columns A-G (7 columns)
            table1Col += 2; // Move past H and I (spacing columns)
            for (let i = 0; i < this.supplierQuotationFiles.length - 1; i++) {
                const headerLength = this.getSupplierQuotationHeaderLength(i);
                table1Col += headerLength; // Move past this Supplier Quotation group
                allSpacingColumns.push(table1Col + 1); // First spacing column
                allSpacingColumns.push(table1Col + 2); // Second spacing column
                table1Col += 2; // Move past spacing columns
            }

            // Table 2 spacing columns: same column positions as table 1 (both tables use same columns)
            // Since both tables are in the same worksheet, they share the same column structure

            // Set spacing columns to 11 pixels
            allSpacingColumns.forEach(colNum => {
                worksheet.getColumn(colNum).width = 11 / 7;
            });

            // Set default width for Supplier Quotation data columns (auto-fit effect)
            // Only set width for columns that aren't spacing columns
            worksheet.columns.forEach((column, index) => {
                // Column index is 0-based, so column H is index 7
                if (index >= 7) { // Column H and beyond
                    const colNum = index + 1; // Convert to 1-based
                    // Only set width if it's not a spacing column and hasn't been set
                    if (!allSpacingColumns.includes(colNum) && !column.width) {
                        column.width = 15; // Default width for Supplier Quotation columns (auto-fit)
                    }
                }
            });

            // Apply Cambria 11 font to all cells that don't have explicit font settings
            // This ensures the entire page uses Cambria 11
            worksheet.eachRow((row, rowNumber) => {
                row.eachCell((cell) => {
                    if (!cell.font || !cell.font.name) {
                        if (!cell.font) {
                            cell.font = { name: 'Cambria', size: 11 };
                        } else {
                            cell.font = { ...cell.font, name: 'Cambria', size: 11 };
                        }
                    }
                });
            });

            // Generate buffer
            const buffer = await workbook.xlsx.writeBuffer();

            const blob = new Blob([buffer], {
                type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            });

            // Ensure filename has .xlsx extension
            let fileName = this.exportFileName.trim();
            if (!fileName.endsWith('.xlsx')) {
                fileName = `${fileName}.xlsx`;
            }

            saveAs(blob, fileName);

            this.loggingService.logExport('excel_exported', {
                fileName,
                fileSize: blob.size,
                recordCount: this.getRecordCount() + this.getRecordCount2()
            }, 'SupplierAnalysisAnalysisComponent');
        } catch (error) {
            this.loggingService.logError(error as Error, 'excel_export_error', 'SupplierAnalysisAnalysisComponent');
            alert('An error occurred while exporting to Excel. Please try again.');
        }
    }
}

