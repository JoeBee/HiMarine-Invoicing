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
    private filesSubscription?: Subscription;

    constructor(
        private supplierAnalysisService: SupplierAnalysisService,
        private loggingService: LoggingService
    ) { }

    async ngOnInit(): Promise<void> {
        await this.loadData();
        
        // Subscribe to file changes
        this.filesSubscription = this.supplierAnalysisService.files$.subscribe(async () => {
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
                    this.invoiceHeaders = invoiceResult.headers;
                }

                // Extract supplier quotation data (all files)
                this.supplierQuotationData = [];
                this.supplierQuotationHeaders = [];
                for (const file of this.supplierQuotationFiles) {
                    const result = await this.supplierAnalysisService.extractDataFromFile(file);
                    // Filter out "POS." columns from headers
                    const filteredHeaders = result.headers.filter(header => 
                        !header.trim().toUpperCase().startsWith('POS.')
                    );
                    this.supplierQuotationHeaders.push(filteredHeaders);
                    // Filter out "POS." columns from data rows
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
            } catch (error) {
                this.loggingService.logError(error as Error, 'data_extraction_error', 'SupplierAnalysisAnalysisComponent');
            }
        }

        this.isLoading = false;
    }

    updateExportFileName(): void {
        if (this.invoiceFiles.length > 0) {
            const invoiceFileName = this.invoiceFiles[0].fileName;
            // Remove extension if present
            const nameWithoutExt = invoiceFileName.replace(/\.(xlsx|xls|xlsm)$/i, '');
            this.exportFileName = `_OUTPUT ${nameWithoutExt}`;
        } else {
            this.exportFileName = '_OUTPUT Invoice';
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
        const value = row[header] !== undefined ? row[header] : '';
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

    async exportToExcel(): Promise<void> {
        if (this.invoiceFiles.length === 0) {
            alert('No invoice file available to export.');
            return;
        }

        this.loggingService.logButtonClick('export_to_excel', 'SupplierAnalysisAnalysisComponent', {
            fileName: this.exportFileName,
            recordCount: this.getRecordCount(),
            fileCount: this.getFileCount()
        });

        try {
            const invoiceFile = this.invoiceFiles[0];
            
            // Read the invoice file
            const fileReader = new FileReader();
            
            fileReader.onload = async (e: any) => {
                try {
                    const arrayBuffer = e.target.result;
                    
                    // Use JSZip to directly modify the Excel file XML (preserves file structure)
                    // This approach avoids ExcelJS modifications that can cause corruption
                    const zip = await JSZip.loadAsync(arrayBuffer);
                    const worksheetFiles = Object.keys(zip.files).filter(name =>
                        name.startsWith('xl/worksheets/sheet') && name.endsWith('.xml')
                    );
                    
                    // Process each worksheet to remove print area and page view settings
                    for (const worksheetFile of worksheetFiles) {
                        const worksheetXml = await zip.file(worksheetFile)?.async('string');
                        if (!worksheetXml) {
                            continue;
                        }
                        
                        let modifiedXml = worksheetXml;
                        
                        // Remove printArea from pageSetup
                        modifiedXml = modifiedXml.replace(/<pageSetup[^>]*printArea="[^"]*"[^>]*>/g, (match) => {
                            return match.replace(/\s+printArea="[^"]*"/g, '');
                        });
                        
                        // Remove fitToPage, fitToWidth, fitToHeight from pageSetup
                        modifiedXml = modifiedXml.replace(/<pageSetup[^>]*>/g, (match) => {
                            return match
                                .replace(/\s+fitToPage="[^"]*"/g, '')
                                .replace(/\s+fitToWidth="[^"]*"/g, '')
                                .replace(/\s+fitToHeight="[^"]*"/g, '');
                        });
                        
                        // Remove page break preview view from sheetViews
                        if (modifiedXml.includes('<sheetViews>')) {
                            modifiedXml = modifiedXml.replace(
                                /<sheetView([^>]*?)(\s*\/?>)/g,
                                (match, attrs, closing) => {
                                    // Remove view attribute but keep other attributes
                                    const cleanAttrs = attrs.replace(/\s*view="[^"]*"/g, '');
                                    return `<sheetView${cleanAttrs}${closing}`;
                                }
                            );
                        }
                        
                        zip.file(worksheetFile, modifiedXml);
                    }
                    
                    // Generate buffer with proper compression (same as invoice-workbook-builder)
                    const buffer = await zip.generateAsync({ type: 'arraybuffer' });
                    
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
                        recordCount: this.getRecordCount()
                    }, 'SupplierAnalysisAnalysisComponent');
                } catch (error) {
                    this.loggingService.logError(error as Error, 'excel_export_error', 'SupplierAnalysisAnalysisComponent');
                    alert('An error occurred while exporting to Excel. Please try again.');
                }
            };
            
            fileReader.onerror = () => {
                this.loggingService.logError(new Error('File reading error'), 'excel_export_error', 'SupplierAnalysisAnalysisComponent');
                alert('An error occurred while reading the invoice file.');
            };
            
            fileReader.readAsArrayBuffer(invoiceFile.file);
        } catch (error) {
            this.loggingService.logError(error as Error, 'excel_export_error', 'SupplierAnalysisAnalysisComponent');
            alert('An error occurred while exporting to Excel. Please try again.');
        }
    }
}

