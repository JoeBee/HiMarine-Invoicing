import { Component, OnInit, OnDestroy } from '@angular/core';
import { CommonModule } from '@angular/common';
import { SupplierAnalysisService, ExcelRowData } from '../../services/supplier-analysis.service';
import { SupplierAnalysisFileInfo } from '../supplier-analysis-inputs/supplier-analysis-inputs.component';
import { LoggingService } from '../../services/logging.service';
import { Subscription } from 'rxjs';

@Component({
    selector: 'app-supplier-analysis-analysis',
    standalone: true,
    imports: [CommonModule],
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
                    this.supplierQuotationData.push(result.rows);
                    this.supplierQuotationHeaders.push(result.headers);
                }

                this.loggingService.logDataProcessing('supplier_analysis_data_loaded', {
                    invoiceFiles: this.invoiceFiles.length,
                    supplierQuotationFiles: this.supplierQuotationFiles.length,
                    rowCount: this.invoiceData.length
                }, 'SupplierAnalysisAnalysisComponent');
            } catch (error) {
                this.loggingService.logError(error as Error, 'data_extraction_error', 'SupplierAnalysisAnalysisComponent');
            }
        }

        this.isLoading = false;
    }

    getAllHeaders(): string[] {
        const allHeaders = new Set<string>(this.invoiceHeaders);
        this.supplierQuotationHeaders.forEach(headers => {
            headers.forEach(header => allHeaders.add(header));
        });
        return Array.from(allHeaders);
    }

    getCellValue(row: ExcelRowData, header: string): any {
        return row[header] !== undefined ? row[header] : '';
    }

    getSupplierQuotationValue(fileIndex: number, rowIndex: number, header: string): any {
        if (fileIndex < this.supplierQuotationData.length && 
            rowIndex < this.supplierQuotationData[fileIndex].length) {
            const row = this.supplierQuotationData[fileIndex][rowIndex];
            return row[header] !== undefined ? row[header] : '';
        }
        return '';
    }

    getSupplierQuotationHeaderLength(index: number): number {
        return this.supplierQuotationHeaders[index]?.length || 1;
    }
}

