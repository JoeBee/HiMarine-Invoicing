import { Component } from '@angular/core';
import { CommonModule } from '@angular/common';
import { DataService, ExcelProcessedData, ExcelItemData } from '../../services/data.service';
import { LoggingService } from '../../services/logging.service';
import * as XLSX from 'xlsx';

interface TabData {
    recordsWithTotal: number;
    sumOfTotals: number;
    currency: string;
}

interface ExcelData {
    [tabName: string]: TabData;
}

@Component({
    selector: 'app-captains-order',
    standalone: true,
    imports: [CommonModule],
    templateUrl: './captains-order.component.html',
    styleUrls: ['./captains-order.component.scss']
})
export class CaptainsOrderComponent {
    isDragOver = false;
    uploadedFile: File | null = null;
    isProcessing = false;
    excelData: ExcelData | null = null;
    errorMessage = '';

    constructor(private dataService: DataService, private loggingService: LoggingService) {
    }

    onDragOver(event: DragEvent): void {
        event.preventDefault();
        event.stopPropagation();
        this.isDragOver = true;
    }

    onDragLeave(event: DragEvent): void {
        event.preventDefault();
        event.stopPropagation();
        this.isDragOver = false;
    }

    onDrop(event: DragEvent): void {
        event.preventDefault();
        event.stopPropagation();
        this.isDragOver = false;

        const files = event.dataTransfer?.files;
        if (files && files.length > 0) {
            this.handleFile(files[0]);
        }
    }

    onFileSelected(event: Event): void {
        const input = event.target as HTMLInputElement;
        if (input.files && input.files.length > 0) {
            this.handleFile(input.files[0]);
        }
    }

    private handleFile(file: File): void {
        // Log file upload attempt
        this.loggingService.logFileUpload(file.name, file.size, file.type, 'captains_order', 'CaptainsOrderComponent');

        // Validate file type
        const validTypes = [
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', // .xlsx
            'application/vnd.ms-excel' // .xls
        ];

        if (!validTypes.includes(file.type) && !file.name.match(/\.(xlsx|xls)$/i)) {
            this.errorMessage = 'Please upload a valid Excel file (.xlsx or .xls)';
            this.loggingService.logError(
                `Invalid file type: ${file.type}`,
                'file_validation',
                'CaptainsOrderComponent',
                {
                    fileName: file.name,
                    fileSize: file.size,
                    fileType: file.type,
                    expectedTypes: validTypes
                }
            );
            return;
        }

        this.uploadedFile = file;
        this.errorMessage = '';
        this.processExcelFile(file);
    }

    private async processExcelFile(file: File): Promise<void> {
        this.isProcessing = true;
        this.excelData = null;

        try {
            const data = await this.readExcelFile(file);
            this.excelData = data;

            // Also process with detailed items for invoice
            const detailedData = await this.readExcelFileWithItems(file);
            this.dataService.setExcelData(detailedData);
        } catch (error) {
            this.errorMessage = 'Error processing Excel file. Please ensure it has the required tabs and format.';
            this.loggingService.logError(
                error as Error,
                'excel_file_processing',
                'CaptainsOrderComponent',
                {
                    fileName: file.name,
                    fileSize: file.size,
                    fileType: file.type,
                    processingStep: 'excel_processing',
                    allTabs: true
                }
            );
        } finally {
            this.isProcessing = false;
        }
    }

    private readExcelFile(file: File): Promise<ExcelData> {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();

            reader.onload = (e) => {
                try {
                    const data = e.target?.result;
                    const workbook = XLSX.read(data, { type: 'binary' });

                    const result: ExcelData = {};

                    // Process all tabs from the workbook
                    workbook.SheetNames.forEach(tabName => {
                        const worksheet = workbook.Sheets[tabName];
                        if (worksheet) {
                            result[tabName] = this.processTabData(worksheet);
                        }
                    });

                    resolve(result);
                } catch (error) {
                    this.loggingService.logError(
                        error as Error,
                        'excel_file_reading',
                        'CaptainsOrderComponent',
                        {
                            fileName: file.name,
                            processingStep: 'read_excel_file',
                            allTabs: true
                        }
                    );
                    reject(error);
                }
            };

            reader.onerror = () => {
                const error = new Error('Failed to read file');
                this.loggingService.logError(
                    error,
                    'file_reader_error',
                    'CaptainsOrderComponent',
                    {
                        fileName: file.name,
                        fileSize: file.size,
                        fileType: file.type,
                        readerType: 'FileReader'
                    }
                );
                reject(error);
            };
            reader.readAsBinaryString(file);
        });
    }

    private readExcelFileWithItems(file: File): Promise<ExcelProcessedData> {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();

            reader.onload = (e) => {
                try {
                    const data = e.target?.result;
                    const workbook = XLSX.read(data, { type: 'binary' });

                    const result: ExcelProcessedData = {};

                    // Process all tabs from the workbook
                    workbook.SheetNames.forEach(tabName => {
                        const worksheet = workbook.Sheets[tabName];
                        if (worksheet) {
                            result[tabName] = this.processTabDataWithItems(worksheet, tabName);
                        }
                    });

                    resolve(result);
                } catch (error) {
                    this.loggingService.logError(
                        error as Error,
                        'excel_file_reading_with_items',
                        'CaptainsOrderComponent',
                        {
                            fileName: file.name,
                            processingStep: 'read_excel_file_with_items',
                            allTabs: true
                        }
                    );
                    reject(error);
                }
            };

            reader.onerror = () => {
                const error = new Error('Failed to read file with items');
                this.loggingService.logError(
                    error,
                    'file_reader_error_with_items',
                    'CaptainsOrderComponent',
                    {
                        fileName: file.name,
                        fileSize: file.size,
                        fileType: file.type,
                        readerType: 'FileReader',
                        processingMode: 'with_items'
                    }
                );
                reject(error);
            };
            reader.readAsBinaryString(file);
        });
    }

    private detectCurrency(value: string): string | null {
        if (!value || typeof value !== 'string') return null;
        const str = value.trim();
        const strUpper = str.toUpperCase();
        
        // Check for specific currency prefixes first (most specific to least specific)
        // Check both original case and uppercase to catch NZ$, nz$, NZ$ etc.
        if (strUpper.includes('NZ$') || strUpper.includes('NZD')) return 'NZ$';
        if (strUpper.includes('A$') || strUpper.includes('AUD')) return 'A$';
        if (strUpper.includes('C$') || strUpper.includes('CAD')) return 'C$';
        if (str.includes('€') || strUpper.includes('EUR')) return '€';
        if (str.includes('£') || strUpper.includes('GBP')) return '£';
        // Check for generic $ last (only if not already matched by NZ$, A$, C$)
        // We need to check that $ is not part of NZ$, A$, or C$
        if (str.includes('$')) {
            // Make sure it's not NZ$, A$, or C$
            if (!strUpper.includes('NZ$') && !strUpper.includes('A$') && !strUpper.includes('C$')) {
                return '$';
            }
        }
        if (strUpper.includes('USD')) return '$';
        
        return null;
    }

    private processTabData(worksheet: XLSX.WorkSheet): TabData {
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: false });

        let recordsWithTotal = 0;
        let sumOfTotals = 0;

        // Detect currency from the first price or total value found
        let detectedCurrency = ''; // Default to empty string
        for (let i = 1; i < jsonData.length; i++) {
            const row = jsonData[i] as any[];
            if (row && row.length >= 7) {
                // Check price column (index 5)
                const priceValue = row[5];
                if (priceValue) {
                    const currency = this.detectCurrency(String(priceValue));
                    if (currency) {
                        detectedCurrency = currency;
                        break;
                    }
                }

                // Also check total column (index 6) if price didn't have currency
                if (!detectedCurrency) {
                    const totalValue = row[6];
                    if (totalValue) {
                        const currency = this.detectCurrency(String(totalValue));
                        if (currency) {
                            detectedCurrency = currency;
                            break;
                        }
                    }
                }
            }
        }

        // Skip header row (index 0)
        for (let i = 1; i < jsonData.length; i++) {
            const row = jsonData[i] as any[];

            // Check if row has enough columns (at least 7 for Total column)
            if (row && row.length >= 7) {
                // Skip rows without a description (column B, index 1)
                const descriptionCell = String(row[1] ?? '').trim();
                if (!descriptionCell) {
                    continue;
                }
                const totalValue = this.parseNumericValue(row[6]); // Total is column G (index 6)

                if (totalValue > 0) {
                    recordsWithTotal++;
                    sumOfTotals += totalValue;
                }
            }
        }

        return {
            recordsWithTotal,
            sumOfTotals,
            currency: detectedCurrency
        };
    }

    private processTabDataWithItems(worksheet: XLSX.WorkSheet, tabName: string): { recordsWithTotal: number; sumOfTotals: number; items: ExcelItemData[] } {
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: false });

        let recordsWithTotal = 0;
        let sumOfTotals = 0;
        const items: ExcelItemData[] = [];

        // Detect currency from the first price or total value found
        let detectedCurrency = ''; // Default to empty string
        for (let i = 1; i < jsonData.length; i++) {
            const row = jsonData[i] as any[];
            if (row && row.length >= 7) {
                // Check price column (index 5)
                const priceValue = row[5];
                if (priceValue) {
                    const currency = this.detectCurrency(String(priceValue));
                    if (currency) {
                        detectedCurrency = currency;
                        console.log(`Detected currency for tab ${tabName} from price: ${currency}`);
                        break;
                    }
                }

                // Also check total column (index 6) if price didn't have currency
                if (!detectedCurrency) {
                    const totalValue = row[6];
                    if (totalValue) {
                        const currency = this.detectCurrency(String(totalValue));
                        if (currency) {
                            detectedCurrency = currency;
                            console.log(`Detected currency for tab ${tabName} from total: ${currency}`);
                            break;
                        }
                    }
                }
            }
        }

        console.log(`Detected currency for tab ${tabName}: ${detectedCurrency}`);

        // Skip header row (index 0)
        for (let i = 1; i < jsonData.length; i++) {
            const row = jsonData[i] as any[];

            // Check if row has enough columns (at least 7 for Total column)
            if (row && row.length >= 7) {
                // Skip rows without a description (column B, index 1)
                const descriptionCell = String(row[1] ?? '').trim();
                if (!descriptionCell) {
                    continue;
                }
                const totalValue = this.parseNumericValue(row[6]); // Total is column G (index 6)

                if (totalValue > 0) {
                    recordsWithTotal++;
                    sumOfTotals += totalValue;

                    // Extract item data (assuming columns: Pos, Description, Remark, Unit, Qty, Price, Total)
                    const item: ExcelItemData = {
                        pos: this.parseNumericValue(row[0]) || (recordsWithTotal),
                        description: String(row[1] || ''),
                        remark: String(row[2] || ''),
                        unit: String(row[3] || 'EACH'),
                        qty: this.parseNumericValue(row[4]) || 1,
                        price: this.parseNumericValue(row[5]) || 0,
                        total: totalValue,
                        tabName: tabName,
                        currency: detectedCurrency
                    };

                    items.push(item);
                }
            }
        }

        return {
            recordsWithTotal,
            sumOfTotals,
            items
        };
    }

    private parseNumericValue(value: any): number {
        if (typeof value === 'number') {
            return value;
        }

        if (typeof value === 'string') {
            // Remove currency prefixes first (most specific to least specific), then other symbols
            let cleaned = value.trim();
            cleaned = cleaned.replace(/NZ\$/gi, '');
            cleaned = cleaned.replace(/A\$/gi, '');
            cleaned = cleaned.replace(/C\$/gi, '');
            cleaned = cleaned.replace(/[€£$,]/g, '');
            const parsed = parseFloat(cleaned);
            return isNaN(parsed) ? 0 : parsed;
        }

        return 0;
    }

    getSummaryTabs(): string[] {
        if (!this.excelData) return [];
        // Get all tabs except "COVER SHEET"
        return Object.keys(this.excelData).filter(tab => tab !== 'COVER SHEET');
    }

    getTotalRecords(): number {
        if (!this.excelData) return 0;

        const tabs = this.getSummaryTabs();
        return tabs.reduce((total, tab) => {
            return total + (this.excelData![tab]?.recordsWithTotal || 0);
        }, 0);
    }

    getGrandTotal(): number {
        if (!this.excelData) return 0;

        const tabs = this.getSummaryTabs();
        return tabs.reduce((total, tab) => {
            return total + (this.excelData![tab]?.sumOfTotals || 0);
        }, 0);
    }

    getPrimaryCurrency(): string {
        if (!this.excelData) return '';

        const tabs = this.getSummaryTabs();
        const currencyCount: { [key: string]: number } = {};
        
        tabs.forEach(tab => {
            // Only count currency from tabs that exist and have currency defined
            if (this.excelData![tab] && this.excelData![tab].currency) {
                const currency = this.excelData![tab].currency;
                currencyCount[currency] = (currencyCount[currency] || 0) + 1;
            }
        });

        // If no currencies found, default to empty string
        const currencies = Object.keys(currencyCount);
        if (currencies.length === 0) return '';

        // Find the most common currency
        let maxCount = 0;
        let mostCommonCurrency = currencies[0];
        for (const [currency, count] of Object.entries(currencyCount)) {
            if (count > maxCount) {
                maxCount = count;
                mostCommonCurrency = currency;
            }
        }

        return mostCommonCurrency;
    }

    getTabCurrency(tab: string): string {
        if (!this.excelData || !this.excelData[tab]) return '';
        return this.excelData[tab].currency || '';
    }
}

