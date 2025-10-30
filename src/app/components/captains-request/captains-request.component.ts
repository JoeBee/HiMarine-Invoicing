import { Component } from '@angular/core';
import { CommonModule } from '@angular/common';
import { DataService, ExcelProcessedData, ExcelItemData } from '../../services/data.service';
import { LoggingService } from '../../services/logging.service';
import * as XLSX from 'xlsx';

interface TabData {
    recordsWithTotal: number;
    sumOfTotals: number;
}

interface ExcelData {
    [tabName: string]: TabData;
}

@Component({
    selector: 'app-captains-request',
    standalone: true,
    imports: [CommonModule],
    templateUrl: './captains-request.component.html',
    styleUrls: ['./captains-request.component.scss']
})
export class CaptainsRequestComponent {
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
        this.loggingService.logFileUpload(file.name, file.size, file.type, 'captains_request', 'CaptainsRequestComponent');

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
                'CaptainsRequestComponent',
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
                'CaptainsRequestComponent',
                {
                    fileName: file.name,
                    fileSize: file.size,
                    fileType: file.type,
                    processingStep: 'excel_processing',
                    expectedTabs: ['COVER SHEET', 'PROVISIONS', 'FRESH PROVISIONS', 'BOND']
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
                    const requiredTabs = ['COVER SHEET', 'PROVISIONS', 'FRESH PROVISIONS', 'BOND'];

                    // Process each required tab
                    requiredTabs.forEach(tabName => {
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
                        'CaptainsRequestComponent',
                        {
                            fileName: file.name,
                            processingStep: 'read_excel_file',
                            requiredTabs: ['COVER SHEET', 'PROVISIONS', 'FRESH PROVISIONS', 'BOND']
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
                    'CaptainsRequestComponent',
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
                    const requiredTabs = ['COVER SHEET', 'PROVISIONS', 'FRESH PROVISIONS', 'BOND'];

                    // Process each required tab
                    requiredTabs.forEach(tabName => {
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
                        'CaptainsRequestComponent',
                        {
                            fileName: file.name,
                            processingStep: 'read_excel_file_with_items',
                            requiredTabs: ['COVER SHEET', 'PROVISIONS', 'FRESH PROVISIONS', 'BOND']
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
                    'CaptainsRequestComponent',
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

    private processTabData(worksheet: XLSX.WorkSheet): TabData {
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

        let recordsWithTotal = 0;
        let sumOfTotals = 0;

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
            sumOfTotals
        };
    }

    private processTabDataWithItems(worksheet: XLSX.WorkSheet, tabName: string): { recordsWithTotal: number; sumOfTotals: number; items: ExcelItemData[] } {
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

        let recordsWithTotal = 0;
        let sumOfTotals = 0;
        const items: ExcelItemData[] = [];

        // Detect currency from the first price or total value found
        let detectedCurrency = '£'; // Default to GBP
        for (let i = 1; i < jsonData.length; i++) {
            const row = jsonData[i] as any[];
            if (row && row.length >= 7) {
                // Check price column (index 5)
                const priceValue = row[5];
                if (priceValue && typeof priceValue === 'string') {
                    const priceStr = String(priceValue).trim();
                    console.log(`Price value: "${priceStr}"`);
                    if (priceStr.includes('$')) {
                        detectedCurrency = '$';
                        break;
                    } else if (priceStr.includes('€')) {
                        detectedCurrency = '€';
                        break;
                    } else if (priceStr.includes('£')) {
                        detectedCurrency = '£';
                        break;
                    }
                }

                // Also check total column (index 6) if price didn't have currency
                if (detectedCurrency === '£') {
                    const totalValue = row[6];
                    if (totalValue && typeof totalValue === 'string') {
                        const totalStr = String(totalValue).trim();
                        if (totalStr.includes('$')) {
                            detectedCurrency = '$';
                            break;
                        } else if (totalStr.includes('€')) {
                            detectedCurrency = '€';
                            break;
                        } else if (totalStr.includes('£')) {
                            detectedCurrency = '£';
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
            // Remove currency symbols and parse
            const cleaned = value.replace(/[$,]/g, '').trim();
            const parsed = parseFloat(cleaned);
            return isNaN(parsed) ? 0 : parsed;
        }

        return 0;
    }

    getTotalRecords(): number {
        if (!this.excelData) return 0;

        const tabs = ['PROVISIONS', 'FRESH PROVISIONS', 'BOND'];
        return tabs.reduce((total, tab) => {
            return total + (this.excelData![tab]?.recordsWithTotal || 0);
        }, 0);
    }

    getGrandTotal(): number {
        if (!this.excelData) return 0;

        const tabs = ['PROVISIONS', 'FRESH PROVISIONS', 'BOND'];
        return tabs.reduce((total, tab) => {
            return total + (this.excelData![tab]?.sumOfTotals || 0);
        }, 0);
    }
}
