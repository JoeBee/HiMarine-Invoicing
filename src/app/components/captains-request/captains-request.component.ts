import { Component } from '@angular/core';
import { CommonModule } from '@angular/common';
import { DataService, ExcelProcessedData, ExcelItemData } from '../../services/data.service';
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

    constructor(private dataService: DataService) { }

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
        // Validate file type
        const validTypes = [
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', // .xlsx
            'application/vnd.ms-excel' // .xls
        ];

        if (!validTypes.includes(file.type) && !file.name.match(/\.(xlsx|xls)$/i)) {
            this.errorMessage = 'Please upload a valid Excel file (.xlsx or .xls)';
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
            console.error('Error processing Excel file:', error);
            this.errorMessage = 'Error processing Excel file. Please ensure it has the required tabs and format.';
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
                    reject(error);
                }
            };

            reader.onerror = () => reject(new Error('Failed to read file'));
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
                            result[tabName] = this.processTabDataWithItems(worksheet);
                        }
                    });

                    resolve(result);
                } catch (error) {
                    reject(error);
                }
            };

            reader.onerror = () => reject(new Error('Failed to read file'));
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

    private processTabDataWithItems(worksheet: XLSX.WorkSheet): { recordsWithTotal: number; sumOfTotals: number; items: ExcelItemData[] } {
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

        let recordsWithTotal = 0;
        let sumOfTotals = 0;
        const items: ExcelItemData[] = [];

        // Skip header row (index 0)
        for (let i = 1; i < jsonData.length; i++) {
            const row = jsonData[i] as any[];

            // Check if row has enough columns (at least 7 for Total column)
            if (row && row.length >= 7) {
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
                        total: totalValue
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
