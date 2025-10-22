import { Injectable } from '@angular/core';
import { BehaviorSubject, Observable } from 'rxjs';
import * as XLSX from 'xlsx';

export interface SupplierFileInfo {
    fileName: string;
    topLeftCell: string;
    descriptionColumn: string;
    priceColumn: string;
    unitColumn: string;
    descriptionHeader: string;
    priceHeader: string;
    unitHeader: string;
    rowCount: number;
    file: File;
}

export interface ProcessedDataRow {
    fileName: string;
    description: string;
    price: number;
    unit: string;
    count: number;
    originalData?: any[]; // Store original row data from XLSX
    originalHeaders?: string[]; // Store original column headers
}

@Injectable({
    providedIn: 'root'
})
export class DataService {
    private supplierFilesSubject = new BehaviorSubject<SupplierFileInfo[]>([]);
    private processedDataSubject = new BehaviorSubject<ProcessedDataRow[]>([]);

    supplierFiles$: Observable<SupplierFileInfo[]> = this.supplierFilesSubject.asObservable();
    processedData$: Observable<ProcessedDataRow[]> = this.processedDataSubject.asObservable();

    constructor() { }

    addSupplierFiles(files: File[]): Promise<void> {
        return new Promise(async (resolve) => {
            const currentFiles = this.supplierFilesSubject.value;
            const newFileInfos: SupplierFileInfo[] = [];

            for (const file of files) {
                const fileInfo = await this.analyzeFile(file);
                newFileInfos.push(fileInfo);
            }

            this.supplierFilesSubject.next([...currentFiles, ...newFileInfos]);
            resolve();
        });
    }

    private async analyzeFile(file: File): Promise<SupplierFileInfo> {
        return new Promise((resolve) => {
            const reader = new FileReader();

            reader.onload = (e: any) => {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];

                // Find the top-left cell of the data table
                const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1:Z100');
                let topLeftCell = 'NOT FOUND';
                let descriptionColumn = 'NOT FOUND';
                let priceColumn = 'NOT FOUND';
                let unitColumn = 'NOT FOUND';
                let descriptionHeader = '';
                let priceHeader = '';
                let unitHeader = '';
                let rowCount = 0;

                // Look for the first row containing "description", "descrption", or "item"
                for (let row = range.s.r; row <= range.e.r; row++) {
                    let foundHeaderRow = false;

                    // Check all cells in this row for header keywords
                    for (let col = range.s.c; col <= range.e.c; col++) {
                        const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
                        const cell = worksheet[cellAddress];

                        if (cell && cell.v) {
                            const cellValue = String(cell.v).toLowerCase();


                            if (cellValue.length > 25) {
                                console.log("cellValue", cellValue);
                                // skip rows with less than 25 characters
                                // continue;
                            }
                            else if (cellValue.includes('description') ||
                                cellValue.includes('descrption') ||
                                cellValue.includes('product') ||
                                cellValue.includes('item')) {

                                console.log("cellValue", cellValue);

                                // This is our header row - find the top left corner
                                topLeftCell = XLSX.utils.encode_cell({ r: row, c: range.s.c });
                                foundHeaderRow = true;

                                // Look for description and price columns in this header row
                                for (let searchCol = range.s.c; searchCol <= range.e.c; searchCol++) {
                                    const headerAddress = XLSX.utils.encode_cell({ r: row, c: searchCol });
                                    const headerCell = worksheet[headerAddress];

                                    if (headerCell && headerCell.v) {
                                        const headerValue = String(headerCell.v)
                                            .toLowerCase().trim();
                                        const headerText = String(headerCell.v);


                                        if (headerValue.includes('description') || headerValue.includes('descrption') ||
                                            headerValue.includes('item') || headerValue.includes('product') ||
                                            headerValue.includes('name')) {
                                            descriptionColumn = XLSX.utils.encode_col(searchCol);
                                            descriptionHeader = headerText;
                                        }

                                        if (headerValue.includes('price') || headerValue.includes('cost') ||
                                            headerValue.includes('amount') || headerValue.includes('value')) {
                                            priceColumn = XLSX.utils.encode_col(searchCol);
                                            priceHeader = headerText;
                                        }

                                        if (headerValue === 'unit' ||
                                            headerValue === 'units') {
                                            unitColumn = XLSX.utils.encode_col(searchCol);
                                            unitHeader = headerText;
                                        }
                                    }
                                }

                                // If columns were not found, set them to "NOT FOUND"
                                if (descriptionColumn === 'NOT FOUND') {
                                    descriptionHeader = 'NOT FOUND';
                                }
                                if (priceColumn === 'NOT FOUND') {
                                    priceHeader = 'NOT FOUND';
                                }
                                if (unitColumn === 'NOT FOUND') {
                                    unitHeader = 'NOT FOUND';
                                }

                                // Count data rows (excluding header row) - only if columns were found
                                if (descriptionColumn !== 'NOT FOUND' && priceColumn !== 'NOT FOUND') {
                                    const descColIndex = XLSX.utils.decode_col(descriptionColumn);
                                    const priceColIndex = XLSX.utils.decode_col(priceColumn);

                                    for (let dataRow = row + 1; dataRow <= range.e.r; dataRow++) {
                                        const descAddress = XLSX.utils.encode_cell({ r: dataRow, c: descColIndex });
                                        const priceAddress = XLSX.utils.encode_cell({ r: dataRow, c: priceColIndex });

                                        const descCell = worksheet[descAddress];
                                        const priceCell = worksheet[priceAddress];

                                        // Count rows that have both description and price data
                                        if (descCell && descCell.v && priceCell && priceCell.v) {
                                            rowCount++;
                                        }
                                    }
                                }

                                break;
                            }
                        }
                    }

                    if (foundHeaderRow) {
                        break;
                    }
                }

                const fileName = file.name.replace(/\.[^/.]+$/, ''); // Remove extension

                resolve({
                    fileName,
                    topLeftCell,
                    descriptionColumn,
                    priceColumn,
                    unitColumn,
                    descriptionHeader,
                    priceHeader,
                    unitHeader,
                    rowCount,
                    file
                });
            };

            reader.readAsArrayBuffer(file);
        });
    }

    async processSupplierFiles(): Promise<void> {
        const supplierFiles = this.supplierFilesSubject.value;
        const allData: ProcessedDataRow[] = [];

        for (const fileInfo of supplierFiles) {
            const rowData = await this.extractDataFromFile(fileInfo);
            allData.push(...rowData);
        }

        // Sort by description, then by price
        allData.sort((a, b) => {
            const descCompare = a.description.localeCompare(b.description);
            if (descCompare !== 0) return descCompare;
            return a.price - b.price;
        });

        this.processedDataSubject.next(allData);
    }

    private async extractDataFromFile(fileInfo: SupplierFileInfo): Promise<ProcessedDataRow[]> {
        return new Promise((resolve) => {
            const reader = new FileReader();

            reader.onload = (e: any) => {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];

                const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1:Z100');
                const topLeftCellRef = XLSX.utils.decode_cell(fileInfo.topLeftCell);
                const descColIndex = XLSX.utils.decode_col(fileInfo.descriptionColumn);
                const priceColIndex = XLSX.utils.decode_col(fileInfo.priceColumn);
                const unitColIndex = XLSX.utils.decode_col(fileInfo.unitColumn);

                const rows: ProcessedDataRow[] = [];

                // Extract original headers from the header row
                const originalHeaders: string[] = [];
                for (let col = range.s.c; col <= range.e.c; col++) {
                    const headerAddress = XLSX.utils.encode_cell({ r: topLeftCellRef.r, c: col });
                    const headerCell = worksheet[headerAddress];
                    originalHeaders.push(headerCell && headerCell.v ? String(headerCell.v) : '');
                }

                // Start from the row after the header
                for (let row = topLeftCellRef.r + 1; row <= range.e.r; row++) {
                    const descAddress = XLSX.utils.encode_cell({ r: row, c: descColIndex });
                    const priceAddress = XLSX.utils.encode_cell({ r: row, c: priceColIndex });
                    const unitAddress = XLSX.utils.encode_cell({ r: row, c: unitColIndex });

                    const descCell = worksheet[descAddress];
                    const priceCell = worksheet[priceAddress];
                    const unitCell = worksheet[unitAddress];

                    if (descCell && descCell.v && priceCell && priceCell.v) {
                        // Extract original row data
                        const originalRowData: any[] = [];
                        for (let col = range.s.c; col <= range.e.c; col++) {
                            const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
                            const cell = worksheet[cellAddress];
                            originalRowData.push(cell && cell.v ? cell.v : '');
                        }

                        rows.push({
                            fileName: fileInfo.fileName,
                            description: String(descCell.v),
                            price: Number(priceCell.v),
                            unit: unitCell && unitCell.v ? String(unitCell.v) : '',
                            count: 0,
                            originalData: originalRowData,
                            originalHeaders: originalHeaders
                        });
                    }
                }

                resolve(rows);
            };

            reader.readAsArrayBuffer(fileInfo.file);
        });
    }

    getProcessedData(): ProcessedDataRow[] {
        return this.processedDataSubject.value;
    }

    updateRowCount(index: number, count: number): void {
        const currentData = this.processedDataSubject.value;
        if (index >= 0 && index < currentData.length) {
            currentData[index].count = count;
            this.processedDataSubject.next([...currentData]);
        }
    }

    hasSupplierFiles(): boolean {
        return this.supplierFilesSubject.value.length > 0;
    }

    clearAll(): void {
        this.supplierFilesSubject.next([]);
        this.processedDataSubject.next([]);
    }
}

