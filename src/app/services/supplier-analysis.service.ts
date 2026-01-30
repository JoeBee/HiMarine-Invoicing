import { Injectable } from '@angular/core';
import { BehaviorSubject, Observable } from 'rxjs';
import * as XLSX from 'xlsx';

export interface SupplierAnalysisFileInfo {
    fileName: string;
    file: File | null;
    rowCount: number;
    topLeftCell: string;
    category: 'Invoice' | 'Supplier Quotations';
    discount: number;
    isBlank?: boolean;
}

export interface SupplierAnalysisFileSet {
    id: number;
    files: SupplierAnalysisFileInfo[];
}

export interface ExcelRowData {
    [key: string]: any;
}

export interface SupplierAnalysisData {
    invoiceData: ExcelRowData[];
    supplierQuotationData: ExcelRowData[][];
    invoiceHeaders: string[];
    supplierQuotationHeaders: string[][];
}

@Injectable({
    providedIn: 'root'
})
export class SupplierAnalysisService {
    private fileSetsSubject = new BehaviorSubject<SupplierAnalysisFileSet[]>([]);
    fileSets$: Observable<SupplierAnalysisFileSet[]> = this.fileSetsSubject.asObservable();

    setFileSets(fileSets: SupplierAnalysisFileSet[]): void {
        this.fileSetsSubject.next(fileSets);
    }

    getFileSets(): SupplierAnalysisFileSet[] {
        return this.fileSetsSubject.value;
    }
    
    updateFileSet(id: number, files: SupplierAnalysisFileInfo[]): void {
        const currentSets = this.getFileSets();
        const existingSetIndex = currentSets.findIndex(s => s.id === id);
        
        let newSets: SupplierAnalysisFileSet[];
        
        if (existingSetIndex !== -1) {
            newSets = [...currentSets];
            newSets[existingSetIndex] = { ...newSets[existingSetIndex], files };
        } else {
            newSets = [...currentSets, { id, files }];
        }
        
        this.fileSetsSubject.next(newSets);
    }
    
    removeFileSet(id: number): void {
        const currentSets = this.getFileSets();
        const newSets = currentSets.filter(s => s.id !== id);
        this.fileSetsSubject.next(newSets);
    }

    // Deprecated methods for backward compatibility during refactor (or removal if I update everything at once)
    // I will remove them and fix the components immediately.

    async extractDataFromFile(fileInfo: SupplierAnalysisFileInfo): Promise<{ headers: string[]; rows: ExcelRowData[] }> {
        // Handle blank files
        if (fileInfo.isBlank || !fileInfo.file) {
            const blankHeaders = ['Description', 'Price', 'Total'];
            const blankRows: ExcelRowData[] = [];
            for (let i = 0; i < fileInfo.rowCount; i++) {
                const blankRow: ExcelRowData = {};
                blankHeaders.forEach(header => blankRow[header] = '');
                blankRows.push(blankRow);
            }
            return Promise.resolve({ headers: blankHeaders, rows: blankRows });
        }

        return new Promise((resolve, reject) => {
            const reader = new FileReader();

            reader.onload = (e: any) => {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const firstSheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[firstSheetName];

                    const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1:Z100');
                    let topLeftCellRef;
                    try {
                        topLeftCellRef = XLSX.utils.decode_cell(fileInfo.topLeftCell);
                    } catch (e) {
                         // Fallback if cell is invalid
                         topLeftCellRef = { c: 0, r: 0 };
                    }

                    // Extract headers
                    const headers: string[] = [];
                    for (let col = range.s.c; col <= range.e.c; col++) {
                        const headerAddress = XLSX.utils.encode_cell({ r: topLeftCellRef.r, c: col });
                        const headerCell = worksheet[headerAddress];
                        headers.push(headerCell && headerCell.v ? String(headerCell.v) : `Column ${col + 1}`);
                    }

                    // Extract data rows
                    const rows: ExcelRowData[] = [];
                    for (let row = topLeftCellRef.r + 1; row <= range.e.r; row++) {
                        // Check if row has data
                        let hasData = false;
                        const rowData: ExcelRowData = {};

                        for (let col = range.s.c; col <= range.e.c; col++) {
                            const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
                            const cell = worksheet[cellAddress];
                            const value = cell && cell.v !== null && cell.v !== undefined ? cell.v : '';
                            
                            if (value !== '' && String(value).trim() !== '') {
                                hasData = true;
                            }
                            
                            const headerName = headers[col - range.s.c];
                            let finalValue = value;

                            const headerLower = headerName.toLowerCase().trim();
                            const isPriceColumn = headerLower === 'price' || headerLower.includes('price');

                            // Always capture Gross Price when it's a price column
                            if (isPriceColumn) {
                                let grossPrice = value;
                                if (typeof value !== 'number' && !isNaN(Number(value)) && String(value).trim() !== '') {
                                    grossPrice = Number(value);
                                }
                                rowData['Gross Price'] = grossPrice;
                            }

                            // Apply discount to Price column
                            if (fileInfo.discount !== undefined && fileInfo.discount !== 0) {
                                if (isPriceColumn) {
                                    if (typeof value === 'number') {
                                        finalValue = value * (1 - fileInfo.discount);
                                    } else if (!isNaN(Number(value)) && String(value).trim() !== '') {
                                        finalValue = Number(value) * (1 - fileInfo.discount);
                                    }
                                }
                            }

                            rowData[headerName] = finalValue;
                        }

                        if (hasData) {
                            // Ensure 'Gross Price' is in the headers list if we added it to rowData
                            if (rowData.hasOwnProperty('Gross Price') && !headers.includes('Gross Price')) {
                                // Find the 'Price' header index to insert 'Gross Price' before it
                                const priceIndex = headers.findIndex(h => {
                                    const hLower = h.toLowerCase().trim();
                                    return hLower === 'price' || hLower.includes('price');
                                });
                                if (priceIndex !== -1) {
                                    headers.splice(priceIndex, 0, 'Gross Price');
                                } else {
                                    headers.push('Gross Price');
                                }
                            }
                            // Recalculate Total = Qty * Price
                            let qtyKey: string | undefined;
                            let priceKey: string | undefined;
                            let totalKey: string | undefined;

                            for (const key of Object.keys(rowData)) {
                                const keyLower = key.toLowerCase().trim();
                                if (keyLower === 'qty' || keyLower === 'quantity') {
                                    qtyKey = key;
                                } else if (keyLower === 'price' || keyLower.includes('price')) {
                                    priceKey = key;
                                } else if (keyLower === 'total' || keyLower.includes('total')) {
                                    totalKey = key;
                                }
                            }

                            if (qtyKey && priceKey && totalKey) {
                                const qty = Number(rowData[qtyKey]);
                                const price = Number(rowData[priceKey]);
                                
                                if (!isNaN(qty) && !isNaN(price)) {
                                    rowData[totalKey] = qty * price;
                                }
                            }

                            rows.push(rowData);
                        } else {
                            // Stop if we hit an empty row
                            break;
                        }
                    }

                    resolve({ headers, rows });
                } catch (error) {
                    reject(error);
                }
            };

            reader.onerror = () => {
                reject(new Error('File reading error'));
            };

            if (fileInfo.file) {
                reader.readAsArrayBuffer(fileInfo.file);
            }
        });
    }
}
