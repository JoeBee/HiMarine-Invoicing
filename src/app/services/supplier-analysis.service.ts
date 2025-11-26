import { Injectable } from '@angular/core';
import { BehaviorSubject, Observable } from 'rxjs';
import { SupplierAnalysisFileInfo } from '../components/supplier-analysis-inputs/supplier-analysis-inputs.component';
import * as XLSX from 'xlsx';

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
    private filesSubject = new BehaviorSubject<SupplierAnalysisFileInfo[]>([]);
    files$: Observable<SupplierAnalysisFileInfo[]> = this.filesSubject.asObservable();
    
    private files2Subject = new BehaviorSubject<SupplierAnalysisFileInfo[]>([]);
    files2$: Observable<SupplierAnalysisFileInfo[]> = this.files2Subject.asObservable();

    setFiles(files: SupplierAnalysisFileInfo[]): void {
        this.filesSubject.next(files);
    }

    getFiles(): SupplierAnalysisFileInfo[] {
        return this.filesSubject.value;
    }
    
    setFiles2(files: SupplierAnalysisFileInfo[]): void {
        this.files2Subject.next(files);
    }

    getFiles2(): SupplierAnalysisFileInfo[] {
        return this.files2Subject.value;
    }

    async extractDataFromFile(fileInfo: SupplierAnalysisFileInfo): Promise<{ headers: string[]; rows: ExcelRowData[] }> {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();

            reader.onload = (e: any) => {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const firstSheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[firstSheetName];

                    const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1:Z100');
                    const topLeftCellRef = XLSX.utils.decode_cell(fileInfo.topLeftCell);

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
                            
                            rowData[headers[col - range.s.c]] = value;
                        }

                        if (hasData) {
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

            reader.readAsArrayBuffer(fileInfo.file);
        });
    }
}

