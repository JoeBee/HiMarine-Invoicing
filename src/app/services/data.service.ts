import { Injectable } from '@angular/core';
import { BehaviorSubject, Observable } from 'rxjs';
import * as XLSX from 'xlsx';

export interface SupplierFileInfo {
    fileName: string;
    topLeftCell: string;
    descriptionColumn: string;
    priceColumn: string;
    unitColumn: string;
    remarksColumn: string;
    descriptionHeader: string;
    priceHeader: string;
    unitHeader: string;
    remarksHeader: string;
    rowCount: number;
    file: File;
    hasData?: boolean; // Track if file has data after processing
    category?: string; // Track which drop-zone the file was dropped on
}

export interface ProcessedDataRow {
    fileName: string;
    description: string;
    price: number;
    unit: string;
    remarks: string;
    count: number;
    included: boolean; // Track if row is included in the final output
    category?: string; // Track which drop-zone the file was dropped on
    originalData?: any[]; // Store original row data from XLSX
    originalHeaders?: string[]; // Store original column headers
}

export interface ExcelItemData {
    pos: number;
    description: string;
    remark: string;
    unit: string;
    qty: number;
    price: number;
    total: number;
    tabName: string;
    currency: string;
}

export interface ExcelProcessedData {
    [tabName: string]: {
        recordsWithTotal: number;
        sumOfTotals: number;
        items: ExcelItemData[];
    };
}

type PriceDividerMap = Record<string, number>;

const DEFAULT_PRICE_DIVIDERS: PriceDividerMap = {
    Bonded: 0.9,
    Provisions: 0.9
};

@Injectable({
    providedIn: 'root'
})
export class DataService {
    private supplierFilesSubject = new BehaviorSubject<SupplierFileInfo[]>([]);
    private processedDataSubject = new BehaviorSubject<ProcessedDataRow[]>([]);
    private priceDividerSubject = new BehaviorSubject<PriceDividerMap>({ ...DEFAULT_PRICE_DIVIDERS });
    private separateFreshProvisionsSubject = new BehaviorSubject<boolean>(false); // Default to "Do not Separate"
    private excelDataSubject = new BehaviorSubject<ExcelProcessedData | null>(null);

    supplierFiles$: Observable<SupplierFileInfo[]> = this.supplierFilesSubject.asObservable();
    processedData$: Observable<ProcessedDataRow[]> = this.processedDataSubject.asObservable();
    priceDivider$: Observable<PriceDividerMap> = this.priceDividerSubject.asObservable();
    separateFreshProvisions$: Observable<boolean> = this.separateFreshProvisionsSubject.asObservable();
    excelData$: Observable<ExcelProcessedData | null> = this.excelDataSubject.asObservable();

    constructor() { }

    addSupplierFiles(files: File[], category?: string): Promise<void> {
        return new Promise(async (resolve) => {
            const currentFiles = this.supplierFilesSubject.value;
            const newFileInfos: SupplierFileInfo[] = [];

            for (const file of files) {
                const fileInfo = await this.analyzeFile(file);
                fileInfo.category = category;
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
                let remarksColumn = 'NOT FOUND';
                let descriptionHeader = '';
                let priceHeader = '';
                let unitHeader = '';
                let remarksHeader = '';
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
                                // console.log("cellValue", cellValue);
                                // skip rows with less than 25 characters
                                // continue;
                            }
                            else if (cellValue.includes('description') ||
                                cellValue.includes('descrption') ||
                                cellValue.includes('product') ||
                                cellValue.includes('item') ||
                                cellValue.includes('name') ||
                                cellValue.includes('product description') ||
                                cellValue.includes('product description (en)')) {

                                // console.log("cellValue", cellValue);

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
                                            headerValue.includes('name') || headerValue.includes('product description') ||
                                            headerValue.includes('product description (en)')) {
                                            descriptionColumn = XLSX.utils.encode_col(searchCol);
                                            descriptionHeader = headerText;
                                        }

                                        if ((headerValue.includes('price') ||
                                            headerValue.includes('cost') ||
                                            headerValue.includes('unit aud') ||
                                            headerValue.includes('value') ||
                                            headerValue.includes('precio')) && headerValue.length < 25) {
                                            priceColumn = XLSX.utils.encode_col(searchCol);
                                            priceHeader = headerText;
                                        }

                                        if (headerValue === 'unit' ||
                                            headerValue === 'units' ||
                                            headerValue === 'uom' ||
                                            headerValue === 'uoms' ||
                                            headerValue === 'u.m.' ||
                                            headerValue === 'um' ||
                                            headerValue === 'u.o.m.' ||
                                            headerValue === 'u m') {
                                            unitColumn = XLSX.utils.encode_col(searchCol);
                                            unitHeader = headerText;
                                        }

                                        if (headerValue.includes('remark') ||
                                            headerValue.includes('comment') ||
                                            headerValue.includes('comentarios') ||
                                            headerValue.includes('presentation')) {
                                            remarksColumn = XLSX.utils.encode_col(searchCol);
                                            remarksHeader = headerText;
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
                                if (remarksColumn === 'NOT FOUND') {
                                    remarksHeader = 'NOT FOUND';
                                }

                                // Count data rows (excluding header row) - only if columns were found
                                if (descriptionColumn !== 'NOT FOUND' && priceColumn !== 'NOT FOUND') {
                                    const descColIndex = XLSX.utils.decode_col(descriptionColumn);
                                    const priceColIndex = XLSX.utils.decode_col(priceColumn);
                                    let consecutiveBlankDescriptions = 0;

                                    for (let dataRow = row + 1; dataRow <= range.e.r; dataRow++) {
                                        const descAddress = XLSX.utils.encode_cell({ r: dataRow, c: descColIndex });
                                        const priceAddress = XLSX.utils.encode_cell({ r: dataRow, c: priceColIndex });

                                        const descCell = worksheet[descAddress];
                                        const priceCell = worksheet[priceAddress];

                                        const hasDescription = descCell && descCell.v && String(descCell.v).trim() !== '';
                                        const hasPrice = priceCell && priceCell.v && String(priceCell.v).trim() !== '';

                                        // Check for 3 consecutive blank descriptions
                                        if (!hasDescription) {
                                            consecutiveBlankDescriptions++;
                                            if (consecutiveBlankDescriptions >= 3) {
                                                break;
                                            }
                                            continue;
                                        }

                                        // Reset consecutive blank descriptions counter when we find a description
                                        consecutiveBlankDescriptions = 0;

                                        // Check if this row has only description (no price)
                                        if (hasDescription && !hasPrice) {
                                            // Look ahead to the next row
                                            const nextRow = dataRow + 1;
                                            let shouldExcludeAndStop = false;

                                            if (nextRow <= range.e.r) {
                                                const nextDescAddress = XLSX.utils.encode_cell({ r: nextRow, c: descColIndex });
                                                const nextPriceAddress = XLSX.utils.encode_cell({ r: nextRow, c: priceColIndex });
                                                const nextDescCell = worksheet[nextDescAddress];
                                                const nextPriceCell = worksheet[nextPriceAddress];

                                                const nextHasDescription = nextDescCell && nextDescCell.v && String(nextDescCell.v).trim() !== '';
                                                const nextHasPrice = nextPriceCell && nextPriceCell.v && String(nextPriceCell.v).trim() !== '';

                                                // If next row has no price and no description, exclude current row and stop
                                                if (!nextHasPrice && !nextHasDescription) {
                                                    shouldExcludeAndStop = true;
                                                }
                                            } else {
                                                // If we're at the end of the range and this row only has description (no price),
                                                // exclude it as it's likely not a valid data row
                                                shouldExcludeAndStop = true;
                                            }

                                            if (shouldExcludeAndStop) {
                                                break;
                                            }
                                        }

                                        // Count this row
                                        rowCount++;
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
                    remarksColumn,
                    descriptionHeader,
                    priceHeader,
                    unitHeader,
                    remarksHeader,
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
        const updatedSupplierFiles: SupplierFileInfo[] = [];

        for (const fileInfo of supplierFiles) {
            const rowData = await this.extractDataFromFile(fileInfo);
            allData.push(...rowData);

            // Update the file info to track if it has data
            const updatedFileInfo: SupplierFileInfo = {
                ...fileInfo,
                hasData: rowData.length > 0
            };
            updatedSupplierFiles.push(updatedFileInfo);
        }

        // Update supplier files with processing status
        this.supplierFilesSubject.next(updatedSupplierFiles);

        // Data is not sorted here - maintains original upload order
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
                const remarksColIndex = XLSX.utils.decode_col(fileInfo.remarksColumn);

                const rows: ProcessedDataRow[] = [];

                // Extract original headers from the header row
                const originalHeaders: string[] = [];
                for (let col = range.s.c; col <= range.e.c; col++) {
                    const headerAddress = XLSX.utils.encode_cell({ r: topLeftCellRef.r, c: col });
                    const headerCell = worksheet[headerAddress];
                    originalHeaders.push(headerCell && headerCell.v ? String(headerCell.v) : '');
                }

                // Start from the row after the header
                let consecutiveBlankDescriptions = 0;

                for (let row = topLeftCellRef.r + 1; row <= range.e.r; row++) {
                    const descAddress = XLSX.utils.encode_cell({ r: row, c: descColIndex });
                    const priceAddress = XLSX.utils.encode_cell({ r: row, c: priceColIndex });
                    const unitAddress = XLSX.utils.encode_cell({ r: row, c: unitColIndex });
                    const remarksAddress = XLSX.utils.encode_cell({ r: row, c: remarksColIndex });

                    const descCell = worksheet[descAddress];
                    const priceCell = worksheet[priceAddress];
                    const unitCell = worksheet[unitAddress];
                    const remarksCell = worksheet[remarksAddress];

                    const hasDescription = descCell && descCell.v && String(descCell.v).trim() !== '';
                    const hasPrice = priceCell && priceCell.v && String(priceCell.v).trim() !== '';

                    // Check for 3 consecutive blank descriptions
                    if (!hasDescription) {
                        consecutiveBlankDescriptions++;
                        if (consecutiveBlankDescriptions >= 3) {
                            break;
                        }
                        continue;
                    }

                    // Reset consecutive blank descriptions counter when we find a description
                    consecutiveBlankDescriptions = 0;

                    // Check if this row has only description (no price)
                    if (hasDescription && !hasPrice) {
                        // Look ahead to the next row
                        const nextRow = row + 1;
                        let shouldExcludeAndStop = false;

                        if (nextRow <= range.e.r) {
                            const nextDescAddress = XLSX.utils.encode_cell({ r: nextRow, c: descColIndex });
                            const nextPriceAddress = XLSX.utils.encode_cell({ r: nextRow, c: priceColIndex });
                            const nextDescCell = worksheet[nextDescAddress];
                            const nextPriceCell = worksheet[nextPriceAddress];

                            const nextHasDescription = nextDescCell && nextDescCell.v && String(nextDescCell.v).trim() !== '';
                            const nextHasPrice = nextPriceCell && nextPriceCell.v && String(nextPriceCell.v).trim() !== '';

                            // If next row has no price and no description, exclude current row and stop
                            if (!nextHasPrice && !nextHasDescription) {
                                shouldExcludeAndStop = true;
                            }
                        } else {
                            // If we're at the end of the range and this row only has description (no price),
                            // exclude it as it's likely not a valid data row
                            shouldExcludeAndStop = true;
                        }

                        if (shouldExcludeAndStop) {
                            break;
                        }
                    }

                    // Extract original row data
                    const originalRowData: any[] = [];
                    for (let col = range.s.c; col <= range.e.c; col++) {
                        const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
                        const cell = worksheet[cellAddress];
                        originalRowData.push(cell && cell.v ? cell.v : '');
                    }

                    // Apply price divider to the price
                    let adjustedPrice = NaN;
                    if (priceCell && priceCell.v) {
                        const originalPrice = Number(priceCell.v);
                        const priceDivider = this.getPriceDividerForCategory(fileInfo.category);
                        const safeDivider = priceDivider > 0 ? priceDivider : 1;
                        adjustedPrice = originalPrice / safeDivider;
                    }

                    rows.push({
                        fileName: fileInfo.fileName,
                        description: String(descCell.v),
                        price: adjustedPrice,
                        unit: unitCell && unitCell.v ? String(unitCell.v) : '',
                        remarks: remarksCell && remarksCell.v ? String(remarksCell.v) : '',
                        count: 0,
                        included: true, // Default to included
                        category: fileInfo.category,
                        originalData: originalRowData,
                        originalHeaders: originalHeaders
                    });
                }

                resolve(rows);
            };

            reader.readAsArrayBuffer(fileInfo.file);
        });
    }

    private getPriceDividerForCategory(category?: string): number {
        const dividers = {
            ...DEFAULT_PRICE_DIVIDERS,
            ...this.priceDividerSubject.value
        };

        if (category && dividers[category] !== undefined) {
            return dividers[category];
        }

        if (dividers['Bonded'] !== undefined) {
            return dividers['Bonded'];
        }

        const firstDivider = Object.values(dividers).find(value => typeof value === 'number');
        return typeof firstDivider === 'number' ? firstDivider : 1;
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

    updateRowIncluded(index: number, included: boolean): void {
        const currentData = this.processedDataSubject.value;
        if (index >= 0 && index < currentData.length) {
            currentData[index].included = included;
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

    setPriceDivider(divider: number, category?: string): void {
        const current = this.priceDividerSubject.value;
        const normalizedCategory = category ?? null;

        let updated: PriceDividerMap;
        if (normalizedCategory) {
            updated = {
                ...DEFAULT_PRICE_DIVIDERS,
                ...current,
                [normalizedCategory]: divider
            };
        } else {
            const categories = Object.keys({ ...DEFAULT_PRICE_DIVIDERS, ...current });
            updated = categories.reduce((map, key) => {
                map[key] = divider;
                return map;
            }, {} as PriceDividerMap);
        }

        this.priceDividerSubject.next(updated);
        // Automatically reprocess data when price divider changes
        if (this.supplierFilesSubject.value.length > 0) {
            this.processSupplierFiles();
        }
    }

    getPriceDivider(category?: string): number {
        return this.getPriceDividerForCategory(category);
    }

    getPriceDividers(): PriceDividerMap {
        return {
            ...DEFAULT_PRICE_DIVIDERS,
            ...this.priceDividerSubject.value
        };
    }

    setExcelData(data: ExcelProcessedData): void {
        this.excelDataSubject.next(data);
    }

    getExcelData(): ExcelProcessedData | null {
        return this.excelDataSubject.value;
    }

    clearExcelData(): void {
        this.excelDataSubject.next(null);
    }

    setSeparateFreshProvisions(separate: boolean): void {
        this.separateFreshProvisionsSubject.next(separate);
    }

    getSeparateFreshProvisions(): boolean {
        return this.separateFreshProvisionsSubject.value;
    }

    async updateFileTopLeftCell(fileName: string, topLeftCell: string): Promise<void> {
        const currentFiles = this.supplierFilesSubject.value;
        const fileIndex = currentFiles.findIndex(f => f.fileName === fileName);

        if (fileIndex === -1) {
            return;
        }

        const fileInfo = currentFiles[fileIndex];

        // Re-analyze the file with the new top left cell
        const updatedFileInfo = await this.analyzeFileWithTopLeft(fileInfo.file, topLeftCell, fileInfo.category);

        // Update the file in the array
        const updatedFiles = [...currentFiles];
        updatedFiles[fileIndex] = updatedFileInfo;

        this.supplierFilesSubject.next(updatedFiles);

        // Reprocess files to update data
        await this.processSupplierFiles();
    }

    private async analyzeFileWithTopLeft(file: File, topLeftCell: string, category?: string): Promise<SupplierFileInfo> {
        return new Promise((resolve) => {
            const reader = new FileReader();

            reader.onload = (e: any) => {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];

                const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1:Z100');
                let topLeftCellRef: XLSX.CellAddress;

                try {
                    topLeftCellRef = XLSX.utils.decode_cell(topLeftCell);
                } catch {
                    topLeftCellRef = XLSX.utils.decode_cell('A1');
                }

                let descriptionColumn = 'NOT FOUND';
                let priceColumn = 'NOT FOUND';
                let unitColumn = 'NOT FOUND';
                let remarksColumn = 'NOT FOUND';
                let descriptionHeader = '';
                let priceHeader = '';
                let unitHeader = '';
                let remarksHeader = '';
                let rowCount = 0;

                // Extract headers from the specified top left cell
                for (let col = range.s.c; col <= range.e.c; col++) {
                    const headerAddress = XLSX.utils.encode_cell({ r: topLeftCellRef.r, c: col });
                    const headerCell = worksheet[headerAddress];

                    if (headerCell && headerCell.v) {
                        const headerValue = String(headerCell.v).toLowerCase().trim();
                        const headerText = String(headerCell.v);

                        if (headerValue.includes('description') || headerValue.includes('descrption') ||
                            headerValue.includes('item') || headerValue.includes('product') ||
                            headerValue.includes('name') || headerValue.includes('product description') ||
                            headerValue.includes('product description (en)')) {
                            descriptionColumn = XLSX.utils.encode_col(col);
                            descriptionHeader = headerText;
                        }

                        if ((headerValue.includes('price') ||
                            headerValue.includes('cost') ||
                            headerValue.includes('unit aud') ||
                            headerValue.includes('value') ||
                            headerValue.includes('precio')) && headerValue.length < 25) {
                            priceColumn = XLSX.utils.encode_col(col);
                            priceHeader = headerText;
                        }

                        if (headerValue === 'unit' ||
                            headerValue === 'units' ||
                            headerValue === 'uom' ||
                            headerValue === 'uoms' ||
                            headerValue === 'u.m.' ||
                            headerValue === 'um' ||
                            headerValue === 'u.o.m.' ||
                            headerValue === 'u m') {
                            unitColumn = XLSX.utils.encode_col(col);
                            unitHeader = headerText;
                        }

                        if (headerValue.includes('remark') ||
                            headerValue.includes('comment') ||
                            headerValue.includes('comentarios') ||
                            headerValue.includes('presentation')) {
                            remarksColumn = XLSX.utils.encode_col(col);
                            remarksHeader = headerText;
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
                if (remarksColumn === 'NOT FOUND') {
                    remarksHeader = 'NOT FOUND';
                }

                // Count rows
                if (descriptionColumn !== 'NOT FOUND' && priceColumn !== 'NOT FOUND') {
                    const descColIndex = XLSX.utils.decode_col(descriptionColumn);
                    const priceColIndex = XLSX.utils.decode_col(priceColumn);
                    let consecutiveBlankDescriptions = 0;

                    for (let dataRow = topLeftCellRef.r + 1; dataRow <= range.e.r; dataRow++) {
                        const descAddress = XLSX.utils.encode_cell({ r: dataRow, c: descColIndex });
                        const priceAddress = XLSX.utils.encode_cell({ r: dataRow, c: priceColIndex });

                        const descCell = worksheet[descAddress];
                        const priceCell = worksheet[priceAddress];

                        const hasDescription = descCell && descCell.v && String(descCell.v).trim() !== '';
                        const hasPrice = priceCell && priceCell.v && String(priceCell.v).trim() !== '';

                        // Check for 3 consecutive blank descriptions
                        if (!hasDescription) {
                            consecutiveBlankDescriptions++;
                            if (consecutiveBlankDescriptions >= 3) {
                                break;
                            }
                            continue;
                        }

                        // Reset consecutive blank descriptions counter when we find a description
                        consecutiveBlankDescriptions = 0;

                        // Check if this row has only description (no price)
                        if (hasDescription && !hasPrice) {
                            // Look ahead to the next row
                            const nextRow = dataRow + 1;
                            let shouldExcludeAndStop = false;

                            if (nextRow <= range.e.r) {
                                const nextDescAddress = XLSX.utils.encode_cell({ r: nextRow, c: descColIndex });
                                const nextPriceAddress = XLSX.utils.encode_cell({ r: nextRow, c: priceColIndex });
                                const nextDescCell = worksheet[nextDescAddress];
                                const nextPriceCell = worksheet[nextPriceAddress];

                                const nextHasDescription = nextDescCell && nextDescCell.v && String(nextDescCell.v).trim() !== '';
                                const nextHasPrice = nextPriceCell && nextPriceCell.v && String(nextPriceCell.v).trim() !== '';

                                // If next row has no price and no description, exclude current row and stop
                                if (!nextHasPrice && !nextHasDescription) {
                                    shouldExcludeAndStop = true;
                                }
                            } else {
                                // If we're at the end of the range and this row only has description (no price),
                                // exclude it as it's likely not a valid data row
                                shouldExcludeAndStop = true;
                            }

                            if (shouldExcludeAndStop) {
                                break;
                            }
                        }

                        // Count this row
                        rowCount++;
                    }
                } else {
                    // Count rows based on any data
                    let consecutiveBlankRows = 0;
                    for (let dataRow = topLeftCellRef.r + 1; dataRow <= range.e.r; dataRow++) {
                        let hasData = false;
                        for (let col = range.s.c; col <= range.e.c; col++) {
                            const cellAddress = XLSX.utils.encode_cell({ r: dataRow, c: col });
                            const cell = worksheet[cellAddress];
                            if (cell && cell.v !== null && cell.v !== undefined && String(cell.v).trim() !== '') {
                                hasData = true;
                                break;
                            }
                        }
                        if (hasData) {
                            rowCount++;
                            consecutiveBlankRows = 0;
                        } else {
                            consecutiveBlankRows++;
                            if (consecutiveBlankRows >= 3) {
                                break;
                            }
                        }
                    }
                }

                const fileName = file.name.replace(/\.[^/.]+$/, ''); // Remove extension

                resolve({
                    fileName,
                    topLeftCell,
                    descriptionColumn,
                    priceColumn,
                    unitColumn,
                    remarksColumn,
                    descriptionHeader,
                    priceHeader,
                    unitHeader,
                    remarksHeader,
                    rowCount,
                    file,
                    category
                });
            };

            reader.readAsArrayBuffer(file);
        });
    }
}

