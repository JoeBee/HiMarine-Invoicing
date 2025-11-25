import { Component, OnInit } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { LoggingService } from '../../services/logging.service';
import { SupplierAnalysisService } from '../../services/supplier-analysis.service';
import * as XLSX from 'xlsx';

export interface SupplierAnalysisFileInfo {
    fileName: string;
    file: File;
    rowCount: number;
    topLeftCell: string;
    category: 'Invoice' | 'Supplier Quotations';
}

@Component({
    selector: 'app-supplier-analysis-inputs',
    standalone: true,
    imports: [CommonModule, FormsModule],
    templateUrl: './supplier-analysis-inputs.component.html',
    styleUrls: ['./supplier-analysis-inputs.component.scss']
})
export class SupplierAnalysisInputsComponent implements OnInit {
    invoiceFiles: SupplierAnalysisFileInfo[] = [];
    supplierQuotationFiles: SupplierAnalysisFileInfo[] = [];
    isDragging = false;
    isProcessing = false;
    hoveredDropzone: string | null = null;

    constructor(
        private loggingService: LoggingService,
        private supplierAnalysisService: SupplierAnalysisService
    ) { }

    ngOnInit(): void {
        // Load existing files from service if available
        const existingFiles = this.supplierAnalysisService.getFiles();
        if (existingFiles.length > 0) {
            this.invoiceFiles = existingFiles.filter(f => f.category === 'Invoice');
            this.supplierQuotationFiles = existingFiles.filter(f => f.category === 'Supplier Quotations');
        }
    }

    onDragOver(event: DragEvent, category: string): void {
        event.preventDefault();
        event.stopPropagation();
        this.isDragging = true;
        this.hoveredDropzone = category;
    }

    onDragLeave(event: DragEvent): void {
        event.preventDefault();
        event.stopPropagation();
        this.isDragging = false;
        this.hoveredDropzone = null;
    }

    async onDrop(event: DragEvent, category: 'Invoice' | 'Supplier Quotations'): Promise<void> {
        event.preventDefault();
        event.stopPropagation();
        this.isDragging = false;
        this.hoveredDropzone = null;

        const files = event.dataTransfer?.files;
        if (files) {
            this.loggingService.logUserAction('files_dropped', {
                category: category,
                filesCount: files.length,
                fileNames: Array.from(files).map(f => f.name)
            }, 'SupplierAnalysisInputsComponent');

            await this.processFiles(files, category);
        }
    }

    onFileSelect(event: Event, category: 'Invoice' | 'Supplier Quotations'): void {
        const input = event.target as HTMLInputElement;
        if (input.files) {
            this.loggingService.logUserAction('files_selected', {
                category: category,
                filesCount: input.files.length,
                fileNames: Array.from(input.files).map(f => f.name)
            }, 'SupplierAnalysisInputsComponent');

            this.processFiles(input.files, category);
        }
    }

    async processFiles(fileList: FileList, category: 'Invoice' | 'Supplier Quotations'): Promise<void> {
        this.isProcessing = true;
        const files = Array.from(fileList).filter(f =>
            f.name.toLowerCase().endsWith('.xlsx') || 
            f.name.toLowerCase().endsWith('.xls') ||
            f.name.toLowerCase().endsWith('.xlsm')
        );

        if (files.length > 0) {
            files.forEach(file => {
                this.loggingService.logFileUpload(
                    file.name,
                    file.size,
                    file.type,
                    category,
                    'SupplierAnalysisInputsComponent'
                );
            });

            try {
                // Invoice: only allow 1 file, replace existing
                if (category === 'Invoice') {
                    // Clear existing invoice files
                    this.invoiceFiles = [];
                    // Process only the first file
                    if (files.length > 0) {
                        const fileInfo = await this.analyzeFile(files[0], category);
                        this.invoiceFiles.push(fileInfo);
                    }
                } else {
                    // Supplier Quotations: allow multiple files
                    for (const file of files) {
                        const fileInfo = await this.analyzeFile(file, category);
                        this.supplierQuotationFiles.push(fileInfo);
                    }
                }

                    this.loggingService.logDataProcessing('files_processed_successfully', {
                        fileCount: files.length,
                        category: category,
                        fileNames: files.map(f => f.name)
                    }, 'SupplierAnalysisInputsComponent');

                    // Update service with all files
                    this.updateServiceFiles();
            } catch (error) {
                this.loggingService.logError(error as Error, 'file_processing', 'SupplierAnalysisInputsComponent', {
                    fileCount: files.length,
                    category: category
                });
            }
        }
        this.isProcessing = false;
    }

    private async analyzeFile(file: File, category: 'Invoice' | 'Supplier Quotations'): Promise<SupplierAnalysisFileInfo> {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();

            reader.onload = (e: any) => {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const firstSheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[firstSheetName];

                    const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1:Z100');
                    
                    // Auto-detect the top left cell of the datatable (similar to captains-request)
                    const topLeftCell = this.findFirstDataCell(worksheet, range) || 'A1';
                    const topLeftCellRef = XLSX.utils.decode_cell(topLeftCell);

                    // Count rows starting from the row after topLeftCell
                    let rowCount = 0;
                    let descriptionColumn = -1;

                    // Normalize function to match headers
                    const normalize = (str: string): string => {
                        return str.toLowerCase()
                            .replace(/[^\w\s]/g, '') // Remove punctuation
                            .replace(/s$/, '') // Remove trailing 's'
                            .trim();
                    };

                    // Find description column in header row
                    for (let col = topLeftCellRef.c; col <= range.e.c; col++) {
                        const headerAddress = XLSX.utils.encode_cell({ r: topLeftCellRef.r, c: col });
                        const headerCell = worksheet[headerAddress];
                        if (headerCell && headerCell.v) {
                            const normalizedValue = normalize(String(headerCell.v));
                            if (normalizedValue === 'description') {
                                descriptionColumn = col;
                                break;
                            }
                        }
                    }

                    // Count data rows
                    if (descriptionColumn !== -1) {
                        for (let dataRow = topLeftCellRef.r + 1; dataRow <= range.e.r; dataRow++) {
                            const descAddress = XLSX.utils.encode_cell({ r: dataRow, c: descriptionColumn });
                            const descCell = worksheet[descAddress];
                            const hasDescription = descCell && descCell.v !== null && descCell.v !== undefined && String(descCell.v).trim() !== '';
                            
                            if (hasDescription) {
                                rowCount++;
                            } else {
                                // Stop counting if we hit consecutive empty rows
                                let consecutiveEmptyRows = 0;
                                for (let checkRow = dataRow; checkRow <= Math.min(dataRow + 2, range.e.r); checkRow++) {
                                    const checkAddress = XLSX.utils.encode_cell({ r: checkRow, c: descriptionColumn });
                                    const checkCell = worksheet[checkAddress];
                                    if (!checkCell || !checkCell.v || String(checkCell.v).trim() === '') {
                                        consecutiveEmptyRows++;
                                    } else {
                                        break;
                                    }
                                }
                                if (consecutiveEmptyRows >= 3) {
                                    break;
                                }
                            }
                        }
                    } else {
                        // Fallback: count rows with any data
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
                            } else {
                                break;
                            }
                        }
                    }

                    const fileName = file.name;

                    resolve({
                        fileName,
                        file,
                        rowCount,
                        topLeftCell,
                        category
                    });
                } catch (error) {
                    reject(error);
                }
            };

            reader.onerror = () => {
                reject(new Error('File reading error'));
            };

            reader.readAsArrayBuffer(file);
        });
    }

    private findFirstDataCell(worksheet: XLSX.WorkSheet, range: XLSX.Range): string | undefined {
        // Required column headers: 'Pos', 'Description', 'Remark', 'Unit', 'Qty', 'Price', 'Total'
        // Ignore punctuation and trailing 's'
        const requiredHeaders = ['pos', 'description', 'remark', 'unit', 'qty', 'price', 'total'];
        const maxRowsToScan = Math.min(range.e.r, 50);
        const maxColsToScan = Math.min(range.e.c, 25);

        // Normalize a string by removing punctuation and trailing 's'
        const normalize = (str: string): string => {
            return str.toLowerCase()
                .replace(/[^\w\s]/g, '') // Remove punctuation
                .replace(/s$/, '') // Remove trailing 's'
                .trim();
        };

        for (let row = range.s.r; row <= maxRowsToScan; row++) {
            // Check if this row contains the required headers
            const headersInRow: string[] = [];
            for (let col = range.s.c; col <= maxColsToScan; col++) {
                const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
                const cell = worksheet[cellAddress];
                if (cell && cell.v !== undefined && cell.v !== null) {
                    const normalizedValue = normalize(String(cell.v));
                    if (normalizedValue && requiredHeaders.includes(normalizedValue)) {
                        headersInRow.push(normalizedValue);
                    }
                }
            }

            // If we found at least one required header in this row, this is likely the header row
            // Return the first cell of this row (leftmost column)
            if (headersInRow.length > 0) {
                return XLSX.utils.encode_cell({ r: row, c: range.s.c });
            }
        }

        return undefined;
    }

    triggerFileInput(category: 'Invoice' | 'Supplier Quotations'): void {
        this.loggingService.logButtonClick('file_input_triggered', 'SupplierAnalysisInputsComponent', {
            category: category
        });

        const fileInput = document.getElementById(`fileInput-${category}`) as HTMLInputElement;
        fileInput?.click();
    }

    getTopLeftCellOptions(): string[] {
        const options: string[] = [];
        for (let col = 0; col < 8; col++) {
            for (let row = 1; row <= 20; row++) {
                const cellRef = XLSX.utils.encode_cell({ r: row - 1, c: col });
                options.push(cellRef);
            }
        }
        return options;
    }

    onTopLeftCellChange(event: Event, fileInfo: SupplierAnalysisFileInfo): void {
        const input = event.target as HTMLInputElement;
        let value = input.value.toUpperCase();

        const regex = /^[A-Z][0-9]{1,2}$/;
        if (value && !regex.test(value)) {
            value = value.replace(/[^A-Z0-9]/g, '');
            if (value && !/^[A-Z]/.test(value)) {
                value = '';
            }
            const match = value.match(/^([A-Z])([0-9]{0,2})/);
            if (match) {
                value = match[0];
            } else if (value.length > 0 && /^[A-Z]/.test(value)) {
                value = value.charAt(0);
            }
        }

        input.value = value;
        fileInfo.topLeftCell = value;

        if (value && value.trim() !== '' && value !== 'NOT FOUND') {
            input.removeAttribute('list');
            void this.reanalyzeFileWithTopLeft(fileInfo, value);
        } else {
            fileInfo.rowCount = 0;
        }
    }

    validateTopLeftCell(fileInfo: SupplierAnalysisFileInfo): void {
        const regex = /^[A-Z][0-9]{1,2}$/;
        if (fileInfo.topLeftCell && fileInfo.topLeftCell !== 'NOT FOUND' && !regex.test(fileInfo.topLeftCell)) {
            fileInfo.topLeftCell = 'A1';
            fileInfo.rowCount = 0;
        }
    }

    private async reanalyzeFileWithTopLeft(fileInfo: SupplierAnalysisFileInfo, topLeftCell: string): Promise<void> {
        this.isProcessing = true;
        try {
            return new Promise((resolve) => {
                const reader = new FileReader();

                reader.onload = (e: any) => {
                    try {
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

                        // Count rows starting from the row after topLeftCell (similar to captains-request)
                        let rowCount = 0;
                        let descriptionColumn = -1;

                        // Normalize function to match headers
                        const normalize = (str: string): string => {
                            return str.toLowerCase()
                                .replace(/[^\w\s]/g, '') // Remove punctuation
                                .replace(/s$/, '') // Remove trailing 's'
                                .trim();
                        };

                        // Find description column in header row
                        for (let col = topLeftCellRef.c; col <= range.e.c; col++) {
                            const headerAddress = XLSX.utils.encode_cell({ r: topLeftCellRef.r, c: col });
                            const headerCell = worksheet[headerAddress];
                            if (headerCell && headerCell.v) {
                                const normalizedValue = normalize(String(headerCell.v));
                                if (normalizedValue === 'description') {
                                    descriptionColumn = col;
                                    break;
                                }
                            }
                        }

                        // Count data rows
                        if (descriptionColumn !== -1) {
                            for (let dataRow = topLeftCellRef.r + 1; dataRow <= range.e.r; dataRow++) {
                                const descAddress = XLSX.utils.encode_cell({ r: dataRow, c: descriptionColumn });
                                const descCell = worksheet[descAddress];
                                const hasDescription = descCell && descCell.v !== null && descCell.v !== undefined && String(descCell.v).trim() !== '';
                                
                                if (hasDescription) {
                                    rowCount++;
                                } else {
                                    // Stop counting if we hit consecutive empty rows
                                    let consecutiveEmptyRows = 0;
                                    for (let checkRow = dataRow; checkRow <= Math.min(dataRow + 2, range.e.r); checkRow++) {
                                        const checkAddress = XLSX.utils.encode_cell({ r: checkRow, c: descriptionColumn });
                                        const checkCell = worksheet[checkAddress];
                                        if (!checkCell || !checkCell.v || String(checkCell.v).trim() === '') {
                                            consecutiveEmptyRows++;
                                        } else {
                                            break;
                                        }
                                    }
                                    if (consecutiveEmptyRows >= 3) {
                                        break;
                                    }
                                }
                            }
                        } else {
                            // Fallback: count rows with any data
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
                                } else {
                                    break;
                                }
                            }
                        }

                        fileInfo.topLeftCell = topLeftCell;
                        fileInfo.rowCount = rowCount;

                        // Update service with all files
                        this.updateServiceFiles();

                        this.loggingService.logUserAction('top_left_cell_changed', {
                            fileName: fileInfo.fileName,
                            newTopLeftCell: topLeftCell,
                            rowCount: rowCount
                        }, 'SupplierAnalysisInputsComponent');

                        resolve();
                    } catch (error) {
                        this.loggingService.logError(error as Error, 'reanalyze_file_error', 'SupplierAnalysisInputsComponent', {
                            fileName: fileInfo.fileName,
                            topLeftCell: topLeftCell
                        });
                        resolve();
                    }
                };

                reader.onerror = () => {
                    this.loggingService.logError(new Error('File reading error'), 'file_reader_error', 'SupplierAnalysisInputsComponent');
                    resolve();
                };

                reader.readAsArrayBuffer(fileInfo.file);
            });
        } finally {
            this.isProcessing = false;
        }
    }

    removeFile(fileName: string, category: 'Invoice' | 'Supplier Quotations'): void {
        if (category === 'Invoice') {
            this.invoiceFiles = this.invoiceFiles.filter(f => f.fileName !== fileName);
        } else {
            this.supplierQuotationFiles = this.supplierQuotationFiles.filter(f => f.fileName !== fileName);
        }
        this.updateServiceFiles();
    }

    private updateServiceFiles(): void {
        const allFiles = this.getAllFiles();
        this.supplierAnalysisService.setFiles(allFiles);
    }

    getAllFiles(): SupplierAnalysisFileInfo[] {
        // Always return Invoice files first, then Supplier Quotation files
        // Sort explicitly to ensure Invoice files always appear first
        const allFiles = [...this.invoiceFiles, ...this.supplierQuotationFiles];
        return allFiles.sort((a, b) => {
            // Invoice files come first (return -1), Supplier Quotations come after (return 1)
            if (a.category === 'Invoice' && b.category === 'Supplier Quotations') return -1;
            if (a.category === 'Supplier Quotations' && b.category === 'Invoice') return 1;
            return 0; // Same category, maintain order
        });
    }

    getRowCountsMatch(): boolean {
        const allFiles = this.getAllFiles();
        if (allFiles.length === 0) {
            return false;
        }
        const firstRowCount = allFiles[0].rowCount;
        return allFiles.every(file => file.rowCount === firstRowCount);
    }

    async openFile(fileInfo: SupplierAnalysisFileInfo): Promise<void> {
        this.loggingService.logUserAction('file_opened', {
            fileName: fileInfo.fileName,
            fileSize: fileInfo.file.size,
            category: fileInfo.category
        }, 'SupplierAnalysisInputsComponent');

        // Create a URL for the file and open it in a new tab
        const url = URL.createObjectURL(fileInfo.file);
        window.open(url, '_blank');

        // Clean up the URL after a short delay to free memory
        setTimeout(() => {
            URL.revokeObjectURL(url);
        }, 1000);
    }
}

