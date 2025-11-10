import { Component, OnInit } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { DataService, SupplierFileInfo } from '../../services/data.service';
import { LoggingService } from '../../services/logging.service';
import * as XLSX from 'xlsx';

interface SortState {
    column: string;
    direction: 'asc' | 'desc';
}

@Component({
    selector: 'app-suppliers-docs',
    standalone: true,
    imports: [CommonModule, FormsModule],
    templateUrl: './suppliers-docs.component.html',
    styleUrls: ['./suppliers-docs.component.scss']
})
export class SuppliersDocsComponent implements OnInit {
    supplierFiles: SupplierFileInfo[] = [];
    isDragging = false;
    isProcessing = false;
    sortState: SortState = { column: '', direction: 'asc' };
    showConfirmDialog = false;
    hoveredDropzone: string | null = null;
    priceDivider: number = 0.9;
    hasSupplierFiles = false;
    buttonDisabled = true; // Always disabled since we removed the process button
    hasNewFiles = false;
    priceDividerChanged = false;
    initialPriceDivider: number = 0.9;
    separateFreshProvisions: boolean = false; // Default to "Do not Separate"
    previewDialogVisible = false;
    previewDialogHeaders: string[] = [];
    previewDialogRows: string[][] = [];
    previewDialogFileName = '';
    previewTopLeftHeader = '';
    previewDialogHighlightIndexes: number[] = [];

    constructor(private dataService: DataService, private loggingService: LoggingService) { }

    ngOnInit(): void {
        this.hasSupplierFiles = this.dataService.hasSupplierFiles();

        // Load price divider from data service
        this.priceDivider = this.dataService.getPriceDivider();
        this.initialPriceDivider = this.priceDivider;

        // Load separate fresh provisions setting from data service
        this.separateFreshProvisions = this.dataService.getSeparateFreshProvisions();

        this.dataService.supplierFiles$.subscribe(files => {
            this.supplierFiles = files;
            this.hasSupplierFiles = this.dataService.hasSupplierFiles();

            // Re-apply sort if user had selected a sort column
            if (this.sortState.column) {
                this.sortData(this.sortState.column);
            }

            this.loggingService.logDataProcessing('files_updated', {
                fileCount: files.length,
                hasData: files.some(f => f.hasData === true)
            }, 'SuppliersDocsComponent');
        });

        // Initialize button state
        this.updateButtonState();
    }

    onDragOver(event: DragEvent, category?: string): void {
        event.preventDefault();
        event.stopPropagation();
        this.isDragging = true;
        this.hoveredDropzone = category || null;
    }

    onDragLeave(event: DragEvent): void {
        event.preventDefault();
        event.stopPropagation();
        this.isDragging = false;
        this.hoveredDropzone = null;

        this.loggingService.logUserAction('drag_leave', {
            category: this.hoveredDropzone || 'unknown'
        }, 'SuppliersDocsComponent');
    }

    async onDrop(event: DragEvent, category?: string): Promise<void> {
        event.preventDefault();
        event.stopPropagation();
        this.isDragging = false;
        this.hoveredDropzone = null;

        const files = event.dataTransfer?.files;
        if (files) {
            this.loggingService.logUserAction('files_dropped', {
                category: category || 'unknown',
                filesCount: files.length,
                fileNames: Array.from(files).map(f => f.name)
            }, 'SuppliersDocsComponent');

            await this.processFiles(files, category);
        }
    }

    onFileSelect(event: Event, category?: string): void {
        const input = event.target as HTMLInputElement;
        if (input.files) {
            this.loggingService.logUserAction('files_selected', {
                category: category || 'unknown',
                filesCount: input.files.length,
                fileNames: Array.from(input.files).map(f => f.name)
            }, 'SuppliersDocsComponent');

            this.processFiles(input.files, category);
        }
    }

    async processFiles(fileList: FileList, category?: string): Promise<void> {
        this.isProcessing = true;
        const files = Array.from(fileList).filter(f =>
            f.name.toLowerCase().endsWith('.xlsx') || f.name.toLowerCase().endsWith('.xls')
        );

        if (files.length > 0) {
            // Log file upload details
            files.forEach(file => {
                this.loggingService.logFileUpload(
                    file.name,
                    file.size,
                    file.type,
                    category || 'unknown',
                    'SuppliersDocsComponent'
                );
            });

            try {
                await this.dataService.addSupplierFiles(files, category);
                // Automatically process files after adding them
                await this.dataService.processSupplierFiles();

                this.loggingService.logDataProcessing('files_processed_successfully', {
                    fileCount: files.length,
                    category: category || 'unknown',
                    fileNames: files.map(f => f.name)
                }, 'SuppliersDocsComponent');

                this.hasNewFiles = false; // Reset since we processed immediately
                this.priceDividerChanged = false; // Reset since we processed immediately
                this.initialPriceDivider = this.priceDivider; // Update initial value
                this.updateButtonState();
            } catch (error) {
                this.loggingService.logError(error as Error, 'file_processing', 'SuppliersDocsComponent', {
                    fileCount: files.length,
                    category: category || 'unknown'
                });
            }
        } else {
            this.loggingService.logUserAction('invalid_files_selected', {
                totalFiles: fileList.length,
                validFiles: 0,
                category: category || 'unknown'
            }, 'SuppliersDocsComponent');
        }
        this.isProcessing = false;
    }

    triggerFileInput(category?: string): void {
        this.loggingService.logButtonClick('file_input_triggered', 'SuppliersDocsComponent', {
            category: category || 'unknown'
        });

        const fileInput = document.getElementById(category ? `fileInput-${category}` : 'fileInput') as HTMLInputElement;
        fileInput?.click();
    }

    clearAllFiles(): void {
        this.loggingService.logButtonClick('clear_all_files_requested', 'SuppliersDocsComponent', {
            currentFileCount: this.supplierFiles.length
        });

        this.showConfirmDialog = true;
    }

    confirmClearAll(): void {
        this.loggingService.logUserAction('clear_all_files_confirmed', {
            filesCleared: this.supplierFiles.length
        }, 'SuppliersDocsComponent');

        this.dataService.clearAll();
        this.showConfirmDialog = false;

        // Reset flags when all files are cleared
        this.hasNewFiles = false;
        this.priceDividerChanged = false;
        this.updateButtonState();
    }

    cancelClearAll(): void {
        this.loggingService.logButtonClick('clear_all_files_cancelled', 'SuppliersDocsComponent');
        this.showConfirmDialog = false;
    }

    sortData(column: string): void {
        const previousDirection = this.sortState.direction;

        if (this.sortState.column === column) {
            this.sortState.direction = this.sortState.direction === 'asc' ? 'desc' : 'asc';
        } else {
            this.sortState.column = column;
            this.sortState.direction = 'asc';
        }

        this.loggingService.logSortChange(column, this.sortState.direction, 'SuppliersDocsComponent');

        this.supplierFiles.sort((a, b) => {
            let aValue: any;
            let bValue: any;

            switch (column) {
                case 'fileName':
                    aValue = a.fileName.toLowerCase();
                    bValue = b.fileName.toLowerCase();
                    break;
                case 'topLeftCell':
                    aValue = a.topLeftCell.toLowerCase();
                    bValue = b.topLeftCell.toLowerCase();
                    break;
                case 'descriptionColumn':
                    aValue = (a.descriptionHeader || a.descriptionColumn).toLowerCase();
                    bValue = (b.descriptionHeader || b.descriptionColumn).toLowerCase();
                    break;
                case 'priceColumn':
                    aValue = (a.priceHeader || a.priceColumn).toLowerCase();
                    bValue = (b.priceHeader || b.priceColumn).toLowerCase();
                    break;
                case 'unitColumn':
                    aValue = (a.unitHeader || a.unitColumn).toLowerCase();
                    bValue = (b.unitHeader || b.unitColumn).toLowerCase();
                    break;
                case 'remarksColumn':
                    aValue = (a.remarksHeader || a.remarksColumn).toLowerCase();
                    bValue = (b.remarksHeader || b.remarksColumn).toLowerCase();
                    break;
                case 'rowCount':
                    aValue = a.rowCount;
                    bValue = b.rowCount;
                    break;
                case 'category':
                    aValue = (a.category || 'N/A').toLowerCase();
                    bValue = (b.category || 'N/A').toLowerCase();
                    break;
                case 'status':
                    // Sort by status: success (true) > pending (undefined) > error (false)
                    let statusA = 0; // error (false)
                    let statusB = 0; // error (false)

                    if (a.hasData === true) statusA = 2; // success
                    else if (a.hasData === undefined) statusA = 1; // pending

                    if (b.hasData === true) statusB = 2; // success
                    else if (b.hasData === undefined) statusB = 1; // pending

                    aValue = statusA;
                    bValue = statusB;
                    break;
                default:
                    return 0;
            }

            if (aValue < bValue) {
                return this.sortState.direction === 'asc' ? -1 : 1;
            }
            if (aValue > bValue) {
                return this.sortState.direction === 'asc' ? 1 : -1;
            }
            return 0;
        });
    }

    getSortIcon(column: string): string {
        if (this.sortState.column !== column) {
            return '↕️';
        }
        return this.sortState.direction === 'asc' ? '↑' : '↓';
    }

    canPreviewFile(fileInfo: SupplierFileInfo): boolean {
        return !!fileInfo.file && !!fileInfo.topLeftCell && fileInfo.topLeftCell !== 'NOT FOUND' && fileInfo.topLeftCell.trim() !== '';
    }

    async openFile(fileInfo: SupplierFileInfo): Promise<void> {
        this.loggingService.logUserAction('file_opened', {
            fileName: fileInfo.fileName,
            fileSize: fileInfo.file.size,
            category: fileInfo.category
        }, 'SuppliersDocsComponent');

        // Create a URL for the file and open it in a new tab
        const url = URL.createObjectURL(fileInfo.file);
        window.open(url, '_blank');

        // Clean up the URL after a short delay to free memory
        setTimeout(() => {
            URL.revokeObjectURL(url);
        }, 1000);
    }

    async viewFilePreview(fileInfo: SupplierFileInfo): Promise<void> {
        if (!this.canPreviewFile(fileInfo)) {
            return;
        }

        this.loggingService.logButtonClick('supplier_file_preview_requested', 'SuppliersDocsComponent', {
            fileName: fileInfo.fileName,
            category: fileInfo.category
        });

        let headers: string[] = [];
        let rows: string[][] = [];

        const processedRows = this.dataService.getProcessedData()
            .filter(row => row.fileName === fileInfo.fileName);

        if (processedRows.length > 0 && processedRows[0].originalHeaders && processedRows[0].originalHeaders.length > 0) {
            headers = processedRows[0].originalHeaders.map(header => header != null ? String(header) : '');
            rows = processedRows.slice(0, 4).map(row => (row.originalData || []).map(cell => cell != null ? String(cell) : ''));
        } else {
            const fallbackPreview = await this.buildPreviewDirectlyFromFile(fileInfo);
            headers = fallbackPreview.headers;
            rows = fallbackPreview.rows;

            if (headers.length === 0 || rows.length === 0) {
                this.loggingService.logUserAction('supplier_file_preview_unavailable', {
                    fileName: fileInfo.fileName
                }, 'SuppliersDocsComponent');
            }
        }

        const { trimmedHeaders, trimmedRows } = this.trimPreviewData(headers, rows);

        this.previewDialogHeaders = trimmedHeaders;
        this.previewDialogRows = trimmedRows;
        this.previewTopLeftHeader = trimmedHeaders.length > 0 ? trimmedHeaders[0] : '';
        this.previewDialogFileName = fileInfo.fileName;
        this.previewDialogHighlightIndexes = this.getPreviewHighlightIndexes(fileInfo, trimmedHeaders.length);
        this.previewDialogVisible = true;
    }

    closePreviewDialog(): void {
        this.previewDialogVisible = false;
        this.previewDialogHighlightIndexes = [];
    }

    onPriceDividerChange(): void {
        this.loggingService.logFormSubmission('price_divider_change', {
            newValue: this.priceDivider,
            previousValue: this.initialPriceDivider
        }, 'SuppliersDocsComponent');

        // Update price divider in data service (this will automatically reprocess)
        this.dataService.setPriceDivider(this.priceDivider);

        // Reset flags since processing happens automatically
        this.hasNewFiles = false;
        this.priceDividerChanged = false;
        this.initialPriceDivider = this.priceDivider;
        this.updateButtonState();
    }

    updateButtonState(): void {
        // Button is always disabled since we removed the process button
        this.buttonDisabled = true;
    }

    onSeparateFreshProvisionsChange(): void {
        this.loggingService.logFormSubmission('separate_fresh_provisions_change', {
            newValue: this.separateFreshProvisions,
            previousValue: !this.separateFreshProvisions
        }, 'SuppliersDocsComponent');

        // Update separate fresh provisions setting in data service
        this.dataService.setSeparateFreshProvisions(this.separateFreshProvisions);
    }

    isPreviewColumnHighlighted(index: number): boolean {
        return this.previewDialogHighlightIndexes.includes(index);
    }

    private trimPreviewData(headers: string[], rows: string[][]): { trimmedHeaders: string[]; trimmedRows: string[][] } {
        const columnCount = Math.max(headers.length, ...rows.map(row => row.length), 0);
        let lastIndexWithData = columnCount - 1;

        while (lastIndexWithData >= 0) {
            const headerValue = (headers[lastIndexWithData] || '').trim();
            const hasDataInColumn = rows.some(row => (row[lastIndexWithData] || '').trim() !== '');

            if (headerValue !== '' || hasDataInColumn) {
                break;
            }

            lastIndexWithData--;
        }

        if (lastIndexWithData < 0) {
            return {
                trimmedHeaders: [],
                trimmedRows: []
            };
        }

        const trimmedHeaders = headers.slice(0, lastIndexWithData + 1);
        const trimmedRows = rows.map(row => row.slice(0, lastIndexWithData + 1));

        return { trimmedHeaders, trimmedRows };
    }

    private getPreviewHighlightIndexes(fileInfo: SupplierFileInfo, columnCount: number): number[] {
        const indexes = new Set<number>();

        if (!fileInfo.topLeftCell || fileInfo.topLeftCell === 'NOT FOUND') {
            return [];
        }

        const baseColumn = XLSX.utils.decode_cell(fileInfo.topLeftCell).c;
        const addIndex = (columnRef?: string, headerValue?: string) => {
            if (!columnRef || columnRef === 'NOT FOUND') {
                if (!headerValue) {
                    return;
                }
                const matchIndex = this.previewDialogHeaders.findIndex(header =>
                    header && header.trim().toLowerCase() === headerValue.trim().toLowerCase()
                );
                if (matchIndex >= 0 && matchIndex < columnCount) {
                    indexes.add(matchIndex);
                }
                return;
            }

            try {
                const columnIndex = XLSX.utils.decode_col(columnRef) - baseColumn;
                if (columnIndex >= 0 && columnIndex < columnCount) {
                    indexes.add(columnIndex);
                }
            } catch {
                // Fallback to header text match if decoding fails
                if (headerValue) {
                    const fallbackIndex = this.previewDialogHeaders.findIndex(header =>
                        header && header.trim().toLowerCase() === headerValue.trim().toLowerCase()
                    );
                    if (fallbackIndex >= 0 && fallbackIndex < columnCount) {
                        indexes.add(fallbackIndex);
                    }
                }
            }
        };

        addIndex(fileInfo.descriptionColumn, fileInfo.descriptionHeader);
        addIndex(fileInfo.priceColumn, fileInfo.priceHeader);
        addIndex(fileInfo.unitColumn, fileInfo.unitHeader);
        addIndex(fileInfo.remarksColumn, fileInfo.remarksHeader);

        return Array.from(indexes.values()).sort((a, b) => a - b);
    }

    private buildPreviewDirectlyFromFile(fileInfo: SupplierFileInfo): Promise<{ headers: string[]; rows: string[][] }> {
        return new Promise((resolve) => {
            const reader = new FileReader();

            reader.onload = (event) => {
                try {
                    const data = new Uint8Array(event.target?.result as ArrayBuffer);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const sheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[sheetName];

                    if (!worksheet) {
                        resolve({ headers: [], rows: [] });
                        return;
                    }

                    resolve(this.extractPreviewFromWorksheet(worksheet, fileInfo));
                } catch (error) {
                    this.loggingService.logError(error as Error, 'supplier_file_preview_read_error', 'SuppliersDocsComponent', {
                        fileName: fileInfo.fileName
                    });
                    resolve({ headers: [], rows: [] });
                }
            };

            reader.onerror = () => {
                const readerError = reader.error?.message || 'Unknown file reader error';
                this.loggingService.logError(new Error(readerError), 'supplier_file_preview_file_reader_error', 'SuppliersDocsComponent', {
                    fileName: fileInfo.fileName
                });
                resolve({ headers: [], rows: [] });
            };

            reader.readAsArrayBuffer(fileInfo.file);
        });
    }

    private extractPreviewFromWorksheet(worksheet: XLSX.WorkSheet, fileInfo: SupplierFileInfo): { headers: string[]; rows: string[][] } {
        const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1:A1');
        const topLeft = XLSX.utils.decode_cell(fileInfo.topLeftCell);

        const headers: string[] = [];
        for (let col = topLeft.c; col <= range.e.c; col++) {
            const cellAddress = XLSX.utils.encode_cell({ r: topLeft.r, c: col });
            const cell = worksheet[cellAddress];
            headers.push(cell && cell.v != null ? String(cell.v) : '');
        }

        const rows: string[][] = [];
        for (let row = topLeft.r + 1; row <= range.e.r && rows.length < 4; row++) {
            const rowData: string[] = [];
            let hasData = false;

            for (let col = topLeft.c; col <= range.e.c; col++) {
                const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
                const cell = worksheet[cellAddress];
                const value = cell && cell.v != null ? String(cell.v) : '';

                if (value.trim() !== '') {
                    hasData = true;
                }

                rowData.push(value);
            }

            if (!hasData) {
                continue;
            }

            rows.push(rowData);
        }

        return { headers, rows };
    }

}
