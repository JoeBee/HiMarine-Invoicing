import { Component, OnInit } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { DataService, SupplierFileInfo } from '../../services/data.service';
import { LoggingService } from '../../services/logging.service';

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

    openFile(fileInfo: SupplierFileInfo): void {
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

}
