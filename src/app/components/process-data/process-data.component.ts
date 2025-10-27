import { Component, OnInit, OnDestroy } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { DataService, ProcessedDataRow } from '../../services/data.service';
import { LoggingService } from '../../services/logging.service';

interface SortState {
    column: string;
    direction: 'asc' | 'desc';
}

@Component({
    selector: 'app-process-data',
    standalone: true,
    imports: [CommonModule, FormsModule],
    templateUrl: './process-data.component.html',
    styleUrls: ['./process-data.component.scss']
})
export class ProcessDataComponent implements OnInit, OnDestroy {
    processedData: ProcessedDataRow[] = [];
    filteredData: ProcessedDataRow[] = [];
    hasSupplierFiles = false;
    isProcessing = false;
    hasProcessedFiles = false; // Track if files have been processed
    previousFileCount = 0; // Track previous file count to detect new files
    sortState: SortState = { column: '', direction: 'asc' };

    // Filter properties
    selectedFileName = '';
    selectedDescription = '';
    descriptionTextFilter = '';
    availableFileNames: string[] = [];
    commonDescriptions = ['Beer', 'Cheese', 'Ice', 'Provision'];

    // Row expansion properties
    expandedRowIndex: number | null = null;

    constructor(private dataService: DataService, private loggingService: LoggingService) { }

    ngOnInit(): void {
        this.loggingService.logSystemEvent('component_initialized', {
            component: 'ProcessDataComponent',
            timestamp: new Date().toISOString()
        }, 'ProcessDataComponent');

        this.hasSupplierFiles = this.dataService.hasSupplierFiles();

        this.dataService.supplierFiles$.subscribe((files) => {
            const currentFileCount = files.length;
            this.hasSupplierFiles = this.dataService.hasSupplierFiles();

            // If new files are added (file count increased) and we had previously processed files, show the button again
            if (currentFileCount > this.previousFileCount && this.hasProcessedFiles) {
                this.hasProcessedFiles = false;
            }

            this.previousFileCount = currentFileCount;

            this.loggingService.logDataProcessing('files_updated', {
                fileCount: files.length,
                hasProcessedFiles: this.hasProcessedFiles
            }, 'ProcessDataComponent');
        });

        this.dataService.processedData$.subscribe(data => {
            this.processedData = data;
            this.filteredData = [...data]; // Initialize filtered data
            this.updateAvailableFileNames();
            this.updateCommonDescriptions();
            this.applyFilters();

            this.loggingService.logDataProcessing('processed_data_updated', {
                totalRecords: data.length,
                filteredRecords: this.filteredData.length
            }, 'ProcessDataComponent');
        });

        // Add document click listener to close expanded rows when clicking outside
        document.addEventListener('click', this.onDocumentClick.bind(this));
    }

    async processSupplierFiles(): Promise<void> {
        this.loggingService.logButtonClick('process_supplier_files', 'ProcessDataComponent', {
            fileCount: this.dataService.hasSupplierFiles() ? 1 : 0
        });

        this.isProcessing = true;

        try {
            await this.dataService.processSupplierFiles();
            this.isProcessing = false;
            this.hasProcessedFiles = true; // Mark that files have been processed

            this.loggingService.logDataProcessing('files_processed_successfully', {
                processedFiles: this.dataService.hasSupplierFiles() ? 1 : 0
            }, 'ProcessDataComponent');
        } catch (error) {
            this.isProcessing = false;
            this.loggingService.logError(error as Error, 'file_processing', 'ProcessDataComponent');
        }
    }

    onCountChange(index: number, event: Event): void {
        const select = event.target as HTMLSelectElement;
        const count = parseInt(select.value, 10);

        // Find the original index in processedData
        const filteredRow = this.filteredData[index];
        const originalIndex = this.processedData.findIndex(row =>
            row.fileName === filteredRow.fileName &&
            row.description === filteredRow.description &&
            row.price === filteredRow.price &&
            row.unit === filteredRow.unit
        );

        if (originalIndex !== -1) {
            this.dataService.updateRowCount(originalIndex, count);

            this.loggingService.logUserAction('item_count_changed', {
                itemDescription: filteredRow.description,
                previousCount: filteredRow.count,
                newCount: count,
                fileName: filteredRow.fileName
            }, 'ProcessDataComponent');
        }
    }

    sortData(column: string): void {
        if (this.sortState.column === column) {
            this.sortState.direction = this.sortState.direction === 'asc' ? 'desc' : 'asc';
        } else {
            this.sortState.column = column;
            this.sortState.direction = 'asc';
        }

        this.filteredData.sort((a, b) => {
            let aValue: any;
            let bValue: any;

            switch (column) {
                case 'fileName':
                    aValue = a.fileName.toLowerCase();
                    bValue = b.fileName.toLowerCase();
                    break;
                case 'description':
                    aValue = a.description.toLowerCase();
                    bValue = b.description.toLowerCase();
                    break;
                case 'unit':
                    aValue = a.unit.toLowerCase();
                    bValue = b.unit.toLowerCase();
                    break;
                case 'remarks':
                    aValue = a.remarks.toLowerCase();
                    bValue = b.remarks.toLowerCase();
                    break;
                case 'price':
                    aValue = a.price;
                    bValue = b.price;
                    break;
                case 'count':
                    aValue = a.count;
                    bValue = b.count;
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

    updateAvailableFileNames(): void {
        const uniqueFileNames = [...new Set(this.processedData.map(row => row.fileName))];
        this.availableFileNames = uniqueFileNames.sort();
    }

    updateCommonDescriptions(): void {
        // Keep the predefined list: Beer, Cheese, Ice, Provision
        // No need to dynamically update since we want only these specific options
        this.commonDescriptions = ['Beer', 'Cheese', 'Ice', 'Provision'];
    }

    onFileNameFilterChange(): void {
        this.applyFilters();
    }

    onDescriptionFilterChange(): void {
        this.applyFilters();
    }

    onDescriptionTextFilterChange(): void {
        this.applyFilters();
    }

    applyFilters(): void {
        let filtered = [...this.processedData];

        // Filter by file name
        if (this.selectedFileName) {
            filtered = filtered.filter(row => row.fileName === this.selectedFileName);
        }

        // Filter by description (searches both Description and File Name columns)
        if (this.selectedDescription) {
            filtered = filtered.filter(row =>
                row.description.toLowerCase().includes(this.selectedDescription.toLowerCase()) ||
                row.fileName.toLowerCase().includes(this.selectedDescription.toLowerCase())
            );
        }

        // Filter by description text
        if (this.descriptionTextFilter.trim()) {
            const searchText = this.descriptionTextFilter.toLowerCase().trim();
            filtered = filtered.filter(row =>
                row.description.toLowerCase().includes(searchText)
            );
        }

        this.filteredData = filtered;
    }

    clearFilters(): void {
        this.selectedFileName = '';
        this.selectedDescription = '';
        this.descriptionTextFilter = '';
        this.applyFilters();
    }

    clearFileNameFilter(): void {
        this.selectedFileName = '';
        this.applyFilters();
    }

    clearDescriptionFilter(): void {
        this.selectedDescription = '';
        this.applyFilters();
    }

    toggleRowExpansion(index: number): void {
        if (this.expandedRowIndex === index) {
            this.expandedRowIndex = null;
        } else {
            this.expandedRowIndex = index;
        }
    }

    isRowExpanded(index: number): boolean {
        return this.expandedRowIndex === index;
    }

    onRowClick(index: number): void {
        this.toggleRowExpansion(index);
    }

    onDocumentClick(event: Event): void {
        // Close expanded row when clicking outside
        const target = event.target as HTMLElement;
        if (!target.closest('.data-table')) {
            this.expandedRowIndex = null;
        }
    }

    ngOnDestroy(): void {
        // Remove event listener when component is destroyed
        document.removeEventListener('click', this.onDocumentClick.bind(this));
    }
}

