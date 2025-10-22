import { Component, OnInit } from '@angular/core';
import { CommonModule } from '@angular/common';
import { DataService, SupplierFileInfo } from '../../services/data.service';

interface SortState {
    column: string;
    direction: 'asc' | 'desc';
}

@Component({
    selector: 'app-suppliers',
    standalone: true,
    imports: [CommonModule],
    templateUrl: './suppliers.component.html',
    styleUrls: ['./suppliers.component.scss']
})
export class SuppliersComponent implements OnInit {
    supplierFiles: SupplierFileInfo[] = [];
    isDragging = false;
    isProcessing = false;
    sortState: SortState = { column: '', direction: 'asc' };

    constructor(private dataService: DataService) { }

    ngOnInit(): void {
        this.dataService.supplierFiles$.subscribe(files => {
            this.supplierFiles = files;
        });
    }

    onDragOver(event: DragEvent): void {
        event.preventDefault();
        event.stopPropagation();
        this.isDragging = true;
    }

    onDragLeave(event: DragEvent): void {
        event.preventDefault();
        event.stopPropagation();
        this.isDragging = false;
    }

    async onDrop(event: DragEvent): Promise<void> {
        event.preventDefault();
        event.stopPropagation();
        this.isDragging = false;

        const files = event.dataTransfer?.files;
        if (files) {
            await this.processFiles(files);
        }
    }

    onFileSelect(event: Event): void {
        const input = event.target as HTMLInputElement;
        if (input.files) {
            this.processFiles(input.files);
        }
    }

    async processFiles(fileList: FileList): Promise<void> {
        this.isProcessing = true;
        const files = Array.from(fileList).filter(f =>
            f.name.toLowerCase().endsWith('.xlsx') || f.name.toLowerCase().endsWith('.xls')
        );

        if (files.length > 0) {
            await this.dataService.addSupplierFiles(files);
        }
        this.isProcessing = false;
    }

    triggerFileInput(): void {
        const fileInput = document.getElementById('fileInput') as HTMLInputElement;
        fileInput?.click();
    }

    clearAllFiles(): void {
        if (confirm('Are you sure you want to clear all uploaded files? This will also clear all processed data.')) {
            this.dataService.clearAll();
        }
    }

    sortData(column: string): void {
        if (this.sortState.column === column) {
            this.sortState.direction = this.sortState.direction === 'asc' ? 'desc' : 'asc';
        } else {
            this.sortState.column = column;
            this.sortState.direction = 'asc';
        }

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
                case 'rowCount':
                    aValue = a.rowCount;
                    bValue = b.rowCount;
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
}

