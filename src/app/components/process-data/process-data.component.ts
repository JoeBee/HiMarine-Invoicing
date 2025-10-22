import { Component, OnInit } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { DataService, ProcessedDataRow } from '../../services/data.service';

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
export class ProcessDataComponent implements OnInit {
    processedData: ProcessedDataRow[] = [];
    hasSupplierFiles = false;
    isProcessing = false;
    sortState: SortState = { column: '', direction: 'asc' };

    constructor(private dataService: DataService) { }

    ngOnInit(): void {
        this.hasSupplierFiles = this.dataService.hasSupplierFiles();

        this.dataService.supplierFiles$.subscribe(() => {
            this.hasSupplierFiles = this.dataService.hasSupplierFiles();
        });

        this.dataService.processedData$.subscribe(data => {
            this.processedData = data;
        });
    }

    async processSupplierFiles(): Promise<void> {
        this.isProcessing = true;
        await this.dataService.processSupplierFiles();
        this.isProcessing = false;
    }

    onCountChange(index: number, event: Event): void {
        const select = event.target as HTMLSelectElement;
        const count = parseInt(select.value, 10);
        this.dataService.updateRowCount(index, count);
    }

    sortData(column: string): void {
        if (this.sortState.column === column) {
            this.sortState.direction = this.sortState.direction === 'asc' ? 'desc' : 'asc';
        } else {
            this.sortState.column = column;
            this.sortState.direction = 'asc';
        }

        this.processedData.sort((a, b) => {
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
}

