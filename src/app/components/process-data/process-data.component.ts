import { Component, OnInit } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { DataService, ProcessedDataRow } from '../../services/data.service';

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

    onIncludeChange(index: number, event: Event): void {
        const checkbox = event.target as HTMLInputElement;
        this.dataService.updateRowInclusion(index, checkbox.checked);
    }
}

