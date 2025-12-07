import { Component, OnInit } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { DataService } from '../../services/data.service';

@Component({
    selector: 'app-suppliers',
    standalone: true,
    imports: [CommonModule, FormsModule],
    templateUrl: './suppliers.component.html',
    styleUrls: ['./suppliers.component.scss']
})
export class SuppliersComponent implements OnInit {
    priceDivider: number = 0.9;
    hasSupplierFiles = false;
    isProcessing = false;
    buttonDisabled = false;

    constructor(private dataService: DataService) { }

    ngOnInit(): void {
        this.hasSupplierFiles = this.dataService.hasSupplierFiles();

        // Load price divider from data service
        this.priceDivider = this.dataService.getPriceDivider();

        // Subscribe to supplier files changes
        this.dataService.supplierFiles$.subscribe(files => {
            this.hasSupplierFiles = this.dataService.hasSupplierFiles();
        });
    }

    onPriceDividerChange(): void {
        // Update price divider in data service
        this.dataService.setPriceDivider(this.priceDivider);
    }

    async processSupplierFiles(): Promise<void> {
        this.isProcessing = true;
        await this.dataService.processSupplierFiles();
        this.isProcessing = false;
        this.buttonDisabled = true; // Disable button after processing
    }
}
