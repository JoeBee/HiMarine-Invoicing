import { Component, OnInit } from '@angular/core';
import { CommonModule } from '@angular/common';
import { DataService, SupplierFileInfo } from '../../services/data.service';

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
            f.name.endsWith('.xlsx') || f.name.endsWith('.xls')
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
}

