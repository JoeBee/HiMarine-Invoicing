import { Component, OnInit } from '@angular/core';
import { CommonModule } from '@angular/common';
import { DataService, ProcessedDataRow } from '../../services/data.service';
import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';

@Component({
    selector: 'app-invoice',
    standalone: true,
    imports: [CommonModule],
    templateUrl: './invoice.component.html',
    styleUrls: ['./invoice.component.scss']
})
export class InvoiceComponent implements OnInit {
    processedData: ProcessedDataRow[] = [];
    hasDataToInvoice = false;

    constructor(private dataService: DataService) { }

    ngOnInit(): void {
        this.dataService.processedData$.subscribe(data => {
            this.processedData = data;
            this.hasDataToInvoice = data.some(row => row.include);
        });
    }

    get includedItems(): ProcessedDataRow[] {
        return this.processedData.filter(row => row.include);
    }

    get includedItemsCount(): number {
        return this.includedItems.length;
    }

    get totalAmount(): number {
        return this.includedItems.reduce((sum, row) => sum + row.price, 0);
    }

    generateInvoice(): void {
        const includedData = this.processedData.filter(row => row.include);

        if (includedData.length === 0) {
            alert('No items selected for invoice. Please check the "Include" checkbox for items you want to invoice.');
            return;
        }

        // Prepare data for Excel
        const worksheetData = [
            ['File Name', 'Description', 'Price'],
            ...includedData.map(row => [
                row.fileName,
                row.description,
                row.price
            ])
        ];

        // Create workbook and worksheet
        const worksheet = XLSX.utils.aoa_to_sheet(worksheetData);

        // Set column widths
        worksheet['!cols'] = [
            { wch: 30 },
            { wch: 50 },
            { wch: 15 }
        ];

        // Format price column as currency
        const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1:C1');
        for (let row = 1; row <= range.e.r; row++) {
            const cellAddress = XLSX.utils.encode_cell({ r: row, c: 2 });
            if (worksheet[cellAddress]) {
                worksheet[cellAddress].z = '"$"#,##0.00';
            }
        }

        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, 'Invoice');

        // Generate Excel file
        const excelBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
        const blob = new Blob([excelBuffer], {
            type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        });

        // Download file
        const fileName = `Invoice_${new Date().toISOString().split('T')[0]}.xlsx`;
        saveAs(blob, fileName);
    }
}

