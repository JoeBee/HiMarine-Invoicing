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
            this.hasDataToInvoice = data.some(row => row.count > 0);
        });
    }

    get includedItems(): ProcessedDataRow[] {
        return this.processedData.filter(row => row.count > 0);
    }

    get includedItemsCount(): number {
        return this.includedItems.length;
    }

    get totalAmount(): number {
        return this.includedItems.reduce((sum, row) => sum + (row.count * row.price), 0);
    }

    generateInvoice(): void {
        const includedData = this.processedData.filter(row => row.count > 0);

        if (includedData.length === 0) {
            alert('No items selected for invoice. Please set a count greater than 0 for items you want to invoice.');
            return;
        }

        // Prepare data for Excel
        const worksheetData = [
            ['File Name', 'Description', 'Count', 'Unit Price', 'Total Price'],
            ...includedData.map(row => [
                row.fileName,
                row.description,
                row.count,
                row.price,
                row.count * row.price
            ])
        ];

        // Create workbook and worksheet
        const worksheet = XLSX.utils.aoa_to_sheet(worksheetData);

        // Set column widths
        worksheet['!cols'] = [
            { wch: 30 },
            { wch: 50 },
            { wch: 10 },
            { wch: 15 },
            { wch: 15 }
        ];

        // Format price columns as currency
        const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1:E1');
        for (let row = 1; row <= range.e.r; row++) {
            // Format unit price column (column 3, index 3)
            const unitPriceAddress = XLSX.utils.encode_cell({ r: row, c: 3 });
            if (worksheet[unitPriceAddress]) {
                worksheet[unitPriceAddress].z = '"$"#,##0.00';
            }
            // Format total price column (column 4, index 4)
            const totalPriceAddress = XLSX.utils.encode_cell({ r: row, c: 4 });
            if (worksheet[totalPriceAddress]) {
                worksheet[totalPriceAddress].z = '"$"#,##0.00';
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

