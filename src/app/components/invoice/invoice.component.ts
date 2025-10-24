import { Component, OnInit } from '@angular/core';
import { CommonModule } from '@angular/common';
import { DataService, ProcessedDataRow } from '../../services/data.service';
import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';
import jsPDF from 'jspdf';

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

    get finalTotalAmount(): number {
        return this.totalAmount * 1.10 / 0.9;
    }

    generateInvoice(): void {
        const includedData = this.processedData.filter(row => row.count > 0);

        if (includedData.length === 0) {
            alert('No items selected for invoice. Please set a count greater than 0 for items you want to invoice.');
            return;
        }

        // Prepare data for Excel
        const worksheetData = [
            ['File Name', 'Description', 'Count', 'Remarks', 'Unit Price', 'Total Price'],
            ...includedData.map(row => {
                const adjustedPrice = row.price * 1.10 / 0.9;
                return [
                    row.fileName,
                    row.description,
                    row.count,
                    row.remarks,
                    adjustedPrice,
                    row.count * adjustedPrice
                ];
            }),
            ['', '', '', '', 'TOTAL AMOUNT:', this.finalTotalAmount]
        ];

        // Create workbook and worksheet
        const worksheet = XLSX.utils.aoa_to_sheet(worksheetData);

        // Set column widths
        worksheet['!cols'] = [
            { wch: 30 },
            { wch: 50 },
            { wch: 10 },
            { wch: 30 },
            { wch: 15 },
            { wch: 15 }
        ];

        // Format price columns as currency
        const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1:F1');
        for (let row = 1; row <= range.e.r; row++) {
            // Format unit price column (column 4, index 4)
            const unitPriceAddress = XLSX.utils.encode_cell({ r: row, c: 4 });
            if (worksheet[unitPriceAddress]) {
                worksheet[unitPriceAddress].z = '"$"#,##0.00';
            }
            // Format total price column (column 5, index 5)
            const totalPriceAddress = XLSX.utils.encode_cell({ r: row, c: 5 });
            if (worksheet[totalPriceAddress]) {
                worksheet[totalPriceAddress].z = '"$"#,##0.00';
            }
        }

        // Format the total amount row as bold and currency
        const totalRowIndex = range.e.r;
        const totalLabelAddress = XLSX.utils.encode_cell({ r: totalRowIndex, c: 4 });
        const totalAmountAddress = XLSX.utils.encode_cell({ r: totalRowIndex, c: 5 });

        if (worksheet[totalLabelAddress]) {
            worksheet[totalLabelAddress].s = { font: { bold: true } };
        }
        if (worksheet[totalAmountAddress]) {
            worksheet[totalAmountAddress].z = '"$"#,##0.00';
            worksheet[totalAmountAddress].s = { font: { bold: true } };
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

    generateInvoicePDF(): void {
        const includedData = this.processedData.filter(row => row.count > 0);

        if (includedData.length === 0) {
            alert('No items selected for invoice. Please set a count greater than 0 for items you want to invoice.');
            return;
        }

        // Create new PDF document
        const doc = new jsPDF();

        // Set up the document
        const pageWidth = doc.internal.pageSize.getWidth();
        const margin = 20;
        const contentWidth = pageWidth - (margin * 2);

        // Add title
        doc.setFontSize(24);
        doc.setFont('helvetica', 'bold');
        doc.text('INVOICE', pageWidth / 2, 30, { align: 'center' });

        // Add date
        doc.setFontSize(12);
        doc.setFont('helvetica', 'normal');
        const currentDate = new Date().toLocaleDateString();
        doc.text(`Date: ${currentDate}`, margin, 50);

        // Add invoice details
        doc.setFontSize(14);
        doc.setFont('helvetica', 'bold');
        doc.text('Invoice Details', margin, 70);

        // Add summary information
        doc.setFontSize(10);
        doc.setFont('helvetica', 'normal');
        doc.text(`Total Items: ${this.includedItemsCount}`, margin, 85);
        doc.text(`Total Amount: ${this.finalTotalAmount.toFixed(2)}`, margin, 95);

        // Create table headers
        const tableTop = 110;
        const availableWidth = pageWidth - (margin * 2);
        const colWidths = [availableWidth * 0.20, availableWidth * 0.25, availableWidth * 0.10, availableWidth * 0.20, availableWidth * 0.125, availableWidth * 0.125];
        const colPositions = [margin];
        for (let i = 1; i < colWidths.length; i++) {
            colPositions.push(colPositions[i - 1] + colWidths[i - 1]);
        }

        // Table headers
        doc.setFontSize(10);
        doc.setFont('helvetica', 'bold');
        const headers = ['File Name', 'Description', 'Count', 'Remarks', 'Unit Price', 'Total'];
        headers.forEach((header, index) => {
            doc.text(header, colPositions[index], tableTop);
        });

        // Draw header line
        doc.line(margin, tableTop + 5, pageWidth - margin, tableTop + 5);

        // Add table rows
        doc.setFont('helvetica', 'normal');
        let currentY = tableTop + 15;

        includedData.forEach((row, index) => {
            // Check if we need a new page
            if (currentY > doc.internal.pageSize.getHeight() - 30) {
                doc.addPage();
                currentY = 20;
            }

            // File name (truncate if too long)
            const maxFileNameLength = Math.floor(colWidths[0] / 3); // Approximate characters per unit width
            const fileName = row.fileName.length > maxFileNameLength ? row.fileName.substring(0, maxFileNameLength - 3) + '...' : row.fileName;
            doc.text(fileName, colPositions[0], currentY);

            // Description (truncate if too long)
            const maxDescLength = Math.floor(colWidths[1] / 3); // Approximate characters per unit width
            const description = row.description.length > maxDescLength ? row.description.substring(0, maxDescLength - 3) + '...' : row.description;
            doc.text(description, colPositions[1], currentY);

            // Count
            doc.text(row.count.toString(), colPositions[2], currentY);

            // Remarks (truncate if too long)
            const maxRemarksLength = Math.floor(colWidths[3] / 3); // Approximate characters per unit width
            const remarks = row.remarks.length > maxRemarksLength ? row.remarks.substring(0, maxRemarksLength - 3) + '...' : row.remarks;
            doc.text(remarks, colPositions[3], currentY);

            // Unit price (with calculation applied)
            const adjustedPrice = row.price * 1.10 / 0.9;
            doc.text(`$${adjustedPrice.toFixed(2)}`, colPositions[4], currentY);

            // Total (with calculation applied)
            const total = (row.count * adjustedPrice).toFixed(2);
            doc.text(`$${total}`, colPositions[5], currentY);

            currentY += 10;
        });

        // Add total line
        currentY += 10;
        doc.line(margin, currentY, pageWidth - margin, currentY);
        currentY += 10;

        doc.setFont('helvetica', 'bold');
        doc.setFontSize(12);
        const totalText = `TOTAL AMOUNT: $${this.finalTotalAmount.toFixed(2)}`;
        const totalTextWidth = doc.getTextWidth(totalText);
        const totalX = Math.max(colPositions[4], pageWidth - margin - totalTextWidth);
        doc.text(totalText, totalX, currentY);

        // Save the PDF
        const fileName = `Invoice_${new Date().toISOString().split('T')[0]}.pdf`;
        doc.save(fileName);
    }
}

