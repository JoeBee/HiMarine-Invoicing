import { Component, OnInit, OnDestroy } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { DataService, ProcessedDataRow } from '../../services/data.service';
import * as ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';

interface SortState {
    column: string;
    direction: 'asc' | 'desc';
}

@Component({
    selector: 'app-price-list',
    standalone: true,
    imports: [CommonModule, FormsModule],
    templateUrl: './price-list.component.html',
    styleUrls: ['./price-list.component.scss']
})
export class PriceListComponent implements OnInit, OnDestroy {
    // Process Data functionality
    processedData: ProcessedDataRow[] = [];
    filteredData: ProcessedDataRow[] = [];
    hasSupplierFiles = false;
    hasProcessedFiles = false; // Track if files have been processed
    previousFileCount = 0; // Track previous file count to detect new files
    isProcessing = false;
    buttonDisabled = false; // Track if button should be disabled
    sortState: SortState = { column: '', direction: 'asc' };

    // Filter properties
    selectedFileName = '';
    selectedDescription = '';
    descriptionTextFilter = '';
    availableFileNames: string[] = [];
    commonDescriptions = ['Beer', 'Cheese', 'Ice', 'Provision'];

    // Row expansion properties
    expandedRowIndex: number | null = null;

    constructor(private dataService: DataService) { }

    ngOnInit(): void {
        this.hasSupplierFiles = this.dataService.hasSupplierFiles();

        this.dataService.supplierFiles$.subscribe((files) => {
            const currentFileCount = files.length;
            this.hasSupplierFiles = this.dataService.hasSupplierFiles();

            // If new files are added (file count increased) and we had previously processed files, show the button again
            if (currentFileCount > this.previousFileCount && this.hasProcessedFiles) {
                this.hasProcessedFiles = false;
                this.buttonDisabled = false; // Re-enable button when new files are added
            }

            this.previousFileCount = currentFileCount;
        });

        this.dataService.processedData$.subscribe(data => {
            this.processedData = data;
            this.filteredData = [...data]; // Initialize filtered data
            this.updateAvailableFileNames();
            this.updateCommonDescriptions();
            this.applyFilters();
        });

        // Add document click listener to close expanded rows when clicking outside
        document.addEventListener('click', this.onDocumentClick.bind(this));
    }

    // Process Data functionality methods
    async processSupplierFiles(): Promise<void> {
        this.isProcessing = true;
        await this.dataService.processSupplierFiles();
        this.isProcessing = false;
        this.hasProcessedFiles = true; // Mark that files have been processed
        this.buttonDisabled = true; // Disable button after processing
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
                case 'category':
                    aValue = (a.category || 'N/A').toLowerCase();
                    bValue = (b.category || 'N/A').toLowerCase();
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

    async exportToExcel(): Promise<void> {
        const workbook = new ExcelJS.Workbook();

        // Create Cover Sheet
        this.createCoverSheet(workbook);

        // Create Provisions sheet
        this.createProvisionsSheet(workbook);

        // Create Fresh Provisions sheet
        this.createFreshProvisionsSheet(workbook);

        // Create Bond sheet
        this.createBondSheet(workbook);

        // Generate Excel file
        const buffer = await workbook.xlsx.writeBuffer();
        const data = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        saveAs(data, 'Marine_Provisions_Price_List.xlsx');
    }

    private createCoverSheet(workbook: ExcelJS.Workbook): void {
        const worksheet = workbook.addWorksheet('COVER SHEET');

        // Remove grid lines from the worksheet
        worksheet.properties.showGridLines = false;

        // Set column widths to show all content properly
        worksheet.getColumn('A').width = 20;  // Column A
        worksheet.getColumn('B').width = 30;  // Column B (for email and labels)
        worksheet.getColumn('C').width = 8;   // Column C (for '$')
        worksheet.getColumn('D').width = 8;   // Column D (for '-')

        // Add title (merged across A1:D1)
        worksheet.mergeCells('A1:D1');
        worksheet.getCell('A1').value = 'MARINE PROVISIONS PRICE LIST';
        worksheet.getCell('A1').font = { bold: true, size: 16 };
        worksheet.getCell('A1').alignment = { horizontal: 'center', vertical: 'middle' };

        // Add date and total items
        worksheet.getCell('A3').value = 'Generated on: ' + new Date().toLocaleDateString();
        worksheet.getCell('A5').value = 'Total Items: ' + this.processedData.length;

        // Add email
        worksheet.getCell('A9').value = 'Email:';
        worksheet.getCell('B9').value = 'office@eos-supply.co.uk';

        // Add PROVISIONS row (B14:D14) with dark blue background and white text
        const provisionsRow = worksheet.getRow(14);
        provisionsRow.getCell(2).value = 'PROVISIONS';
        provisionsRow.getCell(3).value = '$';
        provisionsRow.getCell(4).value = '-';

        // Style PROVISIONS row - dark blue background with white text
        provisionsRow.getCell(2).font = { bold: true, color: { argb: 'FFFFFFFF' } };
        provisionsRow.getCell(2).fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FF4472C4' }
        };
        provisionsRow.getCell(2).alignment = { horizontal: 'left' };
        provisionsRow.getCell(2).border = {
            top: { style: 'thin', color: { argb: 'FF000000' } },
            left: { style: 'thin', color: { argb: 'FF000000' } },
            bottom: { style: 'thin', color: { argb: 'FF000000' } },
            right: { style: 'thin', color: { argb: 'FF000000' } }
        };

        provisionsRow.getCell(3).font = { color: { argb: 'FF000000' } };
        provisionsRow.getCell(3).fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FF4472C4' }
        };
        provisionsRow.getCell(3).alignment = { horizontal: 'center' };
        provisionsRow.getCell(3).border = {
            top: { style: 'thin', color: { argb: 'FF000000' } },
            left: { style: 'thin', color: { argb: 'FF000000' } },
            bottom: { style: 'thin', color: { argb: 'FF000000' } },
            right: { style: 'thin', color: { argb: 'FF000000' } }
        };

        provisionsRow.getCell(4).font = { color: { argb: 'FF000000' } };
        provisionsRow.getCell(4).fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FF4472C4' }
        };
        provisionsRow.getCell(4).alignment = { horizontal: 'center' };
        provisionsRow.getCell(4).border = {
            top: { style: 'thin', color: { argb: 'FF000000' } },
            left: { style: 'thin', color: { argb: 'FF000000' } },
            bottom: { style: 'thin', color: { argb: 'FF000000' } },
            right: { style: 'thin', color: { argb: 'FF000000' } }
        };

        // Add FRESH PROVISIONS row (B15:D15) with light blue background and black text
        const freshProvisionsRow = worksheet.getRow(15);
        freshProvisionsRow.getCell(2).value = 'FRESH PROVISIONS';
        freshProvisionsRow.getCell(3).value = '$';
        freshProvisionsRow.getCell(4).value = '-';

        // Style FRESH PROVISIONS row - light blue background with black text
        freshProvisionsRow.getCell(2).font = { bold: true, color: { argb: 'FF000000' } };
        freshProvisionsRow.getCell(2).fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFB4C6E7' }
        };
        freshProvisionsRow.getCell(2).alignment = { horizontal: 'left' };
        freshProvisionsRow.getCell(2).border = {
            top: { style: 'thin', color: { argb: 'FF000000' } },
            left: { style: 'thin', color: { argb: 'FF000000' } },
            bottom: { style: 'thin', color: { argb: 'FF000000' } },
            right: { style: 'thin', color: { argb: 'FF000000' } }
        };

        freshProvisionsRow.getCell(3).font = { color: { argb: 'FF000000' } };
        freshProvisionsRow.getCell(3).fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFB4C6E7' }
        };
        freshProvisionsRow.getCell(3).alignment = { horizontal: 'center' };
        freshProvisionsRow.getCell(3).border = {
            top: { style: 'thin', color: { argb: 'FF000000' } },
            left: { style: 'thin', color: { argb: 'FF000000' } },
            bottom: { style: 'thin', color: { argb: 'FF000000' } },
            right: { style: 'thin', color: { argb: 'FF000000' } }
        };

        freshProvisionsRow.getCell(4).font = { color: { argb: 'FF000000' } };
        freshProvisionsRow.getCell(4).fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFB4C6E7' }
        };
        freshProvisionsRow.getCell(4).alignment = { horizontal: 'center' };
        freshProvisionsRow.getCell(4).border = {
            top: { style: 'thin', color: { argb: 'FF000000' } },
            left: { style: 'thin', color: { argb: 'FF000000' } },
            bottom: { style: 'thin', color: { argb: 'FF000000' } },
            right: { style: 'thin', color: { argb: 'FF000000' } }
        };

        // Add BOND row (B16:D16) with light blue background and black text
        const bondRow = worksheet.getRow(16);
        bondRow.getCell(2).value = 'BOND';
        bondRow.getCell(3).value = '$';
        bondRow.getCell(4).value = '-';

        // Style BOND row - light blue background with black text
        bondRow.getCell(2).font = { bold: true, color: { argb: 'FF000000' } };
        bondRow.getCell(2).fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFB4C6E7' }
        };
        bondRow.getCell(2).alignment = { horizontal: 'left' };
        bondRow.getCell(2).border = {
            top: { style: 'thin', color: { argb: 'FF000000' } },
            left: { style: 'thin', color: { argb: 'FF000000' } },
            bottom: { style: 'thin', color: { argb: 'FF000000' } },
            right: { style: 'thin', color: { argb: 'FF000000' } }
        };

        bondRow.getCell(3).font = { color: { argb: 'FF000000' } };
        bondRow.getCell(3).fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFB4C6E7' }
        };
        bondRow.getCell(3).alignment = { horizontal: 'center' };
        bondRow.getCell(3).border = {
            top: { style: 'thin', color: { argb: 'FF000000' } },
            left: { style: 'thin', color: { argb: 'FF000000' } },
            bottom: { style: 'thin', color: { argb: 'FF000000' } },
            right: { style: 'thin', color: { argb: 'FF000000' } }
        };

        bondRow.getCell(4).font = { color: { argb: 'FF000000' } };
        bondRow.getCell(4).fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFB4C6E7' }
        };
        bondRow.getCell(4).alignment = { horizontal: 'center' };
        bondRow.getCell(4).border = {
            top: { style: 'thin', color: { argb: 'FF000000' } },
            left: { style: 'thin', color: { argb: 'FF000000' } },
            bottom: { style: 'thin', color: { argb: 'FF000000' } },
            right: { style: 'thin', color: { argb: 'FF000000' } }
        };

        // Add TOTAL ORDER row (B19:E19)
        const totalRow = worksheet.getRow(19);
        totalRow.getCell(2).value = 'TOTAL ORDER, USD';
        totalRow.getCell(4).value = '$';
        totalRow.getCell(5).value = '-';

        // Style TOTAL ORDER row - no background, black text
        totalRow.getCell(2).font = { bold: true, color: { argb: 'FF000000' } };
        totalRow.getCell(4).font = { color: { argb: 'FF000000' } };
        totalRow.getCell(5).font = { color: { argb: 'FF000000' } };
        totalRow.getCell(4).alignment = { horizontal: 'center' };
        totalRow.getCell(5).alignment = { horizontal: 'center' };
    }

    private createProvisionsSheet(workbook: ExcelJS.Workbook): void {
        // Filter data for provisions (meat, alcohol, cigarettes, etc.)
        const provisionsData = this.processedData.filter(item =>
            this.isProvisionItem(item.description)
        );

        const worksheet = workbook.addWorksheet('PROVISIONS');

        // Add headers
        const headers = ['Pos.', 'Description', 'Remark', 'Unit', 'Qty', 'Price', 'Total'];
        const headerRow = worksheet.addRow(headers);

        // Style header row with dark blue background and white text
        headerRow.eachCell((cell, colNumber) => {
            cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FF4472C4' }
            };
            cell.alignment = { horizontal: 'center', vertical: 'middle' };
            cell.border = {
                top: { style: 'thin', color: { argb: 'FF000000' } },
                left: { style: 'thin', color: { argb: 'FF000000' } },
                bottom: { style: 'thin', color: { argb: 'FF000000' } },
                right: { style: 'thin', color: { argb: 'FF000000' } }
            };
        });

        // Add data rows
        provisionsData.forEach((item, index) => {
            const dataRow = worksheet.addRow([
                (index + 1).toString(),
                item.description,
                item.remarks || '-',
                item.unit,
                '', // Empty Qty column
                `$ ${item.price.toFixed(2)}`,
                '' // Empty Total column
            ]);

            // Style data rows with borders
            dataRow.eachCell((cell, colNumber) => {
                cell.border = {
                    top: { style: 'thin', color: { argb: 'FF000000' } },
                    left: { style: 'thin', color: { argb: 'FF000000' } },
                    bottom: { style: 'thin', color: { argb: 'FF000000' } },
                    right: { style: 'thin', color: { argb: 'FF000000' } }
                };
                cell.alignment = { vertical: 'middle' };
            });
        });

        // Set column widths
        worksheet.getColumn(1).width = 8;   // Pos.
        worksheet.getColumn(2).width = 40;  // Description
        worksheet.getColumn(3).width = 20;  // Remark
        worksheet.getColumn(4).width = 8;   // Unit
        worksheet.getColumn(5).width = 8;   // Qty
        worksheet.getColumn(6).width = 12;  // Price
        worksheet.getColumn(7).width = 12; // Total
    }

    private createFreshProvisionsSheet(workbook: ExcelJS.Workbook): void {
        // Filter data for fresh provisions (fruits, vegetables)
        const freshProvisionsData = this.processedData.filter(item =>
            this.isFreshProvisionItem(item.description)
        );

        const worksheet = workbook.addWorksheet('FRESH PROVISIONS');

        // Add headers
        const headers = ['Pos.', 'Description', 'Remark', 'Unit', 'Qty', 'Price', 'Total'];
        const headerRow = worksheet.addRow(headers);

        // Style header row with dark blue background and white text
        headerRow.eachCell((cell, colNumber) => {
            cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FF4472C4' }
            };
            cell.alignment = { horizontal: 'center', vertical: 'middle' };
            cell.border = {
                top: { style: 'thin', color: { argb: 'FF000000' } },
                left: { style: 'thin', color: { argb: 'FF000000' } },
                bottom: { style: 'thin', color: { argb: 'FF000000' } },
                right: { style: 'thin', color: { argb: 'FF000000' } }
            };
        });

        // Add data rows
        freshProvisionsData.forEach((item, index) => {
            const dataRow = worksheet.addRow([
                (index + 1).toString(),
                item.description,
                item.remarks || '-',
                item.unit,
                '', // Empty Qty column
                `$ ${item.price.toFixed(2)}`,
                '' // Empty Total column
            ]);

            // Style data rows with borders
            dataRow.eachCell((cell, colNumber) => {
                cell.border = {
                    top: { style: 'thin', color: { argb: 'FF000000' } },
                    left: { style: 'thin', color: { argb: 'FF000000' } },
                    bottom: { style: 'thin', color: { argb: 'FF000000' } },
                    right: { style: 'thin', color: { argb: 'FF000000' } }
                };
                cell.alignment = { vertical: 'middle' };
            });
        });

        // Set column widths
        worksheet.getColumn(1).width = 8;   // Pos.
        worksheet.getColumn(2).width = 40;  // Description
        worksheet.getColumn(3).width = 20;  // Remark
        worksheet.getColumn(4).width = 8;   // Unit
        worksheet.getColumn(5).width = 8;   // Qty
        worksheet.getColumn(6).width = 12;  // Price
        worksheet.getColumn(7).width = 12; // Total
    }

    private createBondSheet(workbook: ExcelJS.Workbook): void {
        // Filter data for bond items (alcohol, spirits)
        const bondData = this.processedData.filter(item =>
            this.isBondItem(item.description)
        );

        const worksheet = workbook.addWorksheet('BOND');

        // Add headers
        const headers = ['Pos.', 'Description', 'Remark', 'Unit', 'Qty', 'Price', 'Total'];
        const headerRow = worksheet.addRow(headers);

        // Style header row with dark blue background and white text
        headerRow.eachCell((cell, colNumber) => {
            cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FF4472C4' }
            };
            cell.alignment = { horizontal: 'center', vertical: 'middle' };
            cell.border = {
                top: { style: 'thin', color: { argb: 'FF000000' } },
                left: { style: 'thin', color: { argb: 'FF000000' } },
                bottom: { style: 'thin', color: { argb: 'FF000000' } },
                right: { style: 'thin', color: { argb: 'FF000000' } }
            };
        });

        // Add data rows
        bondData.forEach((item, index) => {
            const dataRow = worksheet.addRow([
                (index + 1).toString(),
                item.description,
                item.remarks || '-',
                item.unit,
                '', // Empty Qty column
                `$ ${item.price.toFixed(2)}`,
                '' // Empty Total column
            ]);

            // Style data rows with borders
            dataRow.eachCell((cell, colNumber) => {
                cell.border = {
                    top: { style: 'thin', color: { argb: 'FF000000' } },
                    left: { style: 'thin', color: { argb: 'FF000000' } },
                    bottom: { style: 'thin', color: { argb: 'FF000000' } },
                    right: { style: 'thin', color: { argb: 'FF000000' } }
                };
                cell.alignment = { vertical: 'middle' };
            });
        });

        // Set column widths
        worksheet.getColumn(1).width = 8;   // Pos.
        worksheet.getColumn(2).width = 40;  // Description
        worksheet.getColumn(3).width = 20;  // Remark
        worksheet.getColumn(4).width = 8;   // Unit
        worksheet.getColumn(5).width = 8;   // Qty
        worksheet.getColumn(6).width = 12;  // Price
        worksheet.getColumn(7).width = 12; // Total
    }

    private isProvisionItem(description: string): boolean {
        const provisionKeywords = [
            'beef', 'lamb', 'pork', 'chicken', 'meat', 'fish', 'seafood',
            'whisky', 'vodka', 'rum', 'cognac', 'wine', 'beer', 'alcohol',
            'marlboro', 'philip', 'chesterfield', 'lucky', 'camel', 'cigarette',
            'coca', 'sprite', 'pepsi', 'seven', 'fanta', 'juice'
        ];

        const desc = description.toLowerCase();
        return provisionKeywords.some(keyword => desc.includes(keyword));
    }

    private isFreshProvisionItem(description: string): boolean {
        const freshKeywords = [
            'apple', 'banana', 'orange', 'grape', 'lemon', 'lime', 'avocado',
            'carrot', 'lettuce', 'tomato', 'onion', 'potato', 'cucumber',
            'broccoli', 'cauliflower', 'cabbage', 'spinach', 'kale'
        ];

        const desc = description.toLowerCase();
        return freshKeywords.some(keyword => desc.includes(keyword));
    }

    private isBondItem(description: string): boolean {
        const bondKeywords = [
            'whisky', 'vodka', 'rum', 'cognac', 'wine', 'beer', 'alcohol',
            'spirit', 'liquor'
        ];

        const desc = description.toLowerCase();
        return bondKeywords.some(keyword => desc.includes(keyword));
    }

}
