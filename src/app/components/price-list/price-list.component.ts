import { Component, OnInit, OnDestroy } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { DataService, ProcessedDataRow } from '../../services/data.service';
import { LoggingService } from '../../services/logging.service';
import { FRESH_PROVISIONS_LIST, NOT_FRESH } from '../../constants/fresh-provisions.constants';
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
    // Price List functionality
    processedData: ProcessedDataRow[] = [];
    filteredData: ProcessedDataRow[] = [];
    hasProcessedFiles = false; // Track if files have been processed
    sortState: SortState = { column: '', direction: 'asc' };

    // Filter properties
    selectedFileName = '';
    selectedDescription = '';
    descriptionTextFilter = '';
    availableFileNames: string[] = [];
    commonDescriptions = ['Beer', 'Cheese', 'Ice', 'Provision'];

    // Row expansion properties
    expandedRowIndex: number | null = null;

    // Select all functionality
    selectAllState: 'all' | 'some' | 'none' = 'all';

    // Invalid price tracking
    invalidPriceRecords: ProcessedDataRow[] = [];
    showInvalidPriceDialog = false;

    // Company selection for export
    selectedCompany: string = 'EOS';

    // Currency selection
    selectedCurrency: string = '';

    // Export filename
    exportFileName: string = '';

    // Country selection
    selectedCountry: string = '';

    // Country names
    countries = [
        'Afghanistan', 'Albania', 'Algeria', 'Argentina', 'Armenia', 'Australia',
        'Austria', 'Azerbaijan', 'Bahrain', 'Bangladesh', 'Belarus', 'Belgium',
        'Bolivia', 'Brazil', 'Bulgaria', 'Cambodia', 'Canada', 'Chile', 'China',
        'Colombia', 'Croatia', 'Cyprus', 'Czech Republic', 'Denmark', 'Ecuador',
        'Egypt', 'Estonia', 'Finland', 'France', 'Georgia', 'Germany', 'Ghana',
        'Greece', 'Hungary', 'Iceland', 'India', 'Indonesia', 'Iran', 'Iraq',
        'Ireland', 'Israel', 'Italy', 'Japan', 'Jordan', 'Kazakhstan', 'Kenya',
        'Kuwait', 'Latvia', 'Lebanon', 'Lithuania', 'Luxembourg', 'Malaysia',
        'Malta', 'Mexico', 'Morocco', 'Netherlands', 'New Zealand', 'Nigeria',
        'Norway', 'Oman', 'Pakistan', 'Peru', 'Philippines', 'Poland', 'Portugal',
        'Qatar', 'Romania', 'Russia', 'Saudi Arabia', 'Singapore', 'Slovakia',
        'Slovenia', 'South Africa', 'South Korea', 'Spain', 'Sri Lanka', 'Sweden',
        'Switzerland', 'Thailand', 'Turkey', 'UAE', 'Ukraine', 'United Kingdom',
        'United States', 'Uruguay', 'Vietnam'
    ];

    constructor(private dataService: DataService, private loggingService: LoggingService) { }

    ngOnInit(): void {
        // Initialize export filename
        this.updateExportFileName();

        this.dataService.processedData$.subscribe(data => {
            this.processedData = data;
            this.filteredData = [...data]; // Initialize filtered data
            this.updateAvailableFileNames();
            this.updateCommonDescriptions();
            this.applyFilters();
            this.updateSelectAllState();
            this.updateExportFileName(); // Update filename when data changes

            // Update hasProcessedFiles based on data availability
            this.hasProcessedFiles = data.length > 0;

            this.loggingService.logDataProcessing('data_loaded', {
                totalRecords: data.length,
                filteredRecords: this.filteredData.length
            }, 'PriceListComponent');
        });

        // Subscribe to price divider changes to update the display
        this.dataService.priceDivider$.subscribe(() => {
            // The processed data will be automatically updated by the data service
            // when price divider changes, so we just need to refresh our filtered data
            this.applyFilters();
        });

        // Add document click listener to close expanded rows when clicking outside
        document.addEventListener('click', this.onDocumentClick.bind(this));
    }


    onIncludeChange(index: number, event: Event): void {
        const checkbox = event.target as HTMLInputElement;
        const included = checkbox.checked;

        // Find the original index in processedData
        const filteredRow = this.filteredData[index];
        const originalIndex = this.processedData.findIndex(row =>
            row.fileName === filteredRow.fileName &&
            row.description === filteredRow.description &&
            row.price === filteredRow.price &&
            row.unit === filteredRow.unit
        );

        if (originalIndex !== -1) {
            this.dataService.updateRowIncluded(originalIndex, included);
            this.updateSelectAllState();

            this.loggingService.logDataSelection('item_inclusion_toggled',
                this.filteredData.filter(row => row.included).length,
                this.filteredData.length,
                'PriceListComponent'
            );
        }
    }

    onSelectAllChange(event: Event): void {
        const checkbox = event.target as HTMLInputElement;
        const checked = checkbox.checked;

        this.loggingService.logDataSelection('select_all_toggled',
            checked ? this.filteredData.length : 0,
            this.filteredData.length,
            'PriceListComponent'
        );

        if (checked) {
            // When checking "Select All", only select the currently filtered/visible rows
            this.filteredData.forEach((filteredRow, index) => {
                const originalIndex = this.processedData.findIndex(row =>
                    row.fileName === filteredRow.fileName &&
                    row.description === filteredRow.description &&
                    row.price === filteredRow.price &&
                    row.unit === filteredRow.unit
                );

                if (originalIndex !== -1) {
                    this.dataService.updateRowIncluded(originalIndex, checked);
                }
            });
        } else {
            // When unchecking "Select All", uncheck ALL rows in the entire dataset
            this.processedData.forEach((row, index) => {
                this.dataService.updateRowIncluded(index, false);
            });
        }

        this.updateSelectAllState();
    }

    updateSelectAllState(): void {
        const totalRows = this.filteredData.length;
        const checkedRows = this.filteredData.filter(row => row.included).length;

        if (checkedRows === 0) {
            this.selectAllState = 'none';
        } else if (checkedRows === totalRows) {
            this.selectAllState = 'all';
        } else {
            this.selectAllState = 'some';
        }
    }

    getSelectedCount(): number {
        return this.filteredData.filter(row => row.included).length;
    }

    isPriceValid(price: any): boolean {
        return typeof price === 'number' && !isNaN(price) && isFinite(price);
    }

    validatePrices(): void {
        this.invalidPriceRecords = this.processedData.filter(row => !this.isPriceValid(row.price));
    }

    getValidPriceRecords(): ProcessedDataRow[] {
        return this.processedData.filter(row => this.isPriceValid(row.price));
    }

    showInvalidPricesDialog(): void {
        this.validatePrices();
        this.showInvalidPriceDialog = true;
    }

    closeInvalidPriceDialog(): void {
        this.showInvalidPriceDialog = false;
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
                case 'included':
                    aValue = a.included ? 1 : 0;
                    bValue = b.included ? 1 : 0;
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
        this.updateSelectAllState();
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
        this.loggingService.logButtonClick('export_to_excel', 'PriceListComponent', {
            totalRecords: this.processedData.length,
            selectedRecords: this.filteredData.filter(row => row.included).length
        });

        try {
            const workbook = new ExcelJS.Workbook();

            // Check which data types we have and get the separateFreshProvisions setting
            const separateFreshProvisions = this.dataService.getSeparateFreshProvisions();
            const hasProvisionsData = this.hasProvisionsData();
            const hasBondData = this.hasBondData();

            // Create Cover Sheet (pass data availability flags and separateFreshProvisions setting)
            this.createCoverSheet(workbook, hasProvisionsData, hasBondData, separateFreshProvisions);

            // Create sheets only if they have data
            if (hasProvisionsData) {
                this.createProvisionsSheet(workbook, separateFreshProvisions);
                // Only create Fresh Provisions sheet if separateFreshProvisions is enabled
                if (separateFreshProvisions) {
                    this.createFreshProvisionsSheet(workbook);
                }
            }

            if (hasBondData) {
                this.createBondSheet(workbook);
            }

            this.applyCambriaFont(workbook);

            // Generate Excel file
            const buffer = await workbook.xlsx.writeBuffer();
            const data = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });

            // Generate filename with current date in yyyyMMdd format
            const today = new Date();
            const year = today.getFullYear();
            const month = String(today.getMonth() + 1).padStart(2, '0');
            const day = String(today.getDate()).padStart(2, '0');
            const dateString = `${year}${month}${day}`;

            // Use the export filename from input, or fallback to default if empty
            // If exportFileName is empty, use the computed default (what's shown in the input)
            let fileName = this.exportFileName.trim() || this.generateDefaultFileName();
            
            // Add .xlsx extension if not already present
            if (!fileName.toLowerCase().endsWith('.xlsx')) {
                fileName += '.xlsx';
            }

            saveAs(data, fileName);

            this.loggingService.logExport('excel_export_successful', {
                fileName,
                fileSize: data.size,
                totalRecords: this.processedData.length,
                selectedRecords: this.filteredData.filter(row => row.included).length
            }, 'PriceListComponent');
        } catch (error) {
            this.loggingService.logError(error as Error, 'excel_export', 'PriceListComponent');
        }
    }

    private applyCambriaFont(workbook: ExcelJS.Workbook): void {
        workbook.eachSheet(worksheet => {
            worksheet.eachRow({ includeEmpty: true }, row => {
                row.eachCell({ includeEmpty: true }, cell => {
                    const cellFont = cell.font || {};
                    cell.font = { ...cellFont, name: 'Cambria', size: 11 };

                    const value = cell.value;
                    if (value && typeof value === 'object' && 'richText' in value && Array.isArray((value as ExcelJS.CellRichTextValue).richText)) {
                        const richTextValue = value as ExcelJS.CellRichTextValue;
                        richTextValue.richText = richTextValue.richText.map(part => ({
                            ...part,
                            font: { ...(part.font || {}), name: 'Cambria', size: 11 }
                        }));
                        cell.value = richTextValue;
                    }
                });
            });
        });
    }

    private toUpperCellText(value: string | null | undefined): string {
        if (value === null || value === undefined) {
            return '';
        }
        return value.toString().toUpperCase();
    }

    private createCoverSheet(workbook: ExcelJS.Workbook, hasProvisionsData: boolean, hasBondData: boolean, separateFreshProvisions: boolean): void {
        const worksheet = workbook.addWorksheet('COVER SHEET');

        // Set tab color - white with green underline (using light green for tab color)
        worksheet.properties.tabColor = { argb: 'FF90EE90' }; // Light green

        // Remove grid lines from the worksheet
        worksheet.properties.showGridLines = false;

        // Set view options to hide gridlines
        worksheet.views = [{
            showGridLines: false
        }];

        // Set column widths to show all content properly
        worksheet.getColumn('A').width = 20;  // Column A
        worksheet.getColumn('B').width = 30;  // Column B (for email and labels)
        worksheet.getColumn('C').width = 15;  // Column C (for combined $ and - values)

        // Set default font formatting for column B - Cambria, Regular, size 16
        worksheet.getColumn('B').font = {
            name: 'Cambria',
            size: 16
        };

        // Add email (moved to B9 to match image) - based on selected company
        const email = this.selectedCompany === 'Hi Marine'
            ? 'office@himarinecompany.com'
            : 'office@eos-supply.co.uk';
        worksheet.getCell('B9').value = email;

        // Add PROVISIONS row (B14:C14) with dark blue background and white text - only if we have provisions data
        let currentRow = 14;
        if (hasProvisionsData) {
            const provisionsRow = worksheet.getRow(currentRow);
            provisionsRow.getCell(2).value = 'PROVISIONS';
            // Calculate the total row number dynamically based on data length
            // If separateFreshProvisions is false, include all provisions (including fresh) in PROVISIONS
            // If separateFreshProvisions is true, only include non-fresh provisions
            let provisionsDataLength: number;
            if (separateFreshProvisions) {
                provisionsDataLength = this.getValidPriceRecords().filter(item => this.isProvisionItem(item) && !this.isFreshProvisionItem(item) && item.included).length;
            } else {
                provisionsDataLength = this.getValidPriceRecords().filter(item => this.isProvisionItem(item) && item.included).length;
            }
            const provisionsTotalRow = provisionsDataLength + 3; // +1 for header, +2 for two rows below last data
            if (provisionsDataLength > 0) {
                provisionsRow.getCell(3).value = { formula: `PROVISIONS!G${provisionsTotalRow}` };
            } else {
                provisionsRow.getCell(3).value = 0; // No data, so total is 0
            }

            // Style PROVISIONS row - dark blue background with white text
            provisionsRow.getCell(2).font = {
                name: 'Cambria',
                bold: true,
                size: 16,
                color: { argb: 'FFFFFFFF' }
            };
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

            provisionsRow.getCell(3).font = {
                name: 'Cambria',
                bold: true,
                size: 16,
                color: { argb: 'FF000000' }
            };
            provisionsRow.getCell(3).fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFFFFFFF' }
            };
            provisionsRow.getCell(3).alignment = { horizontal: 'center' };
            provisionsRow.getCell(3).numFmt = this.getCurrencyFormat();
            provisionsRow.getCell(3).border = {
                top: { style: 'thin', color: { argb: 'FF000000' } },
                left: { style: 'thin', color: { argb: 'FF000000' } },
                bottom: { style: 'thin', color: { argb: 'FF000000' } },
                right: { style: 'thin', color: { argb: 'FF000000' } }
            };

            currentRow++;
        }

        // Add FRESH PROVISIONS row - only if we have provisions data AND separateFreshProvisions is enabled
        if (hasProvisionsData && separateFreshProvisions) {
            const freshProvisionsRow = worksheet.getRow(currentRow);
            freshProvisionsRow.getCell(2).value = 'FRESH PROVISIONS';
            // Calculate the total row number dynamically based on data length (only included items with valid prices)
            const freshProvisionsDataLength = this.getValidPriceRecords().filter(item => this.isFreshProvisionItem(item) && item.included).length;
            const freshProvisionsTotalRow = freshProvisionsDataLength + 3; // +1 for header, +2 for two rows below last data
            if (freshProvisionsDataLength > 0) {
                freshProvisionsRow.getCell(3).value = { formula: `'FRESH PROVISIONS'!G${freshProvisionsTotalRow}` };
            } else {
                freshProvisionsRow.getCell(3).value = 0; // No data, so total is 0
            }

            // Style FRESH PROVISIONS row - light blue background with black text
            freshProvisionsRow.getCell(2).font = {
                name: 'Cambria',
                bold: true,
                size: 16,
                color: { argb: 'FF000000' }
            };
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

            freshProvisionsRow.getCell(3).font = {
                name: 'Cambria',
                bold: true,
                size: 16,
                color: { argb: 'FF000000' }
            };
            freshProvisionsRow.getCell(3).fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFFFFFFF' }
            };
            freshProvisionsRow.getCell(3).alignment = { horizontal: 'center' };
            freshProvisionsRow.getCell(3).numFmt = this.getCurrencyFormat();
            freshProvisionsRow.getCell(3).border = {
                top: { style: 'thin', color: { argb: 'FF000000' } },
                left: { style: 'thin', color: { argb: 'FF000000' } },
                bottom: { style: 'thin', color: { argb: 'FF000000' } },
                right: { style: 'thin', color: { argb: 'FF000000' } }
            };

            currentRow++;
        }

        // Add BOND row - only if we have bond data
        if (hasBondData) {
            const bondRow = worksheet.getRow(currentRow);
            bondRow.getCell(2).value = 'BOND';
            // Calculate the total row number dynamically based on data length (only included items with valid prices)
            const bondDataLength = this.getValidPriceRecords().filter(item => this.isBondItem(item) && item.included).length;
            const bondTotalRow = bondDataLength + 3; // +1 for header, +2 for two rows below last data
            if (bondDataLength > 0) {
                bondRow.getCell(3).value = { formula: `BOND!G${bondTotalRow}` };
            } else {
                bondRow.getCell(3).value = 0; // No data, so total is 0
            }

            // Style BOND row - light blue background with black text
            bondRow.getCell(2).font = {
                name: 'Cambria',
                bold: true,
                size: 16,
                color: { argb: 'FF000000' }
            };
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

            bondRow.getCell(3).font = {
                name: 'Cambria',
                bold: true,
                size: 16,
                color: { argb: 'FF000000' }
            };
            bondRow.getCell(3).fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFFFFFFF' }
            };
            bondRow.getCell(3).alignment = { horizontal: 'center' };
            bondRow.getCell(3).numFmt = this.getCurrencyFormat();
            bondRow.getCell(3).border = {
                top: { style: 'thin', color: { argb: 'FF000000' } },
                left: { style: 'thin', color: { argb: 'FF000000' } },
                bottom: { style: 'thin', color: { argb: 'FF000000' } },
                right: { style: 'thin', color: { argb: 'FF000000' } }
            };
        }

        // Add TOTAL ORDER row - dynamically positioned
        const totalRow = worksheet.getRow(19);
        totalRow.getCell(2).value = `TOTAL ORDER, ${this.getCurrencyCode()}`;
        // Build dynamic SUM formula based on which rows are present
        let sumFormula = '=';
        if (hasProvisionsData && hasBondData && separateFreshProvisions) {
            sumFormula = '=SUM(C14:C16)'; // All three rows (provisions, fresh provisions, bond)
        } else if (hasProvisionsData && hasBondData && !separateFreshProvisions) {
            sumFormula = '=SUM(C14:C15)'; // Provisions (includes fresh) and bond
        } else if (hasProvisionsData && !hasBondData && separateFreshProvisions) {
            sumFormula = '=SUM(C14:C15)'; // Only provisions and fresh provisions
        } else if (hasProvisionsData && !hasBondData && !separateFreshProvisions) {
            sumFormula = '=C14'; // Only provisions (includes fresh)
        } else if (!hasProvisionsData && hasBondData) {
            sumFormula = '=C14'; // Only bond
        } else {
            sumFormula = '=0'; // No data
        }
        totalRow.getCell(3).value = { formula: sumFormula };

        // Style TOTAL ORDER row - no background, black text
        totalRow.getCell(2).font = {
            name: 'Cambria',
            bold: true,
            size: 16,
            color: { argb: 'FF000000' }
        };
        totalRow.getCell(3).font = {
            name: 'Cambria',
            bold: true,
            size: 16,
            color: { argb: 'FF000000' }
        };
        totalRow.getCell(3).alignment = { horizontal: 'center' };
        totalRow.getCell(3).numFmt = this.getCurrencyFormat();
    }

    private createProvisionsSheet(workbook: ExcelJS.Workbook, separateFreshProvisions: boolean): void {
        // Filter data for provisions
        // If separateFreshProvisions is false, include all provisions (including fresh)
        // If separateFreshProvisions is true, only include non-fresh provisions
        let provisionsData: ProcessedDataRow[];
        if (separateFreshProvisions) {
            // Data uploaded via "Provisions" dropzone AND NOT fresh provisions AND valid price
            provisionsData = this.getValidPriceRecords().filter(item =>
                this.isProvisionItem(item) && !this.isFreshProvisionItem(item) && item.included
            );
        } else {
            // Data uploaded via "Provisions" dropzone AND valid price (includes fresh provisions)
            provisionsData = this.getValidPriceRecords().filter(item =>
                this.isProvisionItem(item) && item.included
            );
        }

        const worksheet = workbook.addWorksheet('PROVISIONS');

        // Set tab color - medium blue
        worksheet.properties.tabColor = { argb: 'FF4472C4' }; // Medium blue

        // Add headers
        const headers = ['Pos', 'Description', 'Remark', 'Unit', 'Qty', 'Price', 'Total'];
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
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            };
        });

        // Add data rows
        provisionsData.forEach((item, index) => {
            const rowNumber = index + 2; // +2 because header is row 1, data starts at row 2
            const dataRow = worksheet.addRow([
                (index + 1).toString(),
                this.toUpperCellText(item.description),
                this.toUpperCellText(item.remarks || '-'),
                this.toUpperCellText(item.unit),
                0, // Qty column - set to 0 instead of empty string
                this.roundToTwoDecimals(item.price), // Price as number for formula calculation, rounded to 2 decimals
                '' // Empty Total column - will be filled with formula
            ]);

            // Add formula to column G (Total column) - simplified formula
            const totalCell = worksheet.getCell(`G${rowNumber}`);
            totalCell.value = { formula: `=F${rowNumber}*E${rowNumber}` };

            // Format Price column (F) and Total column (G) with Accounting format
            const priceCell = worksheet.getCell(`F${rowNumber}`);
            priceCell.numFmt = this.getCurrencyFormat();

            totalCell.numFmt = this.getCurrencyFormat();

            // Style data rows with simple borders
            dataRow.eachCell((cell, colNumber) => {
                cell.border = {
                    top: { style: 'thin' },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' }
                };
                cell.alignment = { vertical: 'middle' };
            });
            
            // Enable word-wrap for Description column (column 2)
            const descCell = worksheet.getCell(`B${rowNumber}`);
            descCell.alignment = { ...descCell.alignment, wrapText: true };
        });

        // Set column widths
        worksheet.getColumn(1).width = 8;   // Pos.
        worksheet.getColumn(2).width = 40;  // Description
        worksheet.getColumn(3).width = 20;  // Remark
        worksheet.getColumn(4).width = 8;   // Unit
        worksheet.getColumn(5).width = 8;   // Qty
        worksheet.getColumn(6).width = 12;  // Price
        worksheet.getColumn(7).width = 12; // Total

        // Add TOTAL USD row two rows below the last data row
        const lastDataRow = provisionsData.length + 1; // +1 for header row
        const totalRowNumber = lastDataRow + 2; // Two rows below last data

        // Add TOTAL currency label in column F
        const totalLabelCell = worksheet.getCell(`F${totalRowNumber}`);
        totalLabelCell.value = `TOTAL ${this.getCurrencyCode()}`;
        totalLabelCell.font = { bold: true };
        totalLabelCell.alignment = { horizontal: 'left' };

        // Add sum formula in column G - handle empty data case
        const totalValueCell = worksheet.getCell(`G${totalRowNumber}`);
        if (provisionsData.length > 0) {
            totalValueCell.value = { formula: `=SUM(G2:G${lastDataRow})` };
        } else {
            totalValueCell.value = 0; // No data, so total is 0
        }
        totalValueCell.numFmt = this.getCurrencyFormat();
        totalValueCell.font = { bold: true };
        totalValueCell.alignment = { horizontal: 'right' };
    }

    private createFreshProvisionsSheet(workbook: ExcelJS.Workbook): void {
        // Filter data for fresh provisions: Data uploaded via "Provisions" dropzone AND has description containing fresh provisions keywords AND valid price
        const freshProvisionsData = this.getValidPriceRecords().filter(item =>
            this.isFreshProvisionItem(item) && item.included
        );

        const worksheet = workbook.addWorksheet('FRESH PROVISIONS');

        // Set tab color - medium blue (same as PROVISIONS)
        worksheet.properties.tabColor = { argb: 'FF4472C4' }; // Medium blue

        // Add headers
        const headers = ['Pos', 'Description', 'Remark', 'Unit', 'Qty', 'Price', 'Total'];
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
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            };
        });

        // Add data rows
        freshProvisionsData.forEach((item, index) => {
            const rowNumber = index + 2; // +2 because header is row 1, data starts at row 2
            const dataRow = worksheet.addRow([
                (index + 1).toString(),
                this.toUpperCellText(item.description),
                this.toUpperCellText(item.remarks || '-'),
                this.toUpperCellText(item.unit),
                0, // Qty column - set to 0 instead of empty string
                this.roundToTwoDecimals(item.price), // Price as number for formula calculation, rounded to 2 decimals
                '' // Empty Total column - will be filled with formula
            ]);

            // Add formula to column G (Total column) - simplified formula
            const totalCell = worksheet.getCell(`G${rowNumber}`);
            totalCell.value = { formula: `=F${rowNumber}*E${rowNumber}` };

            // Format Price column (F) and Total column (G) with Accounting format
            const priceCell = worksheet.getCell(`F${rowNumber}`);
            priceCell.numFmt = this.getCurrencyFormat();

            totalCell.numFmt = this.getCurrencyFormat();

            // Style data rows with simple borders
            dataRow.eachCell((cell, colNumber) => {
                cell.border = {
                    top: { style: 'thin' },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' }
                };
                cell.alignment = { vertical: 'middle' };
            });
            
            // Enable word-wrap for Description column (column 2)
            const descCell = worksheet.getCell(`B${rowNumber}`);
            descCell.alignment = { ...descCell.alignment, wrapText: true };
        });

        // Set column widths
        worksheet.getColumn(1).width = 8;   // Pos.
        worksheet.getColumn(2).width = 40;  // Description
        worksheet.getColumn(3).width = 20;  // Remark
        worksheet.getColumn(4).width = 8;   // Unit
        worksheet.getColumn(5).width = 8;   // Qty
        worksheet.getColumn(6).width = 12;  // Price
        worksheet.getColumn(7).width = 12; // Total

        // Add TOTAL USD row two rows below the last data row
        const lastDataRow = freshProvisionsData.length + 1; // +1 for header row
        const totalRowNumber = lastDataRow + 2; // Two rows below last data

        // Add TOTAL currency label in column F
        const totalLabelCell = worksheet.getCell(`F${totalRowNumber}`);
        totalLabelCell.value = `TOTAL ${this.getCurrencyCode()}`;
        totalLabelCell.font = { bold: true };
        totalLabelCell.alignment = { horizontal: 'left' };

        // Add sum formula in column G - handle empty data case
        const totalValueCell = worksheet.getCell(`G${totalRowNumber}`);
        if (freshProvisionsData.length > 0) {
            totalValueCell.value = { formula: `=SUM(G2:G${lastDataRow})` };
        } else {
            totalValueCell.value = 0; // No data, so total is 0
        }
        totalValueCell.numFmt = this.getCurrencyFormat();
        totalValueCell.font = { bold: true };
        totalValueCell.alignment = { horizontal: 'right' };
    }

    private createBondSheet(workbook: ExcelJS.Workbook): void {
        // Filter data for bond items: Data uploaded via "Bonded" dropzone AND valid price
        const bondData = this.getValidPriceRecords().filter(item =>
            this.isBondItem(item) && item.included
        );

        const worksheet = workbook.addWorksheet('BOND');

        // Set tab color - light blue
        worksheet.properties.tabColor = { argb: 'FFB4C6E7' }; // Light blue

        // Add headers
        const headers = ['Pos', 'Description', 'Remark', 'Unit', 'Qty', 'Price', 'Total'];
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
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            };
        });

        // Add data rows
        bondData.forEach((item, index) => {
            const rowNumber = index + 2; // +2 because header is row 1, data starts at row 2
            const dataRow = worksheet.addRow([
                (index + 1).toString(),
                this.toUpperCellText(item.description),
                this.toUpperCellText(item.remarks || '-'),
                this.toUpperCellText(item.unit),
                0, // Qty column - set to 0 instead of empty string
                this.roundToTwoDecimals(item.price), // Price as number for formula calculation, rounded to 2 decimals
                '' // Empty Total column - will be filled with formula
            ]);

            // Add formula to column G (Total column) - simplified formula
            const totalCell = worksheet.getCell(`G${rowNumber}`);
            totalCell.value = { formula: `=F${rowNumber}*E${rowNumber}` };

            // Format Price column (F) and Total column (G) with Accounting format
            const priceCell = worksheet.getCell(`F${rowNumber}`);
            priceCell.numFmt = this.getCurrencyFormat();

            totalCell.numFmt = this.getCurrencyFormat();

            // Style data rows with simple borders
            dataRow.eachCell((cell, colNumber) => {
                cell.border = {
                    top: { style: 'thin' },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' }
                };
                cell.alignment = { vertical: 'middle' };
            });
            
            // Enable word-wrap for Description column (column 2)
            const descCell = worksheet.getCell(`B${rowNumber}`);
            descCell.alignment = { ...descCell.alignment, wrapText: true };
        });

        // Set column widths
        worksheet.getColumn(1).width = 8;   // Pos.
        worksheet.getColumn(2).width = 40;  // Description
        worksheet.getColumn(3).width = 20;  // Remark
        worksheet.getColumn(4).width = 8;   // Unit
        worksheet.getColumn(5).width = 8;   // Qty
        worksheet.getColumn(6).width = 12;  // Price
        worksheet.getColumn(7).width = 12; // Total

        // Add TOTAL USD row two rows below the last data row
        const lastDataRow = bondData.length + 1; // +1 for header row
        const totalRowNumber = lastDataRow + 2; // Two rows below last data

        // Add TOTAL currency label in column F
        const totalLabelCell = worksheet.getCell(`F${totalRowNumber}`);
        totalLabelCell.value = `TOTAL ${this.getCurrencyCode()}`;
        totalLabelCell.font = { bold: true };
        totalLabelCell.alignment = { horizontal: 'left' };

        // Add sum formula in column G - handle empty data case
        const totalValueCell = worksheet.getCell(`G${totalRowNumber}`);
        if (bondData.length > 0) {
            totalValueCell.value = { formula: `=SUM(G2:G${lastDataRow})` };
        } else {
            totalValueCell.value = 0; // No data, so total is 0
        }
        totalValueCell.numFmt = this.getCurrencyFormat();
        totalValueCell.font = { bold: true };
        totalValueCell.alignment = { horizontal: 'right' };
    }

    private isProvisionItem(item: ProcessedDataRow): boolean {
        // Data uploaded via "Provisions" dropzone should go in "PROVISIONS" Excel tab
        return item.category === 'Provisions';
    }

    private isFreshProvisionItem(item: ProcessedDataRow): boolean {
        // Data uploaded via "Provisions" dropzone AND has 
        // "Description" containing a word from FRESH_PROVISIONS_LIST 
        // but NOT from NOT_FRESH should go in "FRESH PROVISIONS" Excel tab
        if (item.category !== 'Provisions') {
            return false;
        }

        const desc = item.description.toUpperCase();
        const hasFreshProvisionKeyword = FRESH_PROVISIONS_LIST
            .some((keyword: string) => desc.includes(keyword));
        const hasNotFreshKeyword = NOT_FRESH
            .some((keyword: string) => desc.includes(keyword));

        return hasFreshProvisionKeyword && !hasNotFreshKeyword;
    }

    private isBondItem(item: ProcessedDataRow): boolean {
        // Data uploaded via "Bonded" dropzone should go in "BOND" Excel tab
        return item.category === 'Bonded';
    }

    private hasProvisionsData(): boolean {
        // Check if there are any provisions items with valid prices and included
        const validData = this.getValidPriceRecords();
        return validData.some(item => this.isProvisionItem(item) && item.included);
    }

    private hasBondData(): boolean {
        // Check if there are any bond items with valid prices and included
        const validData = this.getValidPriceRecords();
        return validData.some(item => this.isBondItem(item) && item.included);
    }

    private getCurrencyFormat(): string {
        // Return Excel number format string based on selected currency
        switch (this.selectedCurrency) {
            case 'EUR':
                return '"€"#,##0.00';
            case 'AUD':
                return '"A$"#,##0.00';
            case 'NZD':
                return '"NZ$"#,##0.00';
            case 'USD':
                return '$#,##0.00';
            case 'CAD':
                return '"C$"#,##0.00';
            default:
                return '$#,##0.00'; // Default to USD
        }
    }

    private getCurrencyCode(): string {
        // Return currency code for labels
        return this.selectedCurrency || 'USD';
    }

    private roundToTwoDecimals(value: number): number {
        // Round number to 2 decimal places
        return Math.round(value * 100) / 100;
    }

    updateExportFileName(): void {
        // Update filename when company or data changes
        this.exportFileName = this.generateDefaultFileName();
    }

    private generateDefaultFileName(): string {
        // Start with company prefix
        let companyPrefix = '';
        if (this.selectedCompany === 'Hi Marine') {
            companyPrefix = 'Hi Marine Company Price List ';
        } else {
            companyPrefix = 'EOS Supply Price List ';
        }

        let fileName = companyPrefix;

        // Determine category part based on data
        const hasProvisions = this.processedData.some(item =>
            item.category === 'Provisions' && item.included && this.isPriceValid(item.price)
        );
        const hasBonded = this.processedData.some(item =>
            item.category === 'Bonded' && item.included && this.isPriceValid(item.price)
        );

        // Build category suffix
        let categorySuffix = '';
        if (hasProvisions && hasBonded) {
            categorySuffix = 'Provisions&Bond';
        } else if (hasBonded) {
            categorySuffix = 'Bond';
        } else if (hasProvisions) {
            categorySuffix = 'Provisions';
        } else {
            // Default if no data
            categorySuffix = 'Provisions';
        }

        fileName += categorySuffix;

        // Add country at the end if selected
        if (this.selectedCountry && this.selectedCountry.trim() !== '') {
            fileName += ' ' + this.selectedCountry;
        }

        // Don't add file extension here - it will be added in the template and export

        return fileName;
    }

    onCompanyChange(): void {
        // Update filename when company selection changes
        this.updateExportFileName();
    }

    onCountryChange(): void {
        // Update filename when country selection changes
        this.updateExportFileName();
    }

    getDefaultFileName(): string {
        // Public method to access default filename in template
        return this.generateDefaultFileName();
    }

}
