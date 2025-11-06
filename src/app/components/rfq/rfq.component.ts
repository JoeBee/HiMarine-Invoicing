import { Component } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { LoggingService } from '../../services/logging.service';
import * as XLSX from 'xlsx';

interface TabInfo {
    tabName: string;
    rowCount: number;
    topLeftCell: string;
    product: string;
    qty: string;
    unit: string;
    remark: string;
    isHidden: boolean;
    columnHeaders: string[];
}

interface FileAnalysis {
    fileName: string;
    numberOfTabs: number;
    tabs: TabInfo[];
    file: File;
}

@Component({
    selector: 'app-rfq',
    standalone: true,
    imports: [CommonModule, FormsModule],
    templateUrl: './rfq.component.html',
    styleUrls: ['./rfq.component.scss']
})
export class RfqComponent {
    isDragOver = false;
    uploadedFiles: File[] = [];
    fileAnalyses: FileAnalysis[] = [];
    isProcessing = false;
    errorMessage = '';
    selectedCompany: 'HI US' | 'HI UK' | 'EOS' = 'HI US';

    constructor(private loggingService: LoggingService) { }

    onDragOver(event: DragEvent): void {
        event.preventDefault();
        event.stopPropagation();
        this.isDragOver = true;
    }

    onDragLeave(event: DragEvent): void {
        event.preventDefault();
        event.stopPropagation();
        this.isDragOver = false;
    }

    onDrop(event: DragEvent): void {
        event.preventDefault();
        event.stopPropagation();
        this.isDragOver = false;

        const files = event.dataTransfer?.files;
        if (files && files.length > 0) {
            this.handleFiles(Array.from(files));
        }
    }

    onFileSelected(event: Event): void {
        const input = event.target as HTMLInputElement;
        if (input.files && input.files.length > 0) {
            this.handleFiles(Array.from(input.files));
        }
    }

    private handleFiles(files: File[]): void {
        // Filter only Excel files
        const excelFiles = files.filter(file => {
            const validTypes = [
                'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', // .xlsx
                'application/vnd.ms-excel', // .xls
                'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.macroEnabled' // .xlsm
            ];
            return validTypes.includes(file.type) || file.name.match(/\.(xlsx|xls|xlsm)$/i);
        });

        if (excelFiles.length === 0) {
            this.errorMessage = 'Please upload valid Excel files (.xlsx, .xls, or .xlsm)';
            return;
        }

        this.errorMessage = '';

        // Log file uploads
        excelFiles.forEach(file => {
            this.loggingService.logFileUpload(file.name, file.size, file.type, 'rfq', 'RfqComponent');
        });

        // Add files to uploadedFiles list
        this.uploadedFiles = [...this.uploadedFiles, ...excelFiles];

        // Process all files
        this.processFiles(excelFiles);
    }

    private async processFiles(files: File[]): Promise<void> {
        this.isProcessing = true;

        try {
            for (const file of files) {
                const analysis = await this.analyzeExcelFile(file);
                this.fileAnalyses.push(analysis);
            }
        } catch (error) {
            this.errorMessage = 'Error processing Excel files. Please ensure they are valid Excel files.';
            this.loggingService.logError(
                error as Error,
                'excel_file_processing',
                'RfqComponent',
                {
                    fileCount: files.length
                }
            );
        } finally {
            this.isProcessing = false;
        }
    }

    private analyzeExcelFile(file: File): Promise<FileAnalysis> {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();

            reader.onload = (e: any) => {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, {
                        type: 'array',
                        cellFormula: false,
                        cellHTML: false,
                        cellStyles: false,
                        sheetStubs: false
                    });

                    // Validate workbook structure
                    if (!workbook || (!workbook.Sheets && !workbook.SheetNames)) {
                        throw new Error('Invalid workbook structure - workbook, Sheets, or SheetNames missing');
                    }

                    const tabInfos: TabInfo[] = [];

                    // Get hidden sheet information from workbook properties
                    const hiddenSheets = new Set<string>();

                    // XLSX library stores hidden sheet info in workbook.Workbook.Sheets array
                    // Each sheet entry has a 'state' property: 'visible', 'hidden', or 'veryHidden'
                    if (workbook.Workbook && workbook.Workbook.Sheets) {
                        const sheets = workbook.Workbook.Sheets;

                        // Handle both array and object formats
                        if (Array.isArray(sheets)) {
                            sheets.forEach((sheet: any, index: number) => {
                                const state = sheet?.state || (sheet as any)?.State;
                                const name = sheet?.name || (sheet as any)?.Name || workbook.SheetNames[index];
                                if ((state === 'hidden' || state === 'veryHidden') && name) {
                                    hiddenSheets.add(name);
                                }
                            });
                        } else {
                            // Handle object format - iterate through sheet indices
                            Object.keys(sheets).forEach(key => {
                                const sheet = (sheets as any)[key];
                                const state = sheet?.state || sheet?.State;
                                const name = sheet?.name || sheet?.Name;
                                if ((state === 'hidden' || state === 'veryHidden') && name) {
                                    hiddenSheets.add(name);
                                }
                            });

                            // Also try matching by index
                            workbook.SheetNames.forEach((sheetName: string, index: number) => {
                                const sheet = (sheets as any)[index] || (sheets as any)[index.toString()];
                                if (sheet) {
                                    const state = sheet.state || sheet.State;
                                    if (state === 'hidden' || state === 'veryHidden') {
                                        hiddenSheets.add(sheetName);
                                    }
                                }
                            });
                        }
                    }

                    // Process each sheet (tab) - only include visible sheets that actually exist
                    // Check if workbook has Sheets - if not, try to use SheetNames to access sheets
                    if (!workbook.Sheets || Object.keys(workbook.Sheets).length === 0) {
                        // If Sheets is empty but SheetNames exists, try to access sheets by name
                        if (workbook.SheetNames && workbook.SheetNames.length > 0) {
                            // Try alternative approach - access sheets directly by name
                            workbook.SheetNames.forEach(sheetName => {
                                const isHidden = hiddenSheets.has(sheetName);

                                // Skip hidden sheets (they won't be processed, but we'll show them with indication)
                                // Actually, let's show them but mark them as hidden
                                // Try to get worksheet - might work even if Sheets object looks empty
                                const worksheet = (workbook as any).Sheets?.[sheetName];
                                if (!worksheet) {
                                    // Even if worksheet doesn't exist, add it if it's in SheetNames
                                    if (!isHidden) {
                                        return; // Skip non-hidden sheets that don't exist
                                    }
                                    // For hidden sheets, add with empty data
                                    tabInfos.push({
                                        tabName: sheetName,
                                        rowCount: 0,
                                        topLeftCell: 'NOT FOUND',
                                        product: '',
                                        qty: '',
                                        unit: '',
                                        remark: '',
                                        isHidden: true,
                                        columnHeaders: []
                                    });
                                    return;
                                }

                                const datatableInfo = this.findDatatableInfo(worksheet);
                                const autoSelected = this.autoSelectColumns(datatableInfo.columnHeaders);
                                tabInfos.push({
                                    tabName: sheetName,
                                    rowCount: datatableInfo.rowCount,
                                    topLeftCell: datatableInfo.topLeftCell,
                                    product: autoSelected.product || datatableInfo.product,
                                    qty: autoSelected.qty || datatableInfo.qty,
                                    unit: autoSelected.unit || datatableInfo.unit,
                                    remark: autoSelected.remark || datatableInfo.remark,
                                    isHidden: isHidden,
                                    columnHeaders: datatableInfo.columnHeaders
                                });
                            });

                            // If we successfully processed any sheets, continue
                            if (tabInfos.length > 0) {
                                const fileName = file.name.replace(/\.[^/.]+$/, '');
                                resolve({
                                    fileName: fileName,
                                    numberOfTabs: tabInfos.length,
                                    tabs: tabInfos,
                                    file: file
                                });
                                return;
                            }
                        }

                        // If we get here, we couldn't process any sheets
                        this.loggingService.logError(
                            new Error('Workbook has no Sheets object or sheets are empty'),
                            'workbook_no_sheets',
                            'RfqComponent',
                            {
                                hasSheets: !!workbook.Sheets,
                                sheetNames: workbook.SheetNames || [],
                                sheetsKeys: workbook.Sheets ? Object.keys(workbook.Sheets) : [],
                                workbookKeys: Object.keys(workbook)
                            }
                        );
                        // Return empty result instead of throwing
                        const fileName = file.name.replace(/\.[^/.]+$/, '');
                        resolve({
                            fileName: fileName,
                            numberOfTabs: 0,
                            tabs: [],
                            file: file
                        });
                        return;
                    }

                    // Get all sheet names from Sheets object (they should match SheetNames)
                    const availableSheets = Object.keys(workbook.Sheets);

                    availableSheets.forEach(sheetName => {
                        // Skip metadata sheets (sheets starting with '!')
                        if (sheetName.startsWith('!')) {
                            return;
                        }

                        const isHidden = hiddenSheets.has(sheetName);

                        // Don't skip hidden sheets - we'll show them but mark them as hidden
                        const worksheet = workbook.Sheets[sheetName];
                        if (!worksheet) {
                            // If worksheet doesn't exist but is in SheetNames, add it
                            if (isHidden) {
                                tabInfos.push({
                                    tabName: sheetName,
                                    rowCount: 0,
                                    topLeftCell: 'NOT FOUND',
                                    product: '',
                                    qty: '',
                                    unit: '',
                                    remark: '',
                                    isHidden: true,
                                    columnHeaders: []
                                });
                            }
                            return;
                        }

                        const datatableInfo = this.findDatatableInfo(worksheet);
                        const autoSelected = this.autoSelectColumns(datatableInfo.columnHeaders);
                        tabInfos.push({
                            tabName: sheetName,
                            rowCount: datatableInfo.rowCount,
                            topLeftCell: datatableInfo.topLeftCell,
                            product: autoSelected.product || datatableInfo.product,
                            qty: autoSelected.qty || datatableInfo.qty,
                            unit: autoSelected.unit || datatableInfo.unit,
                            remark: autoSelected.remark || datatableInfo.remark,
                            isHidden: isHidden,
                            columnHeaders: datatableInfo.columnHeaders
                        });
                    });

                    const fileName = file.name.replace(/\.[^/.]+$/, ''); // Remove extension

                    resolve({
                        fileName: fileName,
                        numberOfTabs: tabInfos.length, // Use count of visible tabs only
                        tabs: tabInfos,
                        file: file
                    });
                } catch (error) {
                    this.loggingService.logError(
                        error as Error,
                        'excel_file_reading',
                        'RfqComponent',
                        {
                            fileName: file.name,
                            processingStep: 'read_excel_file'
                        }
                    );
                    reject(error);
                }
            };

            reader.onerror = () => {
                const error = new Error('Failed to read file');
                this.loggingService.logError(
                    error,
                    'file_reader_error',
                    'RfqComponent',
                    {
                        fileName: file.name,
                        fileSize: file.size,
                        fileType: file.type
                    }
                );
                reject(error);
            };

            reader.readAsArrayBuffer(file);
        });
    }

    private findDatatableInfo(worksheet: XLSX.WorkSheet): { rowCount: number; topLeftCell: string; product: string; qty: string; unit: string; remark: string; columnHeaders: string[] } {
        const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1:Z100');
        let headerRow = -1;
        let descriptionColumn = -1;
        let priceColumn = -1;
        let qtyColumn = -1;
        let unitColumn = -1;
        let remarkColumn = -1;
        let topLeftCell = 'NOT FOUND';

        // Find the header row with description and price columns
        for (let row = range.s.r; row <= range.e.r; row++) {
            for (let col = range.s.c; col <= range.e.c; col++) {
                const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
                const cell = worksheet[cellAddress];

                if (cell && cell.v) {
                    const cellValue = String(cell.v).toLowerCase().trim();

                    // Look for description column
                    if ((cellValue.includes('description') ||
                        cellValue.includes('descrption') ||
                        cellValue.includes('product') ||
                        cellValue.includes('item') ||
                        cellValue.includes('name')) && descriptionColumn === -1) {
                        descriptionColumn = col;
                        headerRow = row;
                        // Top left cell is the first column of the header row
                        topLeftCell = XLSX.utils.encode_cell({ r: row, c: range.s.c });
                    }

                    // Look for price column (must be less than 25 characters)
                    if ((cellValue.includes('price') ||
                        cellValue.includes('cost') ||
                        cellValue.includes('amount') ||
                        cellValue.includes('value')) &&
                        cellValue.length < 25 &&
                        priceColumn === -1) {
                        priceColumn = col;
                        if (headerRow === -1) {
                            headerRow = row;
                            // Top left cell is the first column of the header row
                            topLeftCell = XLSX.utils.encode_cell({ r: row, c: range.s.c });
                        }
                    }

                    // Look for quantity column
                    if ((cellValue === 'qty' ||
                        cellValue === 'quantity' ||
                        cellValue.includes('qty')) && qtyColumn === -1) {
                        qtyColumn = col;
                        if (headerRow === -1) {
                            headerRow = row;
                            topLeftCell = XLSX.utils.encode_cell({ r: row, c: range.s.c });
                        }
                    }

                    // Look for unit column
                    if ((cellValue === 'unit' ||
                        cellValue === 'units' ||
                        cellValue === 'uom' ||
                        cellValue === 'uoms') && unitColumn === -1) {
                        unitColumn = col;
                        if (headerRow === -1) {
                            headerRow = row;
                            topLeftCell = XLSX.utils.encode_cell({ r: row, c: range.s.c });
                        }
                    }

                    // Look for remark/comment column
                    if ((cellValue.includes('remark') ||
                        cellValue.includes('comment')) && remarkColumn === -1) {
                        remarkColumn = col;
                        if (headerRow === -1) {
                            headerRow = row;
                            topLeftCell = XLSX.utils.encode_cell({ r: row, c: range.s.c });
                        }
                    }
                }
            }

            // If we found both description and price columns, break
            if (headerRow !== -1 && descriptionColumn !== -1 && priceColumn !== -1) {
                break;
            }
        }

        // Extract all column headers from the header row
        const columnHeaders: string[] = [];
        if (headerRow !== -1) {
            for (let col = range.s.c; col <= range.e.c; col++) {
                const headerAddress = XLSX.utils.encode_cell({ r: headerRow, c: col });
                const headerCell = worksheet[headerAddress];
                if (headerCell && headerCell.v !== null && headerCell.v !== undefined) {
                    const headerValue = String(headerCell.v).trim();
                    if (headerValue) {
                        columnHeaders.push(headerValue);
                    }
                }
            }
        }

        // If we didn't find a header row, return default values
        if (headerRow === -1 || descriptionColumn === -1 || priceColumn === -1) {
            return {
                rowCount: 0,
                topLeftCell: 'NOT FOUND',
                product: '',
                qty: '',
                unit: '',
                remark: '',
                columnHeaders: columnHeaders.length > 0 ? columnHeaders : []
            };
        }

        // Count data rows (excluding header row)
        // A row is counted if it has description data. Price can be 0, empty, or any value.
        let rowCount = 0;
        let firstDataRow = -1;

        for (let dataRow = headerRow + 1; dataRow <= range.e.r; dataRow++) {
            const descAddress = XLSX.utils.encode_cell({ r: dataRow, c: descriptionColumn });
            const descCell = worksheet[descAddress];

            // Count rows that have description data (non-empty)
            const hasDescription = descCell && descCell.v !== null && descCell.v !== undefined && String(descCell.v).trim() !== '';

            if (hasDescription) {
                if (firstDataRow === -1) {
                    firstDataRow = dataRow;
                }
                rowCount++;
            }
        }

        // Extract values from the first data row (if it exists)
        let product = '';
        let qty = '';
        let unit = '';
        let remark = '';

        if (firstDataRow !== -1) {
            // Get product/description from first data row
            if (descriptionColumn !== -1) {
                const descAddress = XLSX.utils.encode_cell({ r: firstDataRow, c: descriptionColumn });
                const descCell = worksheet[descAddress];
                if (descCell && descCell.v !== null && descCell.v !== undefined) {
                    product = String(descCell.v).trim();
                }
            }

            // Get quantity from first data row
            if (qtyColumn !== -1) {
                const qtyAddress = XLSX.utils.encode_cell({ r: firstDataRow, c: qtyColumn });
                const qtyCell = worksheet[qtyAddress];
                if (qtyCell && qtyCell.v !== null && qtyCell.v !== undefined) {
                    qty = String(qtyCell.v).trim();
                }
            }

            // Get unit from first data row
            if (unitColumn !== -1) {
                const unitAddress = XLSX.utils.encode_cell({ r: firstDataRow, c: unitColumn });
                const unitCell = worksheet[unitAddress];
                if (unitCell && unitCell.v !== null && unitCell.v !== undefined) {
                    unit = String(unitCell.v).trim();
                }
            }

            // Get remark from first data row
            if (remarkColumn !== -1) {
                const remarkAddress = XLSX.utils.encode_cell({ r: firstDataRow, c: remarkColumn });
                const remarkCell = worksheet[remarkAddress];
                if (remarkCell && remarkCell.v !== null && remarkCell.v !== undefined) {
                    remark = String(remarkCell.v).trim();
                }
            }
        }

        return { rowCount, topLeftCell, product, qty, unit, remark, columnHeaders };
    }

    removeFile(index: number): void {
        this.loggingService.logUserAction('file_removed', {
            fileName: this.fileAnalyses[index].fileName
        }, 'RfqComponent');

        this.uploadedFiles = this.uploadedFiles.filter((_, i) => i !== index);
        this.fileAnalyses.splice(index, 1);
    }

    clearAllFiles(): void {
        this.loggingService.logUserAction('clear_all_files', {
            fileCount: this.fileAnalyses.length
        }, 'RfqComponent');

        this.uploadedFiles = [];
        this.fileAnalyses = [];
    }

    onCompanyChange(company: 'HI US' | 'HI UK' | 'EOS'): void {
        this.loggingService.logButtonClick(`company_selected_${company}`, 'RfqComponent', {
            selectedCompany: company
        });
        this.selectedCompany = company;
    }

    openFile(analysis: FileAnalysis): void {
        this.loggingService.logUserAction('file_opened', {
            fileName: analysis.fileName,
            fileSize: analysis.file.size,
            fileType: analysis.file.type
        }, 'RfqComponent');

        // Create a URL for the file and open it in a new tab
        const url = URL.createObjectURL(analysis.file);
        window.open(url, '_blank');

        // Clean up the URL after a short delay to free memory
        setTimeout(() => {
            URL.revokeObjectURL(url);
        }, 1000);
    }

    onColumnChange(analysis: FileAnalysis, tab: TabInfo, columnType: 'product' | 'qty' | 'unit' | 'remark'): void {
        this.loggingService.logUserAction('column_selected', {
            fileName: analysis.fileName,
            tabName: tab.tabName,
            columnType: columnType,
            selectedValue: tab[columnType]
        }, 'RfqComponent');
    }

    private autoSelectColumns(columnHeaders: string[]): { product: string; qty: string; unit: string; remark: string } {
        const result = { product: '', qty: '', unit: '', remark: '' };

        // Create a case-insensitive lookup map
        const headerMap = new Map<string, string>();
        columnHeaders.forEach(header => {
            headerMap.set(header.toLowerCase().trim(), header);
        });

        // Product: 'Product Name', 'Description'
        const productOptions = ['Product Name', 'Description', 'Equipment Description'];
        for (const option of productOptions) {
            const found = headerMap.get(option.toLowerCase());
            if (found) {
                result.product = found;
                break;
            }
        }

        // Qty: 'Requested Qty', 'Quantity', 'Qty'
        const qtyOptions = ['Requested Qty', 'Quantity', 'Qty'];
        for (const option of qtyOptions) {
            const found = headerMap.get(option.toLowerCase());
            if (found) {
                result.qty = found;
                break;
            }
        }

        // Unit: 'Unit Type', 'Unit', 'UOM', 'UN'
        const unitOptions = ['Unit Type', 'Unit', 'UOM', 'UN'];
        for (const option of unitOptions) {
            const found = headerMap.get(option.toLowerCase());
            if (found) {
                result.unit = found;
                break;
            }
        }

        // Remark: 'Product No', 'Product No.', 'Remark', 'Impa'
        const remarkOptions = ['Product No', 'Product No.', 'Remark', 'Remarks', 'Impa'];
        for (const option of remarkOptions) {
            const found = headerMap.get(option.toLowerCase());
            if (found) {
                result.remark = found;
                break;
            }
        }

        return result;
    }
}


