import { Injectable } from '@angular/core';
import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';
import JSZip from 'jszip';
import { LoggingService } from './logging.service';
import { buildInvoiceStyleWorkbook, InvoiceWorkbookBank, InvoiceWorkbookData, InvoiceWorkbookItem } from '../utils/invoice-workbook-builder';
import { COUNTRIES, COUNTRY_PORTS } from '../constants/countries.constants';

export interface TabInfo {
    tabName: string;
    rowCount: number;
    topLeftCell: string;
    product: string;
    qty: string;
    unit: string;
    unitPrimary: string;
    unitSecondary: string;
    unitTertiary: string;
    remark: string;
    remarkPrimary: string;
    remarkSecondary: string;
    remarkTertiary: string;
    price: string;
    isHidden: boolean;
    columnHeaders: string[];
    included: boolean;
    previewHeaders: string[];
    previewRows: string[][];
    images: string[]; // Array of data URLs for images found in this worksheet
    fileType?: string; // File type (e.g., 'XLS', 'XLSX') when images cannot be extracted
    category: string;
}

export interface FileAnalysis {
    fileName: string;
    numberOfTabs: number;
    tabs: TabInfo[];
    file: File;
}

export interface ProposalItem {
    pos: number;
    fileName: string;
    tabName: string;
    description: string;
    remark: string;
    unit: string;
    qty: string;
    price?: string;
    total?: string;
}

export interface ProposalTable {
    fileName: string;
    tabName: string;
    rowCount: number;
    items: ProposalItem[];
    category: string;
}

export type CurrencyCode = 'GBP' | 'USD' | 'EUR' | 'AUD' | 'NZD' | 'CAD';

export const CURRENCY_SYMBOL_MAP: Record<CurrencyCode, string> = {
    GBP: '£',
    USD: '$',
    EUR: '€',
    AUD: 'A$',
    NZD: 'NZ$',
    CAD: 'C$'
};

export interface RfqData {
    // Our Company Details
    ourCompanyName: string;
    ourCompanyAddress: string;
    ourCompanyAddress2: string;
    ourCompanyCity: string;
    ourCompanyCountry: string;
    ourCompanyPhone: string;
    ourCompanyEmail: string;
    // Billing Information
    vesselName: string;
    vesselName2: string;
    vesselAddress: string;
    vesselAddress2: string;
    vesselCity: string;
    vesselCountry: string;
    // Bank Details
    bankName: string;
    bankAddress: string;
    iban: string;
    swiftCode: string;
    accountTitle: string;
    accountNumber: string;
    sortCode: string;
    achRouting?: string;
    intermediaryBic?: string;
    // Invoice/Quotation Details
    invoiceNumber: string;
    invoiceDate: string;
    vessel: string;
    country: string;
    port: string;
    category: string;
    invoiceDue: string;
}

@Injectable({
    providedIn: 'root'
})
export class RfqStateService {
    uploadedFiles: File[] = [];
    fileAnalyses: FileAnalysis[] = [];
    isProcessing = false;
    errorMessage = '';
    selectedCompany: 'HI US' | 'HI UK' | 'EOS' = 'HI US';
    selectedCurrency: CurrencyCode = 'USD';

    rfqData: RfqData = {
        ourCompanyName: '',
        ourCompanyAddress: '',
        ourCompanyAddress2: '',
        ourCompanyCity: '',
        ourCompanyCountry: '',
        ourCompanyPhone: '',
        ourCompanyEmail: '',
        vesselName: '',
        vesselName2: '',
        vesselAddress: '',
        vesselAddress2: '',
        vesselCity: '',
        vesselCountry: '',
        bankName: '',
        bankAddress: '',
        iban: '',
        swiftCode: '',
        accountTitle: '',
        accountNumber: '',
        sortCode: '',
        achRouting: '',
        intermediaryBic: '',
        invoiceNumber: 'Invoice ',
        invoiceDate: this.getTodayDate(),
        vessel: '',
        country: '',
        port: '',
        category: 'Provisions',
        invoiceDue: ''
    };

    countries = COUNTRIES;

    availablePorts: string[] = [];

    countryPorts: { [key: string]: string[] } = COUNTRY_PORTS;

    private readonly fallbackColumnOptions = [
        'COLUMN A', 'COLUMN B', 'COLUMN C', 'COLUMN D', 'COLUMN E', 'COLUMN F',
        'COLUMN G', 'COLUMN H', 'COLUMN I', 'COLUMN J', 'COLUMN K', 'COLUMN L'
    ];

    proposalItems: ProposalItem[] = [];
    proposalTables: ProposalTable[] = [];
    private proposalRefreshToken = 0;

    previewDialogVisible = false;
    previewDialogHeaders: string[] = [];
    previewDialogRows: string[][] = [];
    previewDialogFileName = '';
    previewDialogTabName = '';
    private previewDialogHighlightIndexes: number[] = [];

    private proposalExportFileName = '';

    constructor(private readonly loggingService: LoggingService) {
        this.onCompanySelectionChange();
    }

    setSelectedCurrency(currency: CurrencyCode): void {
        this.selectedCurrency = currency;
    }

    setProposalExportFileName(fileName: string | null | undefined): void {
        this.proposalExportFileName = (fileName ?? '').trim();
    }

    getProposalExportFileName(): string {
        return this.proposalExportFileName;
    }

    getSelectedCurrencySymbol(): string {
        return CURRENCY_SYMBOL_MAP[this.selectedCurrency] ?? '$';
    }

    private getTodayDate(): string {
        const today = new Date();
        return today.toISOString().split('T')[0];
    }

    handleFiles(files: File[]): void {
        const excelFiles = files.filter(file => {
            const validTypes = [
                'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                'application/vnd.ms-excel',
                'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.macroEnabled'
            ];
            return validTypes.includes(file.type) || file.name.match(/\.(xlsx|xls|xlsm)$/i);
        });

        if (excelFiles.length === 0) {
            this.errorMessage = 'Please upload valid Excel files (.xlsx, .xls, or .xlsm)';
            return;
        }

        this.errorMessage = '';

        excelFiles.forEach(file => {
            this.loggingService.logFileUpload(file.name, file.size, file.type, 'rfq', 'RfqStateService');
        });

        this.uploadedFiles = [...this.uploadedFiles, ...excelFiles];
        void this.processFiles(excelFiles);
    }

    private async processFiles(files: File[]): Promise<void> {
        this.isProcessing = true;

        try {
            for (const file of files) {
                const analysis = await this.analyzeExcelFile(file);
                this.fileAnalyses.push(analysis);
            }
            await this.refreshProposalPreview();
        } catch (error) {
            this.errorMessage = 'Error processing Excel files. Please ensure they are valid Excel files.';
            this.loggingService.logError(
                error as Error,
                'excel_file_processing',
                'RfqStateService',
                {
                    fileCount: files.length
                }
            );
        } finally {
            this.isProcessing = false;
        }
    }

    private async analyzeExcelFile(file: File): Promise<FileAnalysis> {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();

            reader.onload = async (e: any) => {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, {
                        type: 'array',
                        cellFormula: false,
                        cellHTML: false,
                        cellStyles: false,
                        sheetStubs: false,
                        cellText: true,
                        cellDates: true
                    });

                    if (!workbook) {
                        throw new Error('Failed to read workbook - workbook is null or undefined');
                    }

                    if (!workbook.Sheets && !workbook.SheetNames) {
                        this.loggingService.logError(
                            new Error('Workbook has no Sheets or SheetNames'),
                            'workbook_structure_invalid',
                            'RfqStateService',
                            {
                                fileName: file.name,
                                workbookKeys: workbook ? Object.keys(workbook) : [],
                                hasWorkbook: !!workbook
                            }
                        );
                        throw new Error('Invalid workbook structure - workbook, Sheets, or SheetNames missing');
                    }

                    // Extract images from Excel file (only works for XLSX format)
                    // Check file signature: XLSX files start with PK (ZIP signature)
                    const isXlsx = this.isXlsxFormat(data);
                    const fileType = this.getFileType(file.name, isXlsx);
                    const imagesBySheet = await this.extractImagesFromExcel(data, workbook.SheetNames, isXlsx);

                    const tabInfos: TabInfo[] = [];

                    const hiddenSheets = new Set<string>();

                    if (workbook.Workbook && workbook.Workbook.Sheets) {
                        const sheets = workbook.Workbook.Sheets;

                        if (Array.isArray(sheets)) {
                            sheets.forEach((sheet: any, index: number) => {
                                const state = sheet?.state || (sheet as any)?.State;
                                const name = sheet?.name || (sheet as any)?.Name || workbook.SheetNames[index];
                                if ((state === 'hidden' || state === 'veryHidden') && name) {
                                    hiddenSheets.add(name);
                                }
                            });
                        } else {
                            Object.keys(sheets).forEach(key => {
                                const sheet = (sheets as any)[key];
                                const state = sheet?.state || sheet?.State;
                                const name = sheet?.name || sheet?.Name;
                                if ((state === 'hidden' || state === 'veryHidden') && name) {
                                    hiddenSheets.add(name);
                                }
                            });

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

                    if (!workbook.Sheets || Object.keys(workbook.Sheets).length === 0) {
                        workbook.SheetNames.forEach(sheetName => {
                            const worksheet = workbook.Sheets[sheetName];
                            if (worksheet) {
                                const tabInfo = this.analyzeWorksheet(sheetName, worksheet, hiddenSheets.has(sheetName), file, workbook, imagesBySheet[sheetName] || [], fileType);
                                tabInfos.push(tabInfo);
                            }
                        });
                    } else {
                        workbook.SheetNames.forEach(sheetName => {
                            const worksheet = workbook.Sheets[sheetName];
                            if (worksheet) {
                                const tabInfo = this.analyzeWorksheet(sheetName, worksheet, hiddenSheets.has(sheetName), file, workbook, imagesBySheet[sheetName] || [], fileType);
                                tabInfos.push(tabInfo);
                            }
                        });
                    }

                    const fileAnalysis: FileAnalysis = {
                        fileName: file.name.replace(/\.[^/.]+$/, ''),
                        numberOfTabs: tabInfos.length,
                        tabs: tabInfos,
                        file
                    };

                    resolve(fileAnalysis);

                } catch (error) {
                    reject(error);
                }
            };

            reader.onerror = () => reject(new Error('Failed to read file'));
            reader.readAsArrayBuffer(file);
        });
    }

    private analyzeWorksheet(sheetName: string, worksheet: XLSX.WorkSheet, isHidden: boolean, file: File, workbook: XLSX.WorkBook, images: string[] = [], fileType?: string): TabInfo {
        const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1:Z100');

        let firstDataCell = this.findFirstDataCell(worksheet, range);

        const columnHeaders: string[] = [];
        const topLeftCell = firstDataCell || 'A1';

        if (firstDataCell) {
            const cellRef = XLSX.utils.decode_cell(firstDataCell);
            const headerRow = cellRef.r;
            const startCol = cellRef.c;

            for (let col = startCol; col <= range.e.c; col++) {
                const cellAddress = XLSX.utils.encode_cell({ r: headerRow, c: col });
                const cell = worksheet[cellAddress];

                if (cell && cell.v !== undefined && cell.v !== null) {
                    const value = String(cell.v).trim();
                    if (value !== '') {
                        columnHeaders.push(value);
                    }
                }
            }
        }

        const fileNameWithoutExt = file.name.replace(/\.[^/.]+$/, '');
        const nameParts = fileNameWithoutExt.split(/[\s_-]+/);
        const defaultCategory = nameParts.length > 0 ? nameParts[nameParts.length - 1] : 'Provisions';

        const tabInfo: TabInfo = {
            tabName: sheetName,
            rowCount: 0,
            topLeftCell,
            product: '',
            qty: '',
            unit: '',
            unitPrimary: '',
            unitSecondary: '',
            unitTertiary: '',
            remark: '',
            remarkPrimary: '',
            remarkSecondary: '',
            remarkTertiary: '',
            price: '',
            isHidden,
            columnHeaders,
            included: true,
            previewHeaders: [],
            previewRows: [],
            images: images,
            fileType: fileType,
            category: defaultCategory
        };

        if (columnHeaders.length > 0) {
            const autoSelection = this.autoSelectColumns(columnHeaders);
            tabInfo.product = autoSelection.product;
            tabInfo.qty = autoSelection.qty;
            tabInfo.unitPrimary = autoSelection.unit;
            tabInfo.remarkPrimary = autoSelection.remark;
            tabInfo.price = autoSelection.price;
            this.syncCompositeField(tabInfo, 'unit');
            this.syncCompositeField(tabInfo, 'remark');
        }

        tabInfo.rowCount = this.countDataRows(worksheet, tabInfo, range);
        const preview = this.buildPreviewData(worksheet, tabInfo, range);
        tabInfo.previewHeaders = preview.headers;
        tabInfo.previewRows = preview.rows;
        this.updateInclusionState(tabInfo);

        return tabInfo;
    }

    private findFirstDataCell(worksheet: XLSX.WorkSheet, range: XLSX.Range): string | undefined {
        const keywords = ['description', 'product', 'item', 'name'];
        const maxRowsToScan = Math.min(range.e.r, 50);
        const maxColsToScan = Math.min(range.e.c, 25);

        for (let row = range.s.r; row <= maxRowsToScan; row++) {
            for (let col = range.s.c; col <= maxColsToScan; col++) {
                const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
                const cell = worksheet[cellAddress];
                if (cell && cell.v !== undefined && cell.v !== null && String(cell.v).trim() !== '') {
                    const cellValue = String(cell.v).trim().toLowerCase();
                    if (keywords.some(keyword => cellValue.includes(keyword))) {
                        return cellAddress;
                    }
                }
            }
        }

        return undefined;
    }

    private countDataRows(worksheet: XLSX.WorkSheet, tab: TabInfo, range: XLSX.Range): number {
        if (!tab.topLeftCell) {
            return 0;
        }

        const cellRef = XLSX.utils.decode_cell(tab.topLeftCell);
        const headerRow = cellRef.r;
        const startCol = cellRef.c;

        let descriptionColumn = -1;
        const descriptionKeywords = ['description', 'product', 'item', 'name'];

        for (let i = 0; i < tab.columnHeaders.length; i++) {
            const header = tab.columnHeaders[i].toLowerCase();
            if (descriptionKeywords.some(keyword => header.includes(keyword))) {
                descriptionColumn = startCol + i;
                break;
            }
        }

        let rowCount = 0;
        let consecutiveEmptyRows = 0;

        if (descriptionColumn !== -1) {
            for (let dataRow = headerRow + 1; dataRow <= range.e.r; dataRow++) {
                const descAddress = XLSX.utils.encode_cell({ r: dataRow, c: descriptionColumn });
                const descCell = worksheet[descAddress];
                const hasDescription = descCell && descCell.v !== null && descCell.v !== undefined && String(descCell.v).trim() !== '';
                if (hasDescription) {
                    rowCount++;
                    consecutiveEmptyRows = 0;
                } else {
                    consecutiveEmptyRows++;
                    if (consecutiveEmptyRows >= 3) {
                        break;
                    }
                }
            }
        } else {
            for (let dataRow = headerRow + 1; dataRow <= range.e.r; dataRow++) {
                let hasData = false;
                for (let col = startCol; col <= startCol + tab.columnHeaders.length - 1 && col <= range.e.c; col++) {
                    const cellAddress = XLSX.utils.encode_cell({ r: dataRow, c: col });
                    const cell = worksheet[cellAddress];
                    if (cell && cell.v !== null && cell.v !== undefined && String(cell.v).trim() !== '') {
                        hasData = true;
                        break;
                    }
                }
                if (hasData) {
                    rowCount++;
                    consecutiveEmptyRows = 0;
                } else {
                    consecutiveEmptyRows++;
                    if (consecutiveEmptyRows >= 3) {
                        break;
                    }
                }
            }
        }

        return rowCount;
    }

    async reanalyzeTabWithTopLeft(analysis: FileAnalysis, tab: TabInfo, topLeftCell: string): Promise<void> {
        try {
            const previousIncludedState = tab.included;

            const fileData = await this.readFileAsArrayBuffer(analysis.file);
            const workbook = XLSX.read(fileData, {
                type: 'array',
                cellFormula: false,
                cellHTML: false,
                cellStyles: false,
                sheetStubs: false,
                cellText: true,
                cellDates: true
            });

            const worksheet = workbook.Sheets[tab.tabName];
            if (!worksheet) {
                throw new Error(`Worksheet "${tab.tabName}" not found`);
            }

            const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1:Z100');
            const cellRef = XLSX.utils.decode_cell(topLeftCell);
            const headerRow = cellRef.r;
            const startCol = cellRef.c;

            const columnHeaders: string[] = [];

            for (let col = startCol; col <= range.e.c; col++) {
                const cellAddress = XLSX.utils.encode_cell({ r: headerRow, c: col });
                const cell = worksheet[cellAddress];
                if (cell && cell.v !== undefined && cell.v !== null) {
                    const value = String(cell.v).trim();
                    if (value !== '') {
                        columnHeaders.push(value);
                    }
                }
            }

            tab.columnHeaders = columnHeaders;

            const autoSelection = this.autoSelectColumns(columnHeaders);
            tab.product = autoSelection.product;
            tab.qty = autoSelection.qty;
            tab.unitPrimary = autoSelection.unit;
            tab.unitSecondary = '';
            tab.unitTertiary = '';
            tab.remarkPrimary = autoSelection.remark;
            tab.remarkSecondary = '';
            tab.remarkTertiary = '';
            tab.price = autoSelection.price;
            this.syncCompositeField(tab, 'remark');
            this.syncCompositeField(tab, 'unit');

            tab.topLeftCell = topLeftCell;

            const updatedRange = XLSX.utils.decode_range(worksheet['!ref'] || 'A1:Z100');
            tab.rowCount = this.countDataRows(worksheet, tab, updatedRange);

            const preview = this.buildPreviewData(worksheet, tab, updatedRange);
            tab.previewHeaders = preview.headers;
            tab.previewRows = preview.rows;

            this.updateInclusionState(tab);
            tab.included = previousIncludedState;

            await this.refreshProposalPreview();

            this.loggingService.logUserAction('tab_reanalyzed', {
                fileName: analysis.fileName,
                tabName: tab.tabName,
                topLeftCell: topLeftCell,
                columnCount: columnHeaders.length,
                rowCount: tab.rowCount
            }, 'RfqStateService');

        } catch (error) {
            this.loggingService.logError(
                error as Error,
                'tab_reanalysis_error',
                'RfqStateService',
                {
                    fileName: analysis.fileName,
                    tabName: tab.tabName,
                    topLeftCell: topLeftCell
                }
            );
            throw error;
        }
    }

    private readFileAsArrayBuffer(file: File): Promise<Uint8Array> {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e: any) => {
                try {
                    resolve(new Uint8Array(e.target.result));
                } catch (error) {
                    reject(error);
                }
            };
            reader.onerror = () => reject(new Error('Failed to read file'));
            reader.readAsArrayBuffer(file);
        });
    }

    private async extractImagesFromExcel(fileData: Uint8Array, sheetNames: string[], isXlsx: boolean = true): Promise<Record<string, string[]>> {
        const imagesBySheet: Record<string, string[]> = {};
        
        // Initialize empty arrays for all sheets
        sheetNames.forEach(sheetName => {
            imagesBySheet[sheetName] = [];
        });

        // XLS files are binary format, not ZIP archives, so we can't extract images using JSZip
        if (!isXlsx) {
            this.loggingService.logUserAction('xls_image_extraction_skipped', {
                reason: 'XLS files use binary format, image extraction not supported'
            }, 'RfqStateService');
            return imagesBySheet;
        }

        try {
            // Load the Excel file as a ZIP archive (only works for XLSX/XLSM files)
            const zip = await JSZip.loadAsync(fileData);
            
            // Extract all images from xl/media/ folder
            const mediaFilePromises: Array<Promise<{ name: string; data: Blob }>> = [];
            const allPaths: string[] = [];
            
            // Iterate through all files in the ZIP using Object.keys
            Object.keys(zip.files).forEach(relativePath => {
                const file = zip.files[relativePath];
                allPaths.push(relativePath);
                if (!file.dir && relativePath.startsWith('xl/media/')) {
                    const fileName = relativePath.replace('xl/media/', '');
                    // Check if it's an image file
                    if (fileName.match(/\.(png|jpg|jpeg|gif|bmp|webp)$/i)) {
                        mediaFilePromises.push(
                            file.async('blob').then(blob => ({ name: fileName, data: blob }))
                        );
                    }
                }
            });

            // Log for debugging
            const mediaPaths = allPaths.filter(p => p.startsWith('xl/media/'));
            this.loggingService.logUserAction('image_extraction_debug', {
                totalFiles: allPaths.length,
                mediaFiles: mediaPaths.length,
                mediaPaths: mediaPaths.slice(0, 10), // Log first 10
                imagePromises: mediaFilePromises.length
            }, 'RfqStateService');

            // Wait for all image blobs to be loaded
            if (mediaFilePromises.length > 0) {
                const loadedMediaFiles = await Promise.all(mediaFilePromises);
                
                this.loggingService.logUserAction('images_loaded', {
                    count: loadedMediaFiles.length
                }, 'RfqStateService');
                
                // Convert blobs to data URLs
                for (const mediaFile of loadedMediaFiles) {
                    try {
                        const dataUrl = await this.blobToDataUrl(mediaFile.data);
                        
                        // For now, assign images to all sheets
                        // A more accurate implementation would parse worksheet XML to match images to specific sheets
                        sheetNames.forEach(sheetName => {
                            imagesBySheet[sheetName].push(dataUrl);
                        });
                    } catch (blobError) {
                        this.loggingService.logError(
                            blobError as Error,
                            'blob_to_dataurl_error',
                            'RfqStateService',
                            { fileName: mediaFile.name }
                        );
                    }
                }
            } else {
                this.loggingService.logUserAction('no_images_found', {
                    checkedPaths: mediaPaths.length
                }, 'RfqStateService');
            }
            
        } catch (error) {
            // If image extraction fails, just continue without images
            this.loggingService.logError(
                error as Error,
                'image_extraction_error',
                'RfqStateService',
                { errorMessage: (error as Error).message, stack: (error as Error).stack }
            );
        }

        return imagesBySheet;
    }

    private isXlsxFormat(data: Uint8Array): boolean {
        // XLSX files are ZIP archives and start with PK signature (0x50 0x4B)
        // XLS files are binary BIFF format and start with different bytes
        if (data.length < 4) {
            return false;
        }
        // Check for ZIP signature: PK (0x50 0x4B) at the start
        return data[0] === 0x50 && data[1] === 0x4B;
    }

    private getFileType(fileName: string, isXlsx: boolean): string | undefined {
        // If it's XLSX format, images can be extracted, so no fileType needed
        if (isXlsx) {
            return undefined;
        }
        
        // Determine file type from extension
        const extension = fileName.toLowerCase().split('.').pop();
        if (extension === 'xls') {
            return 'XLS';
        } else if (extension === 'xlsm') {
            // XLSM might be detected as XLSX if it's actually ZIP format
            return undefined;
        }
        
        // For any other format where images can't be extracted
        return extension ? extension.toUpperCase() : 'UNKNOWN';
    }

    private blobToDataUrl(blob: Blob): Promise<string> {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = () => resolve(reader.result as string);
            reader.onerror = reject;
            reader.readAsDataURL(blob);
        });
    }

    removeFile(index: number): void {
        this.uploadedFiles = this.uploadedFiles.filter((_, i) => i !== index);
        this.fileAnalyses.splice(index, 1);
        void this.refreshProposalPreview();
    }

    clearAllFiles(): void {
        this.loggingService.logUserAction('clear_all_files', {
            fileCount: this.fileAnalyses.length
        }, 'RfqStateService');

        this.uploadedFiles = [];
        this.fileAnalyses = [];
        this.proposalItems = [];
        this.proposalTables = [];
        this.proposalRefreshToken++;
    }

    isTabReady(tab: TabInfo): boolean {
        if (!tab.included) {
            return false;
        }

        const hasTopLeft = !!tab.topLeftCell && tab.topLeftCell.trim() !== '';
        const hasRows = tab.rowCount > 0;
        const hasProduct = !!tab.product && tab.product.trim() !== '';
        const hasQty = !!tab.qty && tab.qty.trim() !== '';
        const hasUnit = this.hasValue(tab.unitPrimary);

        return hasTopLeft && hasRows && hasProduct && hasQty && hasUnit;
    }

    canCreateRFQs(): boolean {
        if (this.fileAnalyses.length === 0) {
            return false;
        }

        const includedTabs: { analysis: FileAnalysis; tab: TabInfo }[] = [];
        for (const analysis of this.fileAnalyses) {
            for (const tab of analysis.tabs) {
                if (tab.included) {
                    includedTabs.push({ analysis, tab });
                }
            }
        }

        if (includedTabs.length === 0) {
            return false;
        }

        for (const { tab } of includedTabs) {
            if (!this.isTabReady(tab)) {
                return false;
            }
        }

        return true;
    }

    async createRFQs(): Promise<void> {
        this.loggingService.logUserAction('create_rfqs_clicked', {
            company: this.selectedCompany,
            tableCount: this.proposalTables.length
        }, 'RfqStateService');

        const tablesToExport = this.proposalTables.filter(table => table.items.length > 0);
        if (tablesToExport.length === 0) {
            this.loggingService.logUserAction('create_rfqs_no_tables', {
                reason: 'no_tables_ready'
            }, 'RfqStateService');
            return;
        }

        const selectedBank = this.mapSelectedCompanyToBank();

        for (const table of tablesToExport) {
            try {
                const tableCurrency = this.determineProposalPrimaryCurrency(table.items);
                const fileNameOverride = this.getResolvedProposalExportFileName(table) ?? this.buildProposalFileName(table);
                const workbookData = this.buildProposalWorkbookData(table, tableCurrency, fileNameOverride);

                const { blob, fileName } = await buildInvoiceStyleWorkbook({
                    data: workbookData,
                    selectedBank,
                    primaryCurrency: tableCurrency,
                    appendAtoInvoiceNumber: false,
                    includeFees: false,
                    fileNameOverride
                });

                saveAs(blob, fileName);

                this.loggingService.logExport('proposal_rfq_created', {
                    fileName,
                    fileSize: blob.size,
                    tabName: table.tabName,
                    rowCount: table.rowCount,
                    company: this.selectedCompany
                }, 'RfqStateService');
            } catch (error) {
                this.loggingService.logError(
                    error as Error,
                    'proposal_rfq_export',
                    'RfqStateService',
                    {
                        tabName: table.tabName,
                        fileName: table.fileName,
                        company: this.selectedCompany
                    }
                );
            }
        }
    }

    private getCellString(worksheet: XLSX.WorkSheet, row: number, col: number): string {
        const cell = worksheet[XLSX.utils.encode_cell({ r: row, c: col })];
        if (!cell || cell.v === undefined || cell.v === null) {
            return '';
        }
        return String(cell.v).trim();
    }

    private mapSelectedCompanyToBank(): InvoiceWorkbookBank {
        switch (this.selectedCompany) {
            case 'HI US':
                return 'US';
            case 'HI UK':
                return 'UK';
            case 'EOS':
                return 'EOS';
            default:
                return 'US';
        }
    }

    private determineProposalPrimaryCurrency(_items: ProposalItem[]): string {
        return this.getSelectedCurrencySymbol();
    }

    private detectCurrencyFromString(value: unknown): string | null {
        if (value === undefined || value === null) {
            return null;
        }

        const str = String(value).trim();
        if (!str) {
            return null;
        }

        const upper = str.toUpperCase();
        if (upper.includes('NZ$') || upper.includes('NZD')) return 'NZ$';
        if (upper.includes('A$') || upper.includes('AUD')) return 'A$';
        if (upper.includes('C$') || upper.includes('CAD')) return 'C$';
        if (str.includes('€') || upper.includes('EUR')) return '€';
        if (str.includes('£') || upper.includes('GBP')) return '£';
        if (str.includes('$')) {
            if (!upper.includes('NZ$') && !upper.includes('A$') && !upper.includes('C$')) {
                return '$';
            }
        }
        if (upper.includes('USD')) return '$';
        return null;
    }

    private parseNumericFromMixed(value: unknown): number | null {
        if (value === undefined || value === null) {
            return null;
        }

        if (typeof value === 'number') {
            return Number.isNaN(value) ? null : value;
        }

        if (typeof value === 'string') {
            let cleaned = value.trim();
            if (!cleaned) {
                return null;
            }
            cleaned = cleaned.replace(/NZ\$/gi, '');
            cleaned = cleaned.replace(/A\$/gi, '');
            cleaned = cleaned.replace(/C\$/gi, '');
            cleaned = cleaned.replace(/[€£$,]/g, '');
            cleaned = cleaned.replace(/USD/gi, '');
            cleaned = cleaned.replace(/NZD/gi, '');
            cleaned = cleaned.replace(/AUD/gi, '');
            cleaned = cleaned.replace(/CAD/gi, '');
            cleaned = cleaned.trim();
            if (!cleaned) {
                return null;
            }
            const parsed = parseFloat(cleaned);
            return Number.isNaN(parsed) ? null : parsed;
        }

        return null;
    }

    private mapProposalItemToWorkbookItem(item: ProposalItem, fallbackCurrency: string): InvoiceWorkbookItem {
        const preferredCurrency = this.getSelectedCurrencySymbol();
        const detectedCurrency = this.detectCurrencyFromString(item.price) || this.detectCurrencyFromString(item.total);
        const currency = preferredCurrency || detectedCurrency || fallbackCurrency || '£';

        const remarkTrimmed = (item.remark ?? '').trim();
        const unitTrimmed = (item.unit ?? '').trim();
        const shouldBlankNumericColumns = remarkTrimmed === '' && unitTrimmed === '';

        const qtyParsed = this.parseNumericFromMixed(item.qty);
        let qtyValue: number | string;
        if (shouldBlankNumericColumns) {
            qtyValue = '';
        } else if (qtyParsed !== null) {
            qtyValue = qtyParsed;
        } else if (typeof item.qty === 'string') {
            qtyValue = item.qty.trim();
        } else {
            qtyValue = '';
        }

        const priceParsed = shouldBlankNumericColumns ? 0 : (this.parseNumericFromMixed(item.price) ?? 0);
        const totalParsed = shouldBlankNumericColumns ? 0 : (this.parseNumericFromMixed(item.total) ?? 0);

        return {
            pos: item.pos,
            description: item.description,
            remark: item.remark,
            unit: item.unit,
            qty: qtyValue,
            price: priceParsed,
            total: totalParsed,
            currency,
            blankNumericColumns: shouldBlankNumericColumns
        };
    }

    private sanitizeFileNameSegment(segment: string): string {
        return segment.replace(/[<>:"/\\|?*]/g, '_');
    }

    private buildProposalFileName(table: ProposalTable): string {
        const parts: string[] = ['Proposal'];

        if (this.rfqData.invoiceNumber?.trim()) {
            parts.push(this.rfqData.invoiceNumber.trim());
        }
        if (this.rfqData.vessel?.trim()) {
            parts.push(this.rfqData.vessel.trim());
        }
        if (table.fileName?.trim()) {
            parts.push(table.fileName.trim());
        }
        if (table.tabName?.trim()) {
            parts.push(table.tabName.trim());
        }
        parts.push(this.selectedCompany.replace(/\s+/g, ''));

        return this.sanitizeFileNameSegment(parts.filter(Boolean).join('_'));
    }

    private buildProposalWorkbookData(table: ProposalTable, primaryCurrency: string, fileNameBase: string): InvoiceWorkbookData {
        const workbookItems: InvoiceWorkbookItem[] = table.items.map(item => this.mapProposalItemToWorkbookItem(item, primaryCurrency));

        return {
            items: workbookItems,
            discountPercent: 0,
            deliveryFee: 0,
            portFee: 0,
            agencyFee: 0,
            transportCustomsLaunchFees: 0,
            launchFee: 0,
            ourCompanyName: this.rfqData.ourCompanyName,
            ourCompanyAddress: this.rfqData.ourCompanyAddress,
            ourCompanyAddress2: this.rfqData.ourCompanyAddress2,
            ourCompanyCity: this.rfqData.ourCompanyCity,
            ourCompanyCountry: this.rfqData.ourCompanyCountry,
            ourCompanyPhone: this.rfqData.ourCompanyPhone,
            ourCompanyEmail: this.rfqData.ourCompanyEmail,
            vesselName: this.rfqData.vesselName,
            vesselName2: this.rfqData.vesselName2,
            vesselAddress: this.rfqData.vesselAddress,
            vesselAddress2: this.rfqData.vesselAddress2,
            vesselCity: this.rfqData.vesselCity,
            vesselCountry: this.rfqData.vesselCountry,
            bankName: this.rfqData.bankName,
            bankAddress: this.rfqData.bankAddress,
            iban: this.rfqData.iban,
            swiftCode: this.rfqData.swiftCode,
            accountTitle: this.rfqData.accountTitle,
            accountNumber: this.rfqData.accountNumber,
            sortCode: this.rfqData.sortCode,
            achRouting: this.rfqData.achRouting,
            intermediaryBic: this.rfqData.intermediaryBic,
            invoiceNumber: this.rfqData.invoiceNumber,
            invoiceDate: this.rfqData.invoiceDate,
            vessel: this.rfqData.vessel,
            country: this.rfqData.country,
            port: this.rfqData.port,
            category: table.category,
            invoiceDue: this.rfqData.invoiceDue,
            exportFileName: fileNameBase
        };
    }

    private getResolvedProposalExportFileName(table: ProposalTable): string | null {
        let trimmed = this.proposalExportFileName.trim();
        if (!trimmed) {
            return null;
        }

        // Replace <category> placeholder with actual category
        trimmed = trimmed.replace(/<category>/gi, table.category || '');

        const withoutExtension = trimmed.replace(/\.xlsx$/i, '');
        const sanitized = this.sanitizeFileNameSegment(withoutExtension);
        return sanitized ? sanitized : null;
    }

    updateTableCategory(fileName: string, tabName: string, category: string): void {
        for (const analysis of this.fileAnalyses) {
            if (analysis.fileName === fileName) {
                const tab = analysis.tabs.find(t => t.tabName === tabName);
                if (tab) {
                    tab.category = category;
                    // We don't need to refresh proposal preview here because we are just updating a property
                    // that is passed through to ProposalTable. The UI already updated the ProposalTable object.
                    // However, if we want to ensure consistency if refresh happens later, this is enough.
                    
                    // Also update the current proposalTable instance to match, although binding should handle it
                    const proposalTable = this.proposalTables.find(pt => pt.fileName === fileName && pt.tabName === tabName);
                    if (proposalTable) {
                        proposalTable.category = category;
                    }
                }
                break;
            }
        }
    }

    private async refreshProposalPreview(): Promise<void> {
        const refreshId = ++this.proposalRefreshToken;

        const tables: ProposalTable[] = [];
        const flatItems: ProposalItem[] = [];

        for (const analysis of this.fileAnalyses) {
            const readyTabs = analysis.tabs.filter(tab => this.isTabReady(tab));
            if (readyTabs.length === 0) {
                continue;
            }

            const fileData = await this.readFileAsArrayBuffer(analysis.file);
            const workbook = XLSX.read(fileData, {
                type: 'array',
                cellFormula: false,
                cellHTML: false,
                cellStyles: false,
                sheetStubs: false,
                cellText: true,
                cellDates: true
            });

            for (const tab of readyTabs) {
                const worksheet = workbook.Sheets[tab.tabName];
                if (!worksheet) {
                    continue;
                }

                const cellRef = XLSX.utils.decode_cell(tab.topLeftCell);
                const headerRow = cellRef.r;
                const startCol = cellRef.c;
                const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1:Z100');

                const unitSelectionNames = this.getCompositeSelections(tab, 'unit');
                const remarkSelectionNames = this.getCompositeSelections(tab, 'remark');
                const productCol = this.findColumnIndex(worksheet, headerRow, startCol, tab.product);
                const qtyCol = this.findColumnIndex(worksheet, headerRow, startCol, tab.qty);
                const unitCols = unitSelectionNames.map(selection => this.findColumnIndex(worksheet, headerRow, startCol, selection));
                const remarkCols = remarkSelectionNames.map(selection => this.findColumnIndex(worksheet, headerRow, startCol, selection));
                const hasPriceSelected = this.hasValue(tab.price);
                const priceCol = hasPriceSelected ? this.findColumnIndex(worksheet, headerRow, startCol, tab.price) : -1;

                const hasMissingColumn =
                    productCol === -1 ||
                    qtyCol === -1 ||
                    unitCols.length === 0 ||
                    unitCols.some(index => index === -1) ||
                    (remarkSelectionNames.length > 0 && remarkCols.some(index => index === -1)) ||
                    (hasPriceSelected && priceCol === -1);

                if (hasMissingColumn) {
                    tables.push({
                        fileName: analysis.fileName,
                        tabName: tab.tabName,
                        rowCount: 0,
                        items: [],
                        category: tab.category
                    });
                    continue;
                }

                const tableItems: ProposalItem[] = [];
                let position = 1;
                let consecutiveEmptyRows = 0;

                for (let row = headerRow + 1; row <= range.e.r; row++) {
                    const description = this.getCellString(worksheet, row, productCol);
                    const qty = this.getCellString(worksheet, row, qtyCol);
                    const unitValues = unitCols.map(col => this.getCellString(worksheet, row, col));
                    const unit = this.combineCellValues(unitValues);
                    const remark = remarkCols.length > 0
                        ? this.combineCellValues(remarkCols.map(col => this.getCellString(worksheet, row, col)))
                        : '';
                    const price = priceCol !== -1 ? this.getCellString(worksheet, row, priceCol) : '';

                    const hasAnyValue = description !== '' || qty !== '' || unit !== '' || remark !== '' || price !== '';
                    if (!hasAnyValue) {
                        consecutiveEmptyRows++;
                        if (consecutiveEmptyRows >= 3) {
                            break;
                        }
                        continue;
                    }

                    consecutiveEmptyRows = 0;

                    if (description === '') {
                        continue;
                    }

                    const item: ProposalItem = {
                        pos: position++,
                        fileName: analysis.fileName,
                        tabName: tab.tabName,
                        description,
                        remark,
                        unit,
                        qty,
                        price,
                        total: ''
                    };

                    tableItems.push(item);
                    flatItems.push({ ...item });
                }

                tables.push({
                    fileName: analysis.fileName,
                    tabName: tab.tabName,
                    rowCount: tableItems.length,
                    items: tableItems,
                    category: tab.category
                });
            }
        }

        if (refreshId === this.proposalRefreshToken) {
            if (tables.length === 0) {
                this.proposalItems = [];
                this.proposalTables = [];
            } else {
                this.proposalItems = flatItems;
                this.proposalTables = tables;
            }
        }
    }

    private findColumnIndex(worksheet: XLSX.WorkSheet, headerRow: number, startCol: number, headerName: string): number {
        if (!headerName || headerName.trim() === '') {
            return -1;
        }

        const trimmedHeader = headerName.trim();
        const columnMatch = trimmedHeader.match(/^column\s+([A-Z]{1,2})$/i);
        if (columnMatch) {
            const columnLetter = columnMatch[1].toUpperCase();
            try {
                return XLSX.utils.decode_col(columnLetter);
            } catch {
                return -1;
            }
        }

        const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1:Z100');
        for (let col = startCol; col <= range.e.c; col++) {
            const headerAddress = XLSX.utils.encode_cell({ r: headerRow, c: col });
            const headerCell = worksheet[headerAddress];
            if (headerCell && headerCell.v !== undefined && headerCell.v !== null) {
                const cellValue = String(headerCell.v).trim();
                if (cellValue.toLowerCase() === trimmedHeader.toLowerCase()) {
                    return col;
                }
            }
        }

        return -1;
    }

    onCompanyChange(company: 'HI US' | 'HI UK' | 'EOS'): void {
        this.loggingService.logButtonClick(`company_selected_${company}`, 'RfqStateService', {
            selectedCompany: company
        });
        this.selectedCompany = company;
        this.onCompanySelectionChange();
    }

    onCompanySelectionChange(): void {
        this.clearBankDetails();

        switch (this.selectedCompany) {
            case 'HI UK':
                this.populateUKBankDetails();
                break;
            case 'HI US':
                this.populateUSBankDetails();
                break;
            case 'EOS':
                this.populateEOSBankDetails();
                break;
            default:
                break;
        }
    }

    private clearBankDetails(): void {
        this.rfqData.ourCompanyName = '';
        this.rfqData.ourCompanyAddress = '';
        this.rfqData.ourCompanyAddress2 = '';
        this.rfqData.ourCompanyCity = '';
        this.rfqData.ourCompanyCountry = '';
        this.rfqData.ourCompanyPhone = '';
        this.rfqData.ourCompanyEmail = '';

        this.rfqData.vesselName = '';
        this.rfqData.vesselName2 = '';
        this.rfqData.vesselAddress = '';
        this.rfqData.vesselAddress2 = '';
        this.rfqData.vesselCity = '';
        this.rfqData.vesselCountry = '';

        this.rfqData.bankName = '';
        this.rfqData.bankAddress = '';
        this.rfqData.iban = '';
        this.rfqData.swiftCode = '';
        this.rfqData.accountTitle = '';
        this.rfqData.accountNumber = '';
        this.rfqData.sortCode = '';
        this.rfqData.achRouting = '';
        this.rfqData.intermediaryBic = '';
    }

    private populateUKBankDetails(): void {
        this.rfqData.ourCompanyName = 'HI MARINE COMPANY LIMITED';
        this.rfqData.ourCompanyAddress = '167-169 Great Portland Street';
        this.rfqData.ourCompanyAddress2 = '';
        this.rfqData.ourCompanyCity = 'London, London, W1W 5PF';
        this.rfqData.ourCompanyCountry = 'United Kingdom';
        this.rfqData.ourCompanyPhone = '';
        this.rfqData.ourCompanyEmail = 'office@himarinecompany.com';

        this.rfqData.vesselName = '';
        this.rfqData.vesselName2 = '';
        this.rfqData.vesselAddress = '';
        this.rfqData.vesselAddress2 = '';
        this.rfqData.vesselCity = '';
        this.rfqData.vesselCountry = '';

        this.rfqData.bankName = 'Lloyds Bank plc';
        this.rfqData.bankAddress = '6 Market Place, Oldham, OL11JG, United Kingdom';
        this.rfqData.iban = 'GB84LOYD30962678553260';
        this.rfqData.swiftCode = 'LOYDGB21446';
        this.rfqData.accountTitle = 'HI MARINE COMPANY LIMITED';
        this.rfqData.accountNumber = '78553260';
        this.rfqData.sortCode = '30-96-26';
    }

    private populateUSBankDetails(): void {
        this.rfqData.ourCompanyName = 'HI MARINE COMPANY INC.';
        this.rfqData.ourCompanyAddress = '9407 N.E. Vancouver Mall Drive, Suite 104';
        this.rfqData.ourCompanyAddress2 = '';
        this.rfqData.ourCompanyCity = 'Vancouver, WA  98662';
        this.rfqData.ourCompanyCountry = 'USA';
        this.rfqData.ourCompanyPhone = '+1 857 2045786';
        this.rfqData.ourCompanyEmail = 'office@himarinecompany.com';

        this.rfqData.vesselName = '';
        this.rfqData.vesselName2 = '';
        this.rfqData.vesselAddress = '';
        this.rfqData.vesselAddress2 = '';
        this.rfqData.vesselCity = '';
        this.rfqData.vesselCountry = '';

        this.rfqData.bankName = 'Bank of America';
        this.rfqData.bankAddress = '100 West 33d Street New York, New York 10001';
        this.rfqData.accountNumber = '466002755612';
        this.rfqData.swiftCode = 'BofAUS3N';
        this.rfqData.achRouting = '011000138';
        this.rfqData.accountTitle = 'Hi Marine Company Inc.';
    }

    private populateEOSBankDetails(): void {
        this.rfqData.ourCompanyName = 'EOS SUPPLY LTD';
        this.rfqData.ourCompanyAddress = '85 Great Portland Street, First Floor';
        this.rfqData.ourCompanyAddress2 = '';
        this.rfqData.ourCompanyCity = 'London, England, W1W 7LT';
        this.rfqData.ourCompanyCountry = 'United Kingdom';
        this.rfqData.ourCompanyPhone = '';
        this.rfqData.ourCompanyEmail = '';

        this.rfqData.vesselName = '';
        this.rfqData.vesselName2 = '';
        this.rfqData.vesselAddress = '';
        this.rfqData.vesselAddress2 = '';
        this.rfqData.vesselCity = '';
        this.rfqData.vesselCountry = '';

        this.rfqData.bankName = 'Revolut Ltd';
        this.rfqData.bankAddress = '7 Westferry Circus, London, England, E14 4HD';
        this.rfqData.iban = 'GB64REVO00996912321885';
        this.rfqData.swiftCode = 'REVOGB21XXX';
        this.rfqData.intermediaryBic = 'CHASGB2L';
        this.rfqData.accountTitle = 'EOS SUPPLY LTD';
        this.rfqData.accountNumber = '69340501';
        this.rfqData.sortCode = '04-00-75';
    }

    onCountryChange(): void {
        this.rfqData.port = '';

        if (this.rfqData.country && this.countryPorts[this.rfqData.country]) {
            this.availablePorts = this.countryPorts[this.rfqData.country];
        } else {
            this.availablePorts = [];
        }
    }

    openFile(analysis: FileAnalysis): void {
        this.loggingService.logUserAction('file_opened', {
            fileName: analysis.fileName,
            fileSize: analysis.file.size,
            fileType: analysis.file.type
        }, 'RfqStateService');

        const url = URL.createObjectURL(analysis.file);
        window.open(url, '_blank');
        setTimeout(() => {
            URL.revokeObjectURL(url);
        }, 1000);
    }

    onColumnChange(analysis: FileAnalysis, tab: TabInfo, columnType: 'product' | 'qty' | 'unit' | 'remark' | 'price'): void {
        void this.refreshProposalPreview();
    }

    getTopLeftCellOptions(): string[] {
        const options: string[] = [];
        for (let i = 1; i <= 25; i++) {
            options.push('A' + i);
        }
        return options;
    }

    getTopLeftCellOptionsLimited(): string[] {
        return this.getTopLeftCellOptions().slice(0, 10);
    }

    getColumnOptions(tab: TabInfo): string[] {
        const headers = (tab.columnHeaders ?? []).map(header => header.trim()).filter(Boolean);

        const normalizedHeaders = new Set(headers.map(header => header.toLowerCase()));
        const fallback = this.fallbackColumnOptions.filter(option => !normalizedHeaders.has(option.toLowerCase()));

        return [...headers, ...fallback];
    }

    onCompositeSelectionChanged(analysis: FileAnalysis, tab: TabInfo, columnType: 'unit' | 'remark', level: 0 | 1 | 2): void {
        if (columnType === 'unit') {
            if (!this.hasValue(tab.unitPrimary)) {
                tab.unitSecondary = '';
                tab.unitTertiary = '';
            } else if (level === 1 && !this.hasValue(tab.unitSecondary)) {
                tab.unitTertiary = '';
            }
        } else {
            if (!this.hasValue(tab.remarkPrimary)) {
                tab.remarkSecondary = '';
                tab.remarkTertiary = '';
            } else if (level === 1 && !this.hasValue(tab.remarkSecondary)) {
                tab.remarkTertiary = '';
            }
        }

        this.syncCompositeField(tab, columnType);
        this.onColumnChange(analysis, tab, columnType);
    }

    onTopLeftCellChange(event: Event, tab: TabInfo): void {
        const input = event.target as HTMLInputElement;
        let value = input.value.toUpperCase();

        const regex = /^[A-Z][0-9]{1,2}$/;
        if (value && !regex.test(value)) {
            value = value.replace(/[^A-Z0-9]/g, '');
            if (value && !/^[A-Z]/.test(value)) {
                value = '';
            }
            const match = value.match(/^([A-Z])([0-9]{0,2})/);
            if (match) {
                value = match[0];
            } else if (value.length > 0 && /^[A-Z]/.test(value)) {
                value = value.charAt(0);
            }
        }

        input.value = value;
        tab.topLeftCell = value;

        if (value && value.trim() !== '') {
            input.removeAttribute('list');

            for (const analysis of this.fileAnalyses) {
                if (analysis.tabs.includes(tab)) {
                    void this.reanalyzeTabWithTopLeft(analysis, tab, value);
                    break;
                }
            }
        } else {
            tab.columnHeaders = [];
            tab.product = '';
            tab.qty = '';
            tab.unitPrimary = '';
            tab.unitSecondary = '';
            tab.unitTertiary = '';
            tab.unit = '';
            tab.remark = '';
            tab.remarkPrimary = '';
            tab.remarkSecondary = '';
            tab.remarkTertiary = '';
            tab.price = '';
            tab.rowCount = 0;
            tab.previewHeaders = [];
            tab.previewRows = [];
            this.syncCompositeField(tab, 'unit');
            this.syncCompositeField(tab, 'remark');
        }
    }

    onIncludeChange(tab: TabInfo): void {
        this.loggingService.logUserAction('include_toggled', {
            tabName: tab.tabName,
            included: tab.included
        }, 'RfqStateService');

        void this.refreshProposalPreview();
    }

    private updateInclusionState(tab: TabInfo): void {
        tab.included = !(tab.rowCount === 0 || !tab.topLeftCell || tab.topLeftCell.trim() === '');
    }

    private hasValue(value: string | undefined | null): boolean {
        return !!(value && value.trim() !== '');
    }

    private getCompositeSelections(tab: TabInfo, columnType: 'unit' | 'remark'): string[] {
        if (columnType === 'unit') {
            return [tab.unitPrimary, tab.unitSecondary, tab.unitTertiary]
                .filter((selection): selection is string => this.hasValue(selection));
        }

        return [tab.remarkPrimary, tab.remarkSecondary, tab.remarkTertiary]
            .filter((selection): selection is string => this.hasValue(selection));
    }

    private combineCompositeSelections(values: Array<string | undefined>): string {
        return values
            .map(value => (value ?? '').trim())
            .filter(value => value !== '')
            .join(' ');
    }

    private syncCompositeField(tab: TabInfo, columnType: 'unit' | 'remark'): void {
        const selections = this.getCompositeSelections(tab, columnType);
        const combined = this.combineCompositeSelections(selections);

        if (columnType === 'unit') {
            tab.unit = combined;
        } else {
            tab.remark = combined;
        }
    }

    private combineCellValues(values: string[]): string {
        return this.combineCompositeSelections(values);
    }

    private buildPreviewData(worksheet: XLSX.WorkSheet, tab: TabInfo, range: XLSX.Range): { headers: string[]; rows: string[][] } {
        if (!tab.topLeftCell) {
            return { headers: [], rows: [] };
        }

        const headers = [...(tab.columnHeaders ?? [])];
        if (headers.length === 0) {
            return { headers, rows: [] };
        }

        const cellRef = XLSX.utils.decode_cell(tab.topLeftCell);
        const headerRow = cellRef.r;
        const startCol = cellRef.c;
        const maxRows = Math.min(tab.rowCount, 4);
        const rows: string[][] = [];
        let consecutiveEmptyRows = 0;

        for (let offset = 1; offset <= range.e.r - headerRow && rows.length < maxRows; offset++) {
            const currentRow = headerRow + offset;
            const rowData: string[] = [];
            let hasData = false;

            for (let colIndex = 0; colIndex < headers.length; colIndex++) {
                const column = startCol + colIndex;
                if (column > range.e.c) {
                    rowData.push('');
                    continue;
                }

                const cellAddress = XLSX.utils.encode_cell({ r: currentRow, c: column });
                const cell = worksheet[cellAddress];
                const value = cell?.v != null ? String(cell.v) : '';
                if (value.trim() !== '') {
                    hasData = true;
                }
                rowData.push(value);
            }

            if (!hasData && rowData.every(value => value.trim() === '')) {
                consecutiveEmptyRows++;
                if (consecutiveEmptyRows >= 3) {
                    break;
                }
            } else {
                consecutiveEmptyRows = 0;
            }

            rows.push(rowData);
        }

        return { headers, rows };
    }

    openPreviewDialog(analysis: FileAnalysis, tab: TabInfo): void {
        if (!tab.previewRows || tab.previewRows.length === 0) {
            return;
        }

        this.previewDialogHeaders = [...tab.previewHeaders];
        this.previewDialogRows = tab.previewRows.map(row => [...row]);
        this.previewDialogFileName = analysis.fileName;
        this.previewDialogTabName = tab.tabName;

        const highlightTargets = new Set<string>();
        if (this.hasValue(tab.product)) {
            highlightTargets.add(tab.product.trim().toLowerCase());
        }
        this.getCompositeSelections(tab, 'unit').forEach(value => highlightTargets.add(value.trim().toLowerCase()));
        this.getCompositeSelections(tab, 'remark').forEach(value => highlightTargets.add(value.trim().toLowerCase()));
        if (this.hasValue(tab.qty)) {
            highlightTargets.add(tab.qty.trim().toLowerCase());
        }
        if (this.hasValue(tab.price)) {
            highlightTargets.add(tab.price.trim().toLowerCase());
        }

        this.previewDialogHighlightIndexes = [];
        this.previewDialogHeaders.forEach((header, index) => {
            if (highlightTargets.has(header.trim().toLowerCase())) {
                this.previewDialogHighlightIndexes.push(index);
            }
        });

        this.previewDialogVisible = true;
    }

    isPreviewColumnHighlighted(index: number): boolean {
        return this.previewDialogHighlightIndexes.includes(index);
    }

    closePreviewDialog(): void {
        this.previewDialogVisible = false;
        this.previewDialogHighlightIndexes = [];
    }

    onTopLeftCellFocus(event: Event, fileIndex: number, tabIndex: number): void {
        const input = event.target as HTMLInputElement;
        if (!input.value || input.value.trim() === '') {
            input.setAttribute('list', `topLeftList-${fileIndex}-${tabIndex}`);
        }
    }

    onTopLeftCellSelectChange(event: Event, tab: TabInfo): void {
        const select = event.target as HTMLSelectElement;
        const value = select.value;

        this.loggingService.logButtonClick('top_left_cell_select_change', 'RfqStateService', {
            value: value
        });

        if (value === 'custom') {
            tab.topLeftCell = '';
            setTimeout(() => {
                const input = select.closest('td')?.querySelector('.top-left-input') as HTMLInputElement;
                if (input) {
                    input.focus();
                }
            }, 0);
        }
    }

    validateTopLeftCell(tab: TabInfo): void {
        const regex = /^[A-Z][0-9]{1,2}$/;
        if (tab.topLeftCell && !regex.test(tab.topLeftCell)) {
            if (tab.topLeftCell.trim() !== '') {
                const match = tab.topLeftCell.match(/^([A-Z])([0-9]{1,2})/);
                if (match) {
                    tab.topLeftCell = match[0];
                } else {
                    tab.topLeftCell = '';
                }
            }
        }
    }

    private autoSelectColumns(columnHeaders: string[]): { product: string; qty: string; unit: string; remark: string; price: string } {
        const result = { product: '', qty: '', unit: '', remark: '', price: '' };

        // Helper function to remove all punctuation from a string
        const removePunctuation = (str: string): string => {
            return str.replace(/[^\w\s]/g, '');
        };

        // Normalize headers: lowercase + trim + remove punctuation
        const headerMap = new Map<string, string>();
        columnHeaders.forEach(header => {
            const normalized = removePunctuation(header.toLowerCase().trim());
            headerMap.set(normalized, header);
        });

        const productOptions = ['Product Name', 'Description', 'Equipment Description', 'Item', 'Items'];
        for (const option of productOptions) {
            const normalizedOption = removePunctuation(option.toLowerCase());
            const found = headerMap.get(normalizedOption);
            if (found) {
                result.product = found;
                break;
            }
        }

        const qtyOptions = ['Requested Qty', 'Quantity', 'Qty'];
        for (const option of qtyOptions) {
            const normalizedOption = removePunctuation(option.toLowerCase());
            const found = headerMap.get(normalizedOption);
            if (found) {
                result.qty = found;
                break;
            }
        }

        const unitOptions = ['Unit Type', 'Unit', 'UOM', 'UN'];
        for (const option of unitOptions) {
            const normalizedOption = removePunctuation(option.toLowerCase());
            const found = headerMap.get(normalizedOption);
            if (found) {
                result.unit = found;
                break;
            }
        }

        const remarkOptions = ['Product No', 'Product No.', 'Remark', 'Remarks', 'Impa'];
        for (const option of remarkOptions) {
            const normalizedOption = removePunctuation(option.toLowerCase());
            const found = headerMap.get(normalizedOption);
            if (found) {
                result.remark = found;
                break;
            }
        }

        const priceOptions = ['Price'];
        for (const option of priceOptions) {
            const normalizedOption = removePunctuation(option.toLowerCase());
            const found = headerMap.get(normalizedOption);
            if (found) {
                result.price = found;
                break;
            }
        }

        return result;
    }
}


