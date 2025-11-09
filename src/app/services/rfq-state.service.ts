import { Injectable } from '@angular/core';
import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';
import { LoggingService } from './logging.service';
import { buildInvoiceStyleWorkbook, InvoiceWorkbookBank, InvoiceWorkbookData, InvoiceWorkbookItem } from '../utils/invoice-workbook-builder';

export interface TabInfo {
    tabName: string;
    rowCount: number;
    topLeftCell: string;
    product: string;
    qty: string;
    unit: string;
    remark: string;
    isHidden: boolean;
    columnHeaders: string[];
    included: boolean;
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
}

export interface RfqData {
    // Our Company Details
    ourCompanyName: string;
    ourCompanyAddress: string;
    ourCompanyAddress2: string;
    ourCompanyCity: string;
    ourCompanyCountry: string;
    ourCompanyPhone: string;
    ourCompanyEmail: string;
    // Vessel Details
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
        invoiceNumber: '',
        invoiceDate: this.getTodayDate(),
        vessel: '',
        country: '',
        port: '',
        category: 'Provisions',
        invoiceDue: ''
    };

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

    availablePorts: string[] = [];

    countryPorts: { [key: string]: string[] } = {
        'Australia': ['Sydney', 'Melbourne', 'Brisbane', 'Perth', 'Adelaide', 'Darwin', 'Townsville', 'Newcastle'],
        'Belgium': ['Antwerp', 'Ghent', 'Zeebrugge', 'Brussels', 'Liège', 'Ostend'],
        'Brazil': ['Santos', 'Rio de Janeiro', 'Paranaguá', 'Rio Grande', 'Salvador', 'Fortaleza', 'Recife', 'Vitória'],
        'Canada': ['Vancouver', 'Montreal', 'Halifax', 'Toronto', 'Thunder Bay', 'Saint John', 'Hamilton', 'Quebec City'],
        'China': ['Shanghai', 'Shenzhen', 'Ningbo', 'Qingdao', 'Guangzhou', 'Tianjin', 'Dalian', 'Xiamen', 'Lianyungang', 'Yingkou'],
        'Denmark': ['Copenhagen', 'Aarhus', 'Aalborg', 'Esbjerg', 'Fredericia', 'Helsingør'],
        'Finland': ['Helsinki', 'Turku', 'Kotka', 'Hamina', 'Vaasa', 'Oulu'],
        'France': ['Le Havre', 'Marseille', 'Dunkirk', 'Calais', 'Rouen', 'Nantes', 'La Rochelle', 'Bordeaux'],
        'Germany': ['Hamburg', 'Bremen', 'Wilhelmshaven', 'Lübeck', 'Rostock', 'Kiel', 'Emden', 'Cuxhaven'],
        'Greece': ['Piraeus', 'Thessaloniki', 'Patras', 'Volos', 'Kavala', 'Igoumenitsa'],
        'India': ['Jawaharlal Nehru (Mumbai)', 'Chennai', 'Kolkata', 'Cochin', 'Visakhapatnam', 'Kandla', 'Paradip', 'Tuticorin'],
        'Italy': ['Genoa', 'La Spezia', 'Naples', 'Venice', 'Trieste', 'Livorno', 'Bari', 'Taranto'],
        'Japan': ['Tokyo', 'Yokohama', 'Nagoya', 'Osaka', 'Kobe', 'Chiba', 'Kitakyushu', 'Hakata'],
        'Netherlands': ['Rotterdam', 'Amsterdam', 'Vlissingen', 'Terneuzen', 'IJmuiden', 'Delfzijl'],
        'Norway': ['Oslo', 'Bergen', 'Stavanger', 'Trondheim', 'Tromsø', 'Kristiansand', 'Drammen'],
        'Poland': ['Gdansk', 'Gdynia', 'Szczecin', 'Świnoujście', 'Kołobrzeg'],
        'Russia': ['St. Petersburg', 'Novorossiysk', 'Vladivostok', 'Kaliningrad', 'Murmansk', 'Arkhangelsk', 'Rostov-on-Don'],
        'Singapore': ['Singapore'],
        'South Korea': ['Busan', 'Incheon', 'Ulsan', 'Gwangyang', 'Pyeongtaek', 'Gunsan'],
        'Spain': ['Barcelona', 'Valencia', 'Algeciras', 'Bilbao', 'Las Palmas', 'Vigo', 'Santander', 'Cartagena'],
        'Sweden': ['Gothenburg', 'Stockholm', 'Malmö', 'Helsingborg', 'Gävle', 'Sundsvall'],
        'Turkey': ['Istanbul', 'Izmir', 'Mersin', 'Samsun', 'Trabzon', 'Iskenderun', 'Bandırma'],
        'United Kingdom': ['Felixstowe', 'Southampton', 'London Gateway', 'Liverpool', 'Immingham', 'Hull', 'Bristol', 'Portsmouth', 'Dover', 'Harwich'],
        'United States': ['Los Angeles', 'Long Beach', 'New York/New Jersey', 'Savannah', 'Seattle', 'Oakland', 'Charleston', 'Norfolk', 'Miami', 'Houston'],
        'UAE': ['Jebel Ali (Dubai)', 'Abu Dhabi', 'Sharjah', 'Fujairah', 'Ras Al Khaimah'],
        'South Africa': ['Durban', 'Cape Town', 'Port Elizabeth', 'Richards Bay', 'East London'],
        'Egypt': ['Alexandria', 'Port Said', 'Suez', 'Damietta', 'Safaga'],
        'Saudi Arabia': ['Jeddah', 'Dammam', 'Yanbu', 'Jubail', 'Jizan'],
        'Malaysia': ['Port Klang', 'Tanjung Pelepas', 'Penang', 'Johor', 'Kuantan'],
        'Thailand': ['Laem Chabang', 'Bangkok', 'Map Ta Phut', 'Songkhla'],
        'Vietnam': ['Ho Chi Minh City', 'Hai Phong', 'Da Nang', 'Cai Mep', 'Quy Nhon'],
        'Indonesia': ['Tanjung Priok (Jakarta)', 'Surabaya', 'Belawan (Medan)', 'Semarang', 'Makassar'],
        'Philippines': ['Manila', 'Cebu', 'Davao', 'Cagayan de Oro', 'Iloilo'],
        'Chile': ['Valparaíso', 'San Antonio', 'Iquique', 'Antofagasta', 'Talcahuano'],
        'Argentina': ['Buenos Aires', 'Rosario', 'Bahía Blanca', 'Mar del Plata', 'Necochea'],
        'Mexico': ['Manzanillo', 'Lázaro Cárdenas', 'Veracruz', 'Altamira', 'Ensenada'],
        'Morocco': ['Casablanca', 'Tangier', 'Agadir', 'Mohammedia', 'Safi'],
        'Israel': ['Haifa', 'Ashdod', 'Eilat'],
        'Iran': ['Bandar Abbas', 'Bandar Imam Khomeini', 'Bushehr', 'Chabahar'],
        'Pakistan': ['Karachi', 'Port Qasim', 'Gwadar'],
        'Bangladesh': ['Chittagong', 'Mongla'],
        'Sri Lanka': ['Colombo', 'Hambantota'],
        'New Zealand': ['Auckland', 'Tauranga', 'Wellington', 'Lyttelton', 'Otago'],
        'Ireland': ['Dublin', 'Cork', 'Shannon Foynes', 'Waterford'],
        'Portugal': ['Sines', 'Leixões', 'Lisbon', 'Setúbal', 'Aveiro'],
        'Romania': ['Constanta', 'Galati', 'Braila'],
        'Bulgaria': ['Varna', 'Burgas'],
        'Croatia': ['Rijeka', 'Split', 'Zadar', 'Ploče'],
        'Ukraine': ['Odessa', 'Mariupol', 'Chornomorsk', 'Mykolaiv'],
        'Estonia': ['Tallinn', 'Muuga'],
        'Latvia': ['Riga', 'Ventspils'],
        'Lithuania': ['Klaipėda'],
        'Cyprus': ['Limassol', 'Larnaca'],
        'Malta': ['Valletta', 'Marsaxlokk'],
        'Iceland': ['Reykjavik', 'Akureyri'],
        'Algeria': ['Algiers', 'Oran', 'Annaba', 'Skikda'],
        'Tunisia': ['Tunis', 'Sfax', 'Bizerte'],
        'Libya': ['Tripoli', 'Benghazi', 'Misrata'],
        'Nigeria': ['Lagos', 'Port Harcourt', 'Warri', 'Calabar'],
        'Ghana': ['Tema', 'Takoradi'],
        'Kenya': ['Mombasa'],
        'Tanzania': ['Dar es Salaam'],
        'Oman': ['Sohar', 'Muscat', 'Salalah'],
        'Qatar': ['Doha', 'Ras Laffan'],
        'Kuwait': ['Kuwait City', 'Shuwaikh', 'Shuaiba'],
        'Bahrain': ['Khalifa Bin Salman', 'Mina Salman'],
        'Jordan': ['Aqaba'],
        'Lebanon': ['Beirut', 'Tripoli']
    };

    private readonly fallbackColumnOptions = [
        'COLUMN A', 'COLUMN B', 'COLUMN C', 'COLUMN D', 'COLUMN E', 'COLUMN F',
        'COLUMN G', 'COLUMN H', 'COLUMN I', 'COLUMN J', 'COLUMN K', 'COLUMN L'
    ];

    proposalItems: ProposalItem[] = [];
    proposalTables: ProposalTable[] = [];
    private proposalRefreshToken = 0;

    constructor(private readonly loggingService: LoggingService) {
        this.onCompanySelectionChange();
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
                                const tabInfo = this.analyzeWorksheet(sheetName, worksheet, hiddenSheets.has(sheetName), file, workbook);
                                tabInfos.push(tabInfo);
                            }
                        });
                    } else {
                        workbook.SheetNames.forEach(sheetName => {
                            const worksheet = workbook.Sheets[sheetName];
                            if (worksheet) {
                                const tabInfo = this.analyzeWorksheet(sheetName, worksheet, hiddenSheets.has(sheetName), file, workbook);
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

    private analyzeWorksheet(sheetName: string, worksheet: XLSX.WorkSheet, isHidden: boolean, file: File, workbook: XLSX.WorkBook): TabInfo {
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

        const tabInfo: TabInfo = {
            tabName: sheetName,
            rowCount: 0,
            topLeftCell,
            product: '',
            qty: '',
            unit: '',
            remark: '',
            isHidden,
            columnHeaders,
            included: true
        };

        if (columnHeaders.length > 0) {
            const autoSelection = this.autoSelectColumns(columnHeaders);
            tabInfo.product = autoSelection.product;
            tabInfo.qty = autoSelection.qty;
            tabInfo.unit = autoSelection.unit;
            tabInfo.remark = autoSelection.remark;
        }

        tabInfo.rowCount = this.countDataRows(worksheet, tabInfo, range);
        this.updateInclusionState(tabInfo);

        this.loggingService.logUserAction('sheet_analyzed', {
            fileName: file.name,
            sheetName,
            topLeftCell,
            columnCount: columnHeaders.length,
            rowCount: tabInfo.rowCount,
            isHidden
        }, 'RfqStateService');

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

        if (descriptionColumn !== -1) {
            for (let dataRow = headerRow + 1; dataRow <= range.e.r; dataRow++) {
                const descAddress = XLSX.utils.encode_cell({ r: dataRow, c: descriptionColumn });
                const descCell = worksheet[descAddress];
                const hasDescription = descCell && descCell.v !== null && descCell.v !== undefined && String(descCell.v).trim() !== '';
                if (hasDescription) {
                    rowCount++;
                } else {
                    break;
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
                } else {
                    break;
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
            tab.unit = autoSelection.unit;
            tab.remark = autoSelection.remark;

            tab.topLeftCell = topLeftCell;

            const updatedRange = XLSX.utils.decode_range(worksheet['!ref'] || 'A1:Z100');
            tab.rowCount = this.countDataRows(worksheet, tab, updatedRange);

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

    removeFile(index: number): void {
        this.loggingService.logUserAction('file_removed', {
            fileName: this.fileAnalyses[index].fileName
        }, 'RfqStateService');

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
        const hasUnit = !!tab.unit && tab.unit.trim() !== '';
        const hasRemark = !!tab.remark && tab.remark.trim() !== '';

        return hasTopLeft && hasRows && hasProduct && hasQty && hasUnit && hasRemark;
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
                const fileNameBase = this.buildProposalFileName(table);
                const workbookData = this.buildProposalWorkbookData(table, tableCurrency, fileNameBase);

                const { blob, fileName } = await buildInvoiceStyleWorkbook({
                    data: workbookData,
                    selectedBank,
                    primaryCurrency: tableCurrency,
                    categoryOverride: table.tabName,
                    appendAtoInvoiceNumber: false,
                    includeFees: false,
                    fileNameOverride: fileNameBase
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

    private determineProposalPrimaryCurrency(items: ProposalItem[]): string {
        for (const item of items) {
            const currency = this.detectCurrencyFromString(item.price) || this.detectCurrencyFromString(item.total);
            if (currency) {
                return currency;
            }
        }
        return '£';
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
        const detectedCurrency = this.detectCurrencyFromString(item.price) || this.detectCurrencyFromString(item.total);
        const currency = detectedCurrency || fallbackCurrency || '£';

        const qtyParsed = this.parseNumericFromMixed(item.qty);
        let qtyValue: number | string;
        if (qtyParsed !== null) {
            qtyValue = qtyParsed;
        } else if (typeof item.qty === 'string') {
            qtyValue = item.qty.trim();
        } else {
            qtyValue = '';
        }

        const priceParsed = this.parseNumericFromMixed(item.price) ?? 0;
        const totalParsed = this.parseNumericFromMixed(item.total) ?? 0;

        return {
            pos: item.pos,
            description: item.description,
            remark: item.remark,
            unit: item.unit,
            qty: qtyValue,
            price: priceParsed,
            total: totalParsed,
            currency
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
            category: table.tabName || this.rfqData.category,
            invoiceDue: this.rfqData.invoiceDue,
            exportFileName: fileNameBase
        };
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

                const productCol = this.findColumnIndex(worksheet, headerRow, startCol, tab.product);
                const qtyCol = this.findColumnIndex(worksheet, headerRow, startCol, tab.qty);
                const unitCol = this.findColumnIndex(worksheet, headerRow, startCol, tab.unit);
                const remarkCol = this.findColumnIndex(worksheet, headerRow, startCol, tab.remark);

                if ([productCol, qtyCol, unitCol, remarkCol].some(index => index === -1)) {
                    tables.push({
                        fileName: analysis.fileName,
                        tabName: tab.tabName,
                        rowCount: 0,
                        items: []
                    });
                    continue;
                }

                const tableItems: ProposalItem[] = [];
                let position = 1;

                for (let row = headerRow + 1; row <= range.e.r; row++) {
                    const description = this.getCellString(worksheet, row, productCol);
                    const qty = this.getCellString(worksheet, row, qtyCol);
                    const unit = this.getCellString(worksheet, row, unitCol);
                    const remark = this.getCellString(worksheet, row, remarkCol);

                    const hasAnyValue = description !== '' || qty !== '' || unit !== '' || remark !== '';
                    if (!hasAnyValue) {
                        break;
                    }

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
                        price: '',
                        total: ''
                    };

                    tableItems.push(item);
                    flatItems.push({ ...item });
                }

                tables.push({
                    fileName: analysis.fileName,
                    tabName: tab.tabName,
                    rowCount: tableItems.length,
                    items: tableItems
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

    onColumnChange(analysis: FileAnalysis, tab: TabInfo, columnType: 'product' | 'qty' | 'unit' | 'remark'): void {
        this.loggingService.logUserAction('column_selected', {
            fileName: analysis.fileName,
            tabName: tab.tabName,
            columnType: columnType,
            selectedValue: tab[columnType]
        }, 'RfqStateService');

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
            tab.unit = '';
            tab.remark = '';
            tab.rowCount = 0;
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

    private autoSelectColumns(columnHeaders: string[]): { product: string; qty: string; unit: string; remark: string } {
        const result = { product: '', qty: '', unit: '', remark: '' };

        const headerMap = new Map<string, string>();
        columnHeaders.forEach(header => {
            headerMap.set(header.toLowerCase().trim(), header);
        });

        const productOptions = ['Product Name', 'Description', 'Equipment Description'];
        for (const option of productOptions) {
            const found = headerMap.get(option.toLowerCase());
            if (found) {
                result.product = found;
                break;
            }
        }

        const qtyOptions = ['Requested Qty', 'Quantity', 'Qty'];
        for (const option of qtyOptions) {
            const found = headerMap.get(option.toLowerCase());
            if (found) {
                result.qty = found;
                break;
            }
        }

        const unitOptions = ['Unit Type', 'Unit', 'UOM', 'UN'];
        for (const option of unitOptions) {
            const found = headerMap.get(option.toLowerCase());
            if (found) {
                result.unit = found;
                break;
            }
        }

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


