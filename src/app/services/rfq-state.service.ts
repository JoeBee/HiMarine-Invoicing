import { Injectable } from '@angular/core';
import * as XLSX from 'xlsx';
import * as ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import { LoggingService } from './logging.service';

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
    excluded: boolean;
}

export interface FileAnalysis {
    fileName: string;
    numberOfTabs: number;
    tabs: TabInfo[];
    file: File;
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
            excluded: false
        };

        if (columnHeaders.length > 0) {
            const autoSelection = this.autoSelectColumns(columnHeaders);
            tabInfo.product = autoSelection.product;
            tabInfo.qty = autoSelection.qty;
            tabInfo.unit = autoSelection.unit;
            tabInfo.remark = autoSelection.remark;
        }

        tabInfo.rowCount = this.countDataRows(worksheet, tabInfo, range);
        this.updateExcludedState(tabInfo);

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
            const previousExcludedState = tab.excluded;

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

            this.updateExcludedState(tab);
            tab.excluded = previousExcludedState;

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
    }

    clearAllFiles(): void {
        this.loggingService.logUserAction('clear_all_files', {
            fileCount: this.fileAnalyses.length
        }, 'RfqStateService');

        this.uploadedFiles = [];
        this.fileAnalyses = [];
    }

    canCreateRFQs(): boolean {
        if (this.fileAnalyses.length === 0) {
            return false;
        }

        const nonExcludedTabs: { analysis: FileAnalysis; tab: TabInfo }[] = [];
        for (const analysis of this.fileAnalyses) {
            for (const tab of analysis.tabs) {
                if (!tab.excluded) {
                    nonExcludedTabs.push({ analysis, tab });
                }
            }
        }

        if (nonExcludedTabs.length === 0) {
            return false;
        }

        for (const { tab } of nonExcludedTabs) {
            if (!tab.product || tab.product.trim() === '' ||
                !tab.qty || tab.qty.trim() === '' ||
                !tab.unit || tab.unit.trim() === '' ||
                !tab.remark || tab.remark.trim() === '') {
                return false;
            }
        }

        return true;
    }

    createRFQs(): void {
        this.loggingService.logUserAction('create_rfqs_clicked', {
            company: this.selectedCompany,
            fileCount: this.fileAnalyses.length
        }, 'RfqStateService');

        const nonExcludedTabs: { analysis: FileAnalysis; tab: TabInfo }[] = [];
        for (const analysis of this.fileAnalyses) {
            for (const tab of analysis.tabs) {
                if (!tab.excluded) {
                    nonExcludedTabs.push({ analysis, tab });
                }
            }
        }

        for (const { analysis, tab } of nonExcludedTabs) {
            if (this.selectedCompany === 'EOS') {
                void this.createEOSWorkbook(analysis, tab);
            } else {
                void this.createHIMarineWorkbook(analysis, tab);
            }
        }
    }

    private async createEOSWorkbook(analysis: FileAnalysis, tab: TabInfo): Promise<void> {
        try {
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
                console.error(`Worksheet "${tab.tabName}" not found`);
                return;
            }

            const cellRef = XLSX.utils.decode_cell(tab.topLeftCell);
            const headerRow = cellRef.r;
            const startCol = cellRef.c;
            const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1:Z100');

            const productCol = this.findColumnIndex(worksheet, headerRow, startCol, tab.product);
            const qtyCol = this.findColumnIndex(worksheet, headerRow, startCol, tab.qty);
            const unitCol = this.findColumnIndex(worksheet, headerRow, startCol, tab.unit);
            const remarkCol = this.findColumnIndex(worksheet, headerRow, startCol, tab.remark);

            const dataRows: any[] = [];
            for (let row = headerRow + 1; row <= range.e.r; row++) {
                const productCell = XLSX.utils.encode_cell({ r: row, c: productCol });
                const productValue = worksheet[productCell]?.v;
                if (productValue && String(productValue).trim() !== '') {
                    const qtyValue = worksheet[XLSX.utils.encode_cell({ r: row, c: qtyCol })]?.v || '';
                    const unitValue = worksheet[XLSX.utils.encode_cell({ r: row, c: unitCol })]?.v || '';
                    const remarkValue = worksheet[XLSX.utils.encode_cell({ r: row, c: remarkCol })]?.v || '';

                    dataRows.push({
                        description: String(productValue).trim(),
                        remark: String(remarkValue).trim(),
                        unit: String(unitValue).trim(),
                        qty: String(qtyValue).trim()
                    });
                } else {
                    break;
                }
            }

            const newWorkbook = new ExcelJS.Workbook();
            const newWorksheet = newWorkbook.addWorksheet('Sheet1');
            newWorksheet.properties.showGridLines = false;
            newWorksheet.views = [{ showGridLines: false }];

            const formatDateAsText = (dateString: string): string => {
                if (!dateString) return '';
                const date = new Date(dateString);
                const months = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
                const month = months[date.getMonth()];
                const day = date.getDate().toString().padStart(2, '0');
                const year = date.getFullYear();
                return `${month} ${day}, ${year}`;
            };

            newWorksheet.getCell('E2').value = this.rfqData.ourCompanyName || 'EOS SUPPLY LTD';
            if (this.rfqData.ourCompanyPhone) {
                newWorksheet.getCell('E4').value = `Phone: ${this.rfqData.ourCompanyPhone}`;
            }
            newWorksheet.getCell('E6').value = this.rfqData.ourCompanyEmail || 'office@eos-supply.co.uk';

            let companyRow = 8;
            if (this.rfqData.ourCompanyName) {
                newWorksheet.getCell(`A${companyRow}`).value = this.rfqData.ourCompanyName;
                companyRow++;
            }
            if (this.rfqData.ourCompanyEmail) {
                newWorksheet.getCell(`A${companyRow}`).value = this.rfqData.ourCompanyEmail;
                companyRow++;
            }
            if (this.rfqData.ourCompanyAddress) {
                newWorksheet.getCell(`A${companyRow}`).value = this.rfqData.ourCompanyAddress;
                companyRow++;
            }
            if (this.rfqData.ourCompanyAddress2) {
                newWorksheet.getCell(`A${companyRow}`).value = this.rfqData.ourCompanyAddress2;
                companyRow++;
            }
            const cityCountry = [this.rfqData.ourCompanyCity, this.rfqData.ourCompanyCountry].filter(Boolean).join(', ');
            if (cityCountry) {
                newWorksheet.getCell(`A${companyRow}`).value = cityCountry;
                companyRow++;
            }

            let bankRow = 15;
            if (this.rfqData.bankName) {
                newWorksheet.getCell(`A${bankRow}`).value = `Bank Name: ${this.rfqData.bankName}`;
                bankRow++;
            }
            if (this.rfqData.bankAddress) {
                newWorksheet.getCell(`A${bankRow}`).value = `Bank Address: ${this.rfqData.bankAddress}`;
                bankRow++;
            }
            if (this.rfqData.iban) {
                newWorksheet.getCell(`A${bankRow}`).value = `IBAN: ${this.rfqData.iban}`;
                bankRow++;
            }
            if (this.rfqData.swiftCode) {
                newWorksheet.getCell(`A${bankRow}`).value = `SWIFTBIC: ${this.rfqData.swiftCode}`;
                bankRow++;
            }
            if (this.rfqData.intermediaryBic) {
                newWorksheet.getCell(`A${bankRow}`).value = `Intermediary BIC: ${this.rfqData.intermediaryBic}`;
                bankRow++;
            }
            if (this.rfqData.accountTitle) {
                newWorksheet.getCell(`A${bankRow}`).value = `Title on Account: ${this.rfqData.accountTitle}`;
                bankRow++;
            }

            if (this.rfqData.accountNumber || this.rfqData.sortCode) {
                newWorksheet.getCell(`A${bankRow}`).value = 'UK DOMESTIC WIRES:';
                bankRow++;
                if (this.rfqData.accountNumber) {
                    newWorksheet.getCell(`A${bankRow}`).value = `Account number: ${this.rfqData.accountNumber}`;
                    bankRow++;
                }
                if (this.rfqData.sortCode) {
                    newWorksheet.getCell(`A${bankRow}`).value = `Sort code: ${this.rfqData.sortCode}`;
                    bankRow++;
                }
            }

            let invoiceRow = 15;
            newWorksheet.getCell(`E${invoiceRow}`).value = '№';
            invoiceRow++;
            newWorksheet.getCell(`E${invoiceRow}`).value = 'Invoice Date';
            if (this.rfqData.invoiceDate) {
                newWorksheet.getCell(`F${invoiceRow}`).value = formatDateAsText(this.rfqData.invoiceDate);
            }
            invoiceRow++;
            newWorksheet.getCell(`E${invoiceRow}`).value = 'Vessel';
            if (this.rfqData.vessel) {
                newWorksheet.getCell(`F${invoiceRow}`).value = this.rfqData.vessel;
            }
            invoiceRow++;
            newWorksheet.getCell(`E${invoiceRow}`).value = 'Country';
            if (this.rfqData.country) {
                newWorksheet.getCell(`F${invoiceRow}`).value = this.rfqData.country;
            }
            invoiceRow++;
            newWorksheet.getCell(`E${invoiceRow}`).value = 'Port';
            if (this.rfqData.port) {
                newWorksheet.getCell(`F${invoiceRow}`).value = this.rfqData.port;
            }
            invoiceRow++;
            newWorksheet.getCell(`E${invoiceRow}`).value = 'Category';
            if (this.rfqData.category) {
                newWorksheet.getCell(`F${invoiceRow}`).value = this.rfqData.category;
            }
            invoiceRow++;
            newWorksheet.getCell(`E${invoiceRow}`).value = 'Invoice Due';
            if (this.rfqData.invoiceDue) {
                newWorksheet.getCell(`F${invoiceRow}`).value = this.rfqData.invoiceDue;
            }

            newWorksheet.getCell('A27').value = 'Pos.';
            newWorksheet.getCell('B27').value = 'Description';
            newWorksheet.getCell('C27').value = 'Remark';
            newWorksheet.getCell('D27').value = 'Unit';
            newWorksheet.getCell('E27').value = 'Qty';
            newWorksheet.getCell('F27').value = 'Price';
            newWorksheet.getCell('G27').value = 'Total';

            dataRows.forEach((row, index) => {
                const rowNum = 28 + index;
                newWorksheet.getCell(`A${rowNum}`).value = index + 1;
                newWorksheet.getCell(`B${rowNum}`).value = row.description;
                newWorksheet.getCell(`C${rowNum}`).value = row.remark;
                newWorksheet.getCell(`D${rowNum}`).value = row.unit;
                newWorksheet.getCell(`E${rowNum}`).value = row.qty;
                newWorksheet.getCell(`F${rowNum}`).value = '';
                newWorksheet.getCell(`G${rowNum}`).value = '$';
            });

            const maxRow = 27 + dataRows.length;
            newWorksheet.pageSetup.printArea = `A1:G${maxRow}`;
            newWorksheet.views = [{
                showGridLines: false
            }];

            newWorksheet.getColumn('A').width = 30 / 7.59;
            newWorksheet.getColumn('B').width = 381 / 7.59;
            newWorksheet.getColumn('C').width = 174 / 7.59;
            newWorksheet.getColumn('D').width = 72 / 7.59;
            newWorksheet.getColumn('E').width = 138 / 7.59;
            newWorksheet.getColumn('F').width = 115 / 7.59;
            newWorksheet.getColumn('G').width = 119 / 7.59;

            const fileName = `${analysis.fileName}_${tab.tabName}_EOS_RFQ.xlsx`;
            const buffer = await newWorkbook.xlsx.writeBuffer();
            const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
            saveAs(blob, fileName);

            this.loggingService.logUserAction('eos_workbook_created', {
                fileName: fileName,
                tabName: tab.tabName,
                rowCount: dataRows.length
            }, 'RfqStateService');

        } catch (error) {
            console.error('Error creating EOS workbook:', error);
            this.loggingService.logError(
                error as Error,
                'eos_workbook_creation_error',
                'RfqStateService',
                { fileName: analysis.fileName, tabName: tab.tabName }
            );
        }
    }

    private async createHIMarineWorkbook(analysis: FileAnalysis, tab: TabInfo): Promise<void> {
        try {
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
                console.error(`Worksheet "${tab.tabName}" not found`);
                return;
            }

            const cellRef = XLSX.utils.decode_cell(tab.topLeftCell);
            const headerRow = cellRef.r;
            const startCol = cellRef.c;
            const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1:Z100');

            const productCol = this.findColumnIndex(worksheet, headerRow, startCol, tab.product);
            const qtyCol = this.findColumnIndex(worksheet, headerRow, startCol, tab.qty);
            const unitCol = this.findColumnIndex(worksheet, headerRow, startCol, tab.unit);
            const remarkCol = this.findColumnIndex(worksheet, headerRow, startCol, tab.remark);

            const dataRows: any[] = [];
            for (let row = headerRow + 1; row <= range.e.r; row++) {
                const productCell = XLSX.utils.encode_cell({ r: row, c: productCol });
                const productValue = worksheet[productCell]?.v;
                if (productValue && String(productValue).trim() !== '') {
                    const qtyValue = worksheet[XLSX.utils.encode_cell({ r: row, c: qtyCol })]?.v || '';
                    const unitValue = worksheet[XLSX.utils.encode_cell({ r: row, c: unitCol })]?.v || '';
                    const remarkValue = worksheet[XLSX.utils.encode_cell({ r: row, c: remarkCol })]?.v || '';

                    dataRows.push({
                        description: String(productValue).trim(),
                        remark: String(remarkValue).trim(),
                        unit: String(unitValue).trim(),
                        qty: String(qtyValue).trim()
                    });
                } else {
                    break;
                }
            }

            const newWorkbook = new ExcelJS.Workbook();
            const newWorksheet = newWorkbook.addWorksheet('Sheet1');
            newWorksheet.properties.showGridLines = false;
            newWorksheet.views = [{ showGridLines: false }];

            const formatDateAsText = (dateString: string): string => {
                if (!dateString) return '';
                const date = new Date(dateString);
                const months = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
                const month = months[date.getMonth()];
                const day = date.getDate().toString().padStart(2, '0');
                const year = date.getFullYear();
                return `${month} ${day}, ${year}`;
            };

            newWorksheet.getCell('A2').value = 'hi marine co';

            let companyRow = 9;
            if (this.rfqData.ourCompanyName) {
                newWorksheet.getCell(`A${companyRow}`).value = this.rfqData.ourCompanyName;
                companyRow++;
            }
            if (this.rfqData.ourCompanyAddress) {
                newWorksheet.getCell(`A${companyRow}`).value = this.rfqData.ourCompanyAddress;
                companyRow++;
            }
            if (this.rfqData.ourCompanyAddress2) {
                newWorksheet.getCell(`A${companyRow}`).value = this.rfqData.ourCompanyAddress2;
                companyRow++;
            }
            const cityCountry = [this.rfqData.ourCompanyCity, this.rfqData.ourCompanyCountry].filter(Boolean).join(', ');
            if (cityCountry) {
                newWorksheet.getCell(`A${companyRow}`).value = cityCountry;
                companyRow++;
            }
            if (this.rfqData.ourCompanyPhone) {
                newWorksheet.getCell(`A${companyRow}`).value = `Phone: ${this.rfqData.ourCompanyPhone}`;
                companyRow++;
            }
            if (this.rfqData.ourCompanyEmail) {
                newWorksheet.getCell(`A${companyRow}`).value = `Email: ${this.rfqData.ourCompanyEmail}`;
                companyRow++;
            }

            let vesselRow = 9;
            if (this.rfqData.vesselName) {
                newWorksheet.getCell(`E${vesselRow}`).value = this.rfqData.vesselName;
                vesselRow++;
            }
            if (this.rfqData.vesselName2) {
                newWorksheet.getCell(`E${vesselRow}`).value = this.rfqData.vesselName2;
                vesselRow++;
            }
            if (this.rfqData.vesselAddress) {
                newWorksheet.getCell(`E${vesselRow}`).value = this.rfqData.vesselAddress;
                vesselRow++;
            }
            if (this.rfqData.vesselAddress2) {
                newWorksheet.getCell(`E${vesselRow}`).value = this.rfqData.vesselAddress2;
                vesselRow++;
            }
            const vesselCityCountry = [this.rfqData.vesselCity, this.rfqData.vesselCountry].filter(Boolean).join(', ');
            if (vesselCityCountry) {
                newWorksheet.getCell(`E${vesselRow}`).value = vesselCityCountry;
                vesselRow++;
            }

            let bankRow = 16;
            if (this.rfqData.bankName) {
                newWorksheet.getCell(`A${bankRow}`).value = 'Bank Name:';
                newWorksheet.getCell(`B${bankRow}`).value = this.rfqData.bankName;
                bankRow++;
            }
            if (this.rfqData.bankAddress) {
                newWorksheet.getCell(`A${bankRow}`).value = 'Bank Address:';
                newWorksheet.getCell(`B${bankRow}`).value = this.rfqData.bankAddress;
                bankRow++;
            }
            if (this.rfqData.accountNumber) {
                newWorksheet.getCell(`A${bankRow}`).value = 'Account No:';
                newWorksheet.getCell(`B${bankRow}`).value = this.rfqData.accountNumber;
                bankRow++;
            }
            if (this.rfqData.swiftCode) {
                newWorksheet.getCell(`A${bankRow}`).value = 'SWIFT CODE:';
                newWorksheet.getCell(`B${bankRow}`).value = this.rfqData.swiftCode;
                bankRow++;
            }
            if (this.selectedCompany === 'HI US' && this.rfqData.achRouting) {
                newWorksheet.getCell(`A${bankRow}`).value = 'ACH Routing:';
                newWorksheet.getCell(`B${bankRow}`).value = this.rfqData.achRouting;
                bankRow++;
            } else if (this.selectedCompany === 'HI UK' && this.rfqData.sortCode) {
                newWorksheet.getCell(`A${bankRow}`).value = 'Sort Code:';
                newWorksheet.getCell(`B${bankRow}`).value = this.rfqData.sortCode;
                bankRow++;
            }
            if (this.rfqData.accountTitle) {
                newWorksheet.getCell(`A${bankRow}`).value = 'Title on Account:';
                newWorksheet.getCell(`B${bankRow}`).value = this.rfqData.accountTitle;
                bankRow++;
            }

            let invoiceRow = 15;
            newWorksheet.getCell(`D${invoiceRow}`).value = '№';
            if (this.rfqData.invoiceNumber) {
                newWorksheet.getCell(`F${invoiceRow}`).value = this.rfqData.invoiceNumber;
            }
            invoiceRow++;
            newWorksheet.getCell(`D${invoiceRow}`).value = 'Invoice Date';
            if (this.rfqData.invoiceDate) {
                newWorksheet.getCell(`F${invoiceRow}`).value = formatDateAsText(this.rfqData.invoiceDate);
            }
            invoiceRow++;
            newWorksheet.getCell(`D${invoiceRow}`).value = 'Vessel';
            if (this.rfqData.vessel) {
                newWorksheet.getCell(`F${invoiceRow}`).value = this.rfqData.vessel;
            }
            invoiceRow++;
            newWorksheet.getCell(`D${invoiceRow}`).value = 'Country';
            if (this.rfqData.country) {
                newWorksheet.getCell(`F${invoiceRow}`).value = this.rfqData.country;
            }
            invoiceRow++;
            newWorksheet.getCell(`D${invoiceRow}`).value = 'Port';
            if (this.rfqData.port) {
                newWorksheet.getCell(`F${invoiceRow}`).value = this.rfqData.port;
            }
            invoiceRow++;
            newWorksheet.getCell(`D${invoiceRow}`).value = 'Category';
            if (this.rfqData.category) {
                newWorksheet.getCell(`F${invoiceRow}`).value = this.rfqData.category;
            }
            invoiceRow++;
            newWorksheet.getCell(`D${invoiceRow}`).value = 'Invoice Due';
            if (this.rfqData.invoiceDue) {
                newWorksheet.getCell(`F${invoiceRow}`).value = this.rfqData.invoiceDue;
            }

            newWorksheet.getCell('A26').value = 'Pos.';
            newWorksheet.getCell('B26').value = 'Description';
            newWorksheet.getCell('C26').value = 'Remark';
            newWorksheet.getCell('D26').value = 'Unit';
            newWorksheet.getCell('E26').value = 'Qty';
            newWorksheet.getCell('F26').value = 'Price';
            newWorksheet.getCell('G26').value = 'Total';

            dataRows.forEach((row, index) => {
                const rowNum = 27 + index;
                newWorksheet.getCell(`A${rowNum}`).value = index + 1;
                newWorksheet.getCell(`B${rowNum}`).value = row.description;
                newWorksheet.getCell(`C${rowNum}`).value = row.remark;
                newWorksheet.getCell(`D${rowNum}`).value = row.unit;
                newWorksheet.getCell(`E${rowNum}`).value = row.qty;
                newWorksheet.getCell(`F${rowNum}`).value = '';
                newWorksheet.getCell(`G${rowNum}`).value = '$ -';
            });

            const maxRow = 26 + dataRows.length;
            newWorksheet.pageSetup.printArea = `A1:G${maxRow}`;
            newWorksheet.views = [{
                showGridLines: false
            }];

            newWorksheet.getColumn('A').width = 56 / 7.59;
            newWorksheet.getColumn('B').width = 295 / 7.59;
            newWorksheet.getColumn('C').width = 56 / 7.59;
            newWorksheet.getColumn('D').width = 80 / 7.59;
            newWorksheet.getColumn('E').width = 82 / 7.59;
            newWorksheet.getColumn('F').width = 131 / 7.59;
            newWorksheet.getColumn('G').width = 121 / 7.59;

            const fileName = `${analysis.fileName}_${tab.tabName}_${this.selectedCompany}_RFQ.xlsx`;
            const buffer = await newWorkbook.xlsx.writeBuffer();
            const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
            saveAs(blob, fileName);

            this.loggingService.logUserAction('himarine_workbook_created', {
                fileName: fileName,
                tabName: tab.tabName,
                company: this.selectedCompany,
                rowCount: dataRows.length
            }, 'RfqStateService');

        } catch (error) {
            console.error('Error creating HI Marine workbook:', error);
            this.loggingService.logError(
                error as Error,
                'himarine_workbook_creation_error',
                'RfqStateService',
                { fileName: analysis.fileName, tabName: tab.tabName, company: this.selectedCompany }
            );
        }
    }

    private findColumnIndex(worksheet: XLSX.WorkSheet, headerRow: number, startCol: number, headerName: string): number {
        const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1:Z100');
        for (let col = startCol; col <= range.e.c; col++) {
            const headerAddress = XLSX.utils.encode_cell({ r: headerRow, c: col });
            const headerCell = worksheet[headerAddress];
            if (headerCell && headerCell.v) {
                const cellValue = String(headerCell.v).trim();
                if (cellValue === headerName) {
                    return col;
                }
            }
        }
        return startCol;
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

    onExcludeChange(tab: TabInfo): void {
        this.loggingService.logUserAction('exclude_toggled', {
            tabName: tab.tabName,
            excluded: tab.excluded
        }, 'RfqStateService');
    }

    private updateExcludedState(tab: TabInfo): void {
        tab.excluded = tab.rowCount === 0 || !tab.topLeftCell || tab.topLeftCell.trim() === '';
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


