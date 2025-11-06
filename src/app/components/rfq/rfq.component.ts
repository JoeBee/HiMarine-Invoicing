import { Component, ChangeDetectorRef } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { LoggingService } from '../../services/logging.service';
import * as XLSX from 'xlsx';
import * as ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';

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
    excluded: boolean;
}

interface FileAnalysis {
    fileName: string;
    numberOfTabs: number;
    tabs: TabInfo[];
    file: File;
}

interface RfqData {
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
    achRouting?: string; // For US bank details
    intermediaryBic?: string; // For EOS bank details
    // Invoice/Quotation Details
    invoiceNumber: string;
    invoiceDate: string;
    vessel: string;
    country: string;
    port: string;
    category: string;
    invoiceDue: string;
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

    // RFQ Data structure
    rfqData: RfqData = {
        // Our Company Details
        ourCompanyName: '',
        ourCompanyAddress: '',
        ourCompanyAddress2: '',
        ourCompanyCity: '',
        ourCompanyCountry: '',
        ourCompanyPhone: '',
        ourCompanyEmail: '',
        // Vessel Details
        vesselName: '',
        vesselName2: '',
        vesselAddress: '',
        vesselAddress2: '',
        vesselCity: '',
        vesselCountry: '',
        // Bank Details
        bankName: '',
        bankAddress: '',
        iban: '',
        swiftCode: '',
        accountTitle: '',
        accountNumber: '',
        sortCode: '',
        achRouting: '',
        intermediaryBic: '',
        // Invoice/Quotation Details
        invoiceNumber: '',
        invoiceDate: this.getTodayDate(),
        vessel: '',
        country: '',
        port: '',
        category: 'Provisions',
        invoiceDue: ''
    };

    // Country dropdown options (same as invoice component)
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

    // Available ports for selected country
    availablePorts: string[] = [];

    // Country to ports mapping (same as invoice component)
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

    constructor(
        private loggingService: LoggingService,
        private cdr: ChangeDetectorRef
    ) {
        // Initialize with default company data
        this.onCompanySelectionChange();
    }

    private getTodayDate(): string {
        const today = new Date();
        return today.toISOString().split('T')[0]; // Returns YYYY-MM-DD format
    }

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
                        sheetStubs: false,
                        // Options to handle hidden columns and protected files
                        cellText: true,
                        cellDates: true
                    });

                    // Validate workbook structure
                    if (!workbook) {
                        throw new Error('Failed to read workbook - workbook is null or undefined');
                    }

                    // Log workbook structure for debugging
                    if (!workbook.Sheets && !workbook.SheetNames) {
                        this.loggingService.logError(
                            new Error('Workbook has no Sheets or SheetNames'),
                            'workbook_structure_invalid',
                            'RfqComponent',
                            {
                                fileName: file.name,
                                workbookKeys: workbook ? Object.keys(workbook) : [],
                                hasWorkbook: !!workbook
                            }
                        );
                        throw new Error('Invalid workbook structure - workbook, Sheets, or SheetNames missing');
                    }

                    // Log if Sheets is empty but SheetNames exists
                    if ((!workbook.Sheets || Object.keys(workbook.Sheets).length === 0) &&
                        workbook.SheetNames && workbook.SheetNames.length > 0) {
                        console.warn('Workbook has SheetNames but Sheets object is empty', {
                            fileName: file.name,
                            sheetNames: workbook.SheetNames,
                            workbookKeys: Object.keys(workbook)
                        });
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
                                        topLeftCell: '',
                                        product: '',
                                        qty: '',
                                        unit: '',
                                        remark: '',
                                        isHidden: true,
                                        columnHeaders: [],
                                        excluded: true // Auto-exclude hidden sheets with no data
                                    });
                                    return;
                                }

                                const datatableInfo = this.findDatatableInfo(worksheet);
                                const autoSelected = this.autoSelectColumns(datatableInfo.columnHeaders);
                                // Auto-exclude if rowCount is 0 or topLeftCell is blank
                                const shouldExclude = datatableInfo.rowCount === 0 || !datatableInfo.topLeftCell || datatableInfo.topLeftCell.trim() === '';
                                tabInfos.push({
                                    tabName: sheetName,
                                    rowCount: datatableInfo.rowCount,
                                    topLeftCell: datatableInfo.topLeftCell,
                                    product: autoSelected.product || datatableInfo.product,
                                    qty: autoSelected.qty || datatableInfo.qty,
                                    unit: autoSelected.unit || datatableInfo.unit,
                                    remark: autoSelected.remark || datatableInfo.remark,
                                    isHidden: isHidden,
                                    columnHeaders: datatableInfo.columnHeaders,
                                    excluded: shouldExclude
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
                                    topLeftCell: '',
                                    product: '',
                                    qty: '',
                                    unit: '',
                                    remark: '',
                                    isHidden: true,
                                    columnHeaders: [],
                                    excluded: true // Auto-exclude hidden sheets with no data
                                });
                            }
                            return;
                        }

                        const datatableInfo = this.findDatatableInfo(worksheet);
                        const autoSelected = this.autoSelectColumns(datatableInfo.columnHeaders);
                        // Auto-exclude if rowCount is 0 or topLeftCell is blank
                        const shouldExclude = datatableInfo.rowCount === 0 || !datatableInfo.topLeftCell || datatableInfo.topLeftCell.trim() === '';
                        tabInfos.push({
                            tabName: sheetName,
                            rowCount: datatableInfo.rowCount,
                            topLeftCell: datatableInfo.topLeftCell,
                            product: autoSelected.product || datatableInfo.product,
                            qty: autoSelected.qty || datatableInfo.qty,
                            unit: autoSelected.unit || datatableInfo.unit,
                            remark: autoSelected.remark || datatableInfo.remark,
                            isHidden: isHidden,
                            columnHeaders: datatableInfo.columnHeaders,
                            excluded: shouldExclude
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
        let topLeftCell = '';

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
                topLeftCell: '',
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

    private async reanalyzeTabWithTopLeft(analysis: FileAnalysis, tab: TabInfo, topLeftCell: string): Promise<void> {
        try {
            // Read the Excel file
            const file = analysis.file;
            const fileData = await this.readFileAsArrayBuffer(file);
            const workbook = XLSX.read(fileData, {
                type: 'array',
                cellFormula: false,
                cellHTML: false,
                cellStyles: false,
                sheetStubs: false,
                // Options to handle hidden columns and protected files
                cellText: true,
                cellDates: true
            });

            // Find the worksheet by tab name - try exact match first, then case-insensitive
            let worksheet = workbook.Sheets[tab.tabName];
            if (!worksheet) {
                // Try case-insensitive match
                const sheetNames = Object.keys(workbook.Sheets);
                const matchingSheet = sheetNames.find(name => name.toLowerCase() === tab.tabName.toLowerCase());
                if (matchingSheet) {
                    worksheet = workbook.Sheets[matchingSheet];
                }
            }

            if (!worksheet) {
                this.loggingService.logError(
                    new Error(`Worksheet "${tab.tabName}" not found`),
                    'worksheet_not_found',
                    'RfqComponent',
                    {
                        tabName: tab.tabName,
                        fileName: analysis.fileName,
                        availableSheets: Object.keys(workbook.Sheets)
                    }
                );
                return;
            }

            // Parse the topLeftCell (e.g., "A19" -> row 18, col 0)
            // XLSX uses 0-based indexing, so A19 = row 18 (0-indexed)
            let cellRef;
            try {
                cellRef = XLSX.utils.decode_cell(topLeftCell);
            } catch (error) {
                this.loggingService.logError(
                    new Error(`Invalid cell reference: ${topLeftCell}`),
                    'invalid_cell_reference',
                    'RfqComponent',
                    { topLeftCell, tabName: tab.tabName, fileName: analysis.fileName }
                );
                return;
            }

            const headerRow = cellRef.r;
            const startCol = cellRef.c;

            // Get the range of the worksheet
            const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1:Z100');

            // Extract all column headers from the header row starting from topLeftCell
            const columnHeaders: string[] = [];
            // Extend range to check more columns (up to column Z or beyond if needed)
            const maxCol = Math.max(range.e.c, 25); // At least check up to column Z (25)

            // Also check if we need to go beyond the detected range
            // Try reading up to column N (13) or more to handle cases where range detection is limited
            const extendedMaxCol = Math.max(maxCol, 13); // At least up to column N

            // Read all columns, including potentially hidden ones
            // Don't stop at first empty cell - continue until we find a sequence of empty cells
            let emptyCellCount = 0;
            const maxEmptyCells = 2; // Stop after 2 consecutive empty cells

            for (let col = startCol; col <= extendedMaxCol; col++) {
                const headerAddress = XLSX.utils.encode_cell({ r: headerRow, c: col });
                const headerCell = worksheet[headerAddress];

                // Check both .v (value) and .w (formatted text) properties
                let cellValue: any = null;
                if (headerCell) {
                    cellValue = headerCell.v !== null && headerCell.v !== undefined ? headerCell.v :
                        (headerCell.w !== null && headerCell.w !== undefined ? headerCell.w : null);
                }

                if (cellValue !== null && cellValue !== undefined) {
                    const headerValue = String(cellValue).trim();
                    if (headerValue) {
                        columnHeaders.push(headerValue);
                        emptyCellCount = 0; // Reset empty cell counter
                    } else {
                        emptyCellCount++;
                        if (emptyCellCount >= maxEmptyCells) {
                            break;
                        }
                    }
                } else {
                    emptyCellCount++;
                    if (emptyCellCount >= maxEmptyCells) {
                        break;
                    }
                }
            }

            // Log for debugging
            this.loggingService.logUserAction('headers_extracted', {
                fileName: analysis.fileName,
                tabName: tab.tabName,
                topLeftCell: topLeftCell,
                headerRow: headerRow,
                startCol: startCol,
                columnCount: columnHeaders.length,
                headers: columnHeaders,
                worksheetRange: worksheet['!ref'],
                extendedMaxCol: extendedMaxCol
            }, 'RfqComponent');

            // If no headers found, try a different approach - read cells directly
            if (columnHeaders.length === 0) {
                console.warn('No headers found with standard approach, trying alternative method');
                // Try reading cells directly without relying on range
                for (let col = 0; col <= 20; col++) { // Try up to column U
                    const headerAddress = XLSX.utils.encode_cell({ r: headerRow, c: col });
                    const headerCell = worksheet[headerAddress];
                    if (headerCell) {
                        const cellValue = headerCell.v !== null && headerCell.v !== undefined ? headerCell.v :
                            (headerCell.w !== null && headerCell.w !== undefined ? headerCell.w : null);
                        if (cellValue !== null && cellValue !== undefined) {
                            const headerValue = String(cellValue).trim();
                            if (headerValue) {
                                columnHeaders.push(headerValue);
                            }
                        }
                    }
                }

                // If we found headers with alternative method, log it
                if (columnHeaders.length > 0) {
                    this.loggingService.logUserAction('headers_extracted_alternative', {
                        fileName: analysis.fileName,
                        tabName: tab.tabName,
                        topLeftCell: topLeftCell,
                        columnCount: columnHeaders.length,
                        headers: columnHeaders
                    }, 'RfqComponent');
                }
            }

            // Update the tab's column headers
            tab.columnHeaders = columnHeaders;

            // Re-run auto-select to update Product, Qty, Unit, Remark
            const autoSelected = this.autoSelectColumns(columnHeaders);
            if (autoSelected.product) {
                tab.product = autoSelected.product;
            }
            if (autoSelected.qty) {
                tab.qty = autoSelected.qty;
            }
            if (autoSelected.unit) {
                tab.unit = autoSelected.unit;
            }
            if (autoSelected.remark) {
                tab.remark = autoSelected.remark;
            }

            // Recalculate row count based on the new header row
            // Find description column in the new headers
            let descriptionColumn = -1;
            for (let i = 0; i < columnHeaders.length; i++) {
                const header = columnHeaders[i].toLowerCase().trim();
                if (header.includes('description') ||
                    header.includes('descrption') ||
                    header.includes('product') ||
                    header.includes('item') ||
                    header.includes('name')) {
                    descriptionColumn = startCol + i;
                    break;
                }
            }

            // Count data rows if we found a description column
            if (descriptionColumn !== -1) {
                let rowCount = 0;
                for (let dataRow = headerRow + 1; dataRow <= range.e.r; dataRow++) {
                    const descAddress = XLSX.utils.encode_cell({ r: dataRow, c: descriptionColumn });
                    const descCell = worksheet[descAddress];
                    const hasDescription = descCell && descCell.v !== null && descCell.v !== undefined && String(descCell.v).trim() !== '';
                    if (hasDescription) {
                        rowCount++;
                    }
                }
                tab.rowCount = rowCount;
            } else {
                // If no description column found, count all non-empty rows after header
                let rowCount = 0;
                for (let dataRow = headerRow + 1; dataRow <= range.e.r; dataRow++) {
                    let hasData = false;
                    for (let col = startCol; col <= startCol + columnHeaders.length - 1 && col <= range.e.c; col++) {
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
                        // Stop at first empty row
                        break;
                    }
                }
                tab.rowCount = rowCount;
            }

            this.loggingService.logUserAction('tab_reanalyzed', {
                fileName: analysis.fileName,
                tabName: tab.tabName,
                topLeftCell: topLeftCell,
                columnCount: columnHeaders.length,
                rowCount: tab.rowCount
            }, 'RfqComponent');

            // Force change detection to update the UI
            this.cdr.detectChanges();

        } catch (error) {
            this.loggingService.logError(
                error as Error,
                'tab_reanalysis_error',
                'RfqComponent',
                {
                    fileName: analysis.fileName,
                    tabName: tab.tabName,
                    topLeftCell: topLeftCell
                }
            );
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

    canCreateRFQs(): boolean {
        // Check if there is no data in the datatable
        if (this.fileAnalyses.length === 0) {
            return false;
        }

        // Get all non-excluded tabs
        const nonExcludedTabs: { analysis: FileAnalysis; tab: TabInfo }[] = [];
        for (const analysis of this.fileAnalyses) {
            for (const tab of analysis.tabs) {
                if (!tab.excluded) {
                    nonExcludedTabs.push({ analysis, tab });
                }
            }
        }

        // Button is disabled if all records have EXCLUDE checked
        if (nonExcludedTabs.length === 0) {
            return false;
        }

        // Button is disabled if any non-excluded row has empty Product, Qty, Unit, or Remark
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
        }, 'RfqComponent');

        // Get all non-excluded tabs
        const nonExcludedTabs: { analysis: FileAnalysis; tab: TabInfo }[] = [];
        for (const analysis of this.fileAnalyses) {
            for (const tab of analysis.tabs) {
                if (!tab.excluded) {
                    nonExcludedTabs.push({ analysis, tab });
                }
            }
        }

        // Create Excel workbooks for each non-excluded tab
        for (const { analysis, tab } of nonExcludedTabs) {
            if (this.selectedCompany === 'EOS') {
                this.createEOSWorkbook(analysis, tab);
            } else {
                this.createHIMarineWorkbook(analysis, tab);
            }
        }
    }

    private async createEOSWorkbook(analysis: FileAnalysis, tab: TabInfo): Promise<void> {
        try {
            // Read the original Excel file to get the data
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

            // Parse the topLeftCell to get header row
            const cellRef = XLSX.utils.decode_cell(tab.topLeftCell);
            const headerRow = cellRef.r;
            const startCol = cellRef.c;
            const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1:Z100');

            // Find column indices for Product, Qty, Unit, Remark
            const productCol = this.findColumnIndex(worksheet, headerRow, startCol, tab.product);
            const qtyCol = this.findColumnIndex(worksheet, headerRow, startCol, tab.qty);
            const unitCol = this.findColumnIndex(worksheet, headerRow, startCol, tab.unit);
            const remarkCol = this.findColumnIndex(worksheet, headerRow, startCol, tab.remark);

            // Extract data rows
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
                    break; // Stop at first empty row
                }
            }

            // Create new workbook with EOS template using ExcelJS
            const newWorkbook = new ExcelJS.Workbook();
            const newWorksheet = newWorkbook.addWorksheet('Sheet1');

            // Remove grid lines for cleaner look (same as invoice component)
            newWorksheet.properties.showGridLines = false;
            newWorksheet.views = [{ showGridLines: false }];

            // Helper to format date as "November 03, 2025"
            const formatDateAsText = (dateString: string): string => {
                if (!dateString) return '';
                const date = new Date(dateString);
                const months = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
                const month = months[date.getMonth()];
                const day = date.getDate().toString().padStart(2, '0');
                const year = date.getFullYear();
                return `${month} ${day}, ${year}`;
            };

            // Set up EOS template structure based on image 1
            // Header section (right side)
            newWorksheet.getCell('E2').value = this.rfqData.ourCompanyName || 'EOS SUPPLY LTD';
            if (this.rfqData.ourCompanyPhone) {
                newWorksheet.getCell('E4').value = `Phone: ${this.rfqData.ourCompanyPhone}`;
            }
            newWorksheet.getCell('E6').value = this.rfqData.ourCompanyEmail || 'office@eos-supply.co.uk';

            // Sender's address (left side)
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

            // Banking details
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

            // UK Domestic Wires (for EOS)
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

            // Invoice/Quotation specifics (right side)
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

            // Table headers (Row 27)
            newWorksheet.getCell('A27').value = 'Pos.';
            newWorksheet.getCell('B27').value = 'Description';
            newWorksheet.getCell('C27').value = 'Remark';
            newWorksheet.getCell('D27').value = 'Unit';
            newWorksheet.getCell('E27').value = 'Qty';
            newWorksheet.getCell('F27').value = 'Price';
            newWorksheet.getCell('G27').value = 'Total';

            // Add data rows
            dataRows.forEach((row, index) => {
                const rowNum = 28 + index;
                newWorksheet.getCell(`A${rowNum}`).value = index + 1;
                newWorksheet.getCell(`B${rowNum}`).value = row.description;
                newWorksheet.getCell(`C${rowNum}`).value = row.remark;
                newWorksheet.getCell(`D${rowNum}`).value = row.unit;
                newWorksheet.getCell(`E${rowNum}`).value = row.qty;
                newWorksheet.getCell(`F${rowNum}`).value = ''; // Price empty
                newWorksheet.getCell(`G${rowNum}`).value = '$'; // Total with $ sign
            });

            // Set print area (Print Active Sheets)
            const maxRow = 27 + dataRows.length;
            newWorksheet.pageSetup.printArea = `A1:G${maxRow}`;

            // Set view to page break preview (ExcelJS may not fully support this, but grid lines are removed)
            newWorksheet.views = [{
                showGridLines: false
            }];

            // Set column widths for EOS (ExcelJS uses character width)
            // Excel default: 8.43 chars = 64px, so 1 char ≈ 7.59px
            // Using more accurate conversion: pixels / 7.59
            newWorksheet.getColumn('A').width = 30 / 7.59;   // A: 30px
            newWorksheet.getColumn('B').width = 381 / 7.59;  // B: 381px
            newWorksheet.getColumn('C').width = 174 / 7.59;  // C: 174px
            newWorksheet.getColumn('D').width = 72 / 7.59;   // D: 72px
            newWorksheet.getColumn('E').width = 138 / 7.59;  // E: 138px
            newWorksheet.getColumn('F').width = 115 / 7.59;  // F: 115px
            newWorksheet.getColumn('G').width = 119 / 7.59;  // G: 119px

            // Generate filename and save
            const fileName = `${analysis.fileName}_${tab.tabName}_EOS_RFQ.xlsx`;
            const buffer = await newWorkbook.xlsx.writeBuffer();
            const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
            saveAs(blob, fileName);

            this.loggingService.logUserAction('eos_workbook_created', {
                fileName: fileName,
                tabName: tab.tabName,
                rowCount: dataRows.length
            }, 'RfqComponent');

        } catch (error) {
            console.error('Error creating EOS workbook:', error);
            this.loggingService.logError(
                error as Error,
                'eos_workbook_creation_error',
                'RfqComponent',
                { fileName: analysis.fileName, tabName: tab.tabName }
            );
        }
    }

    private async createHIMarineWorkbook(analysis: FileAnalysis, tab: TabInfo): Promise<void> {
        try {
            // Read the original Excel file to get the data
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

            // Parse the topLeftCell to get header row
            const cellRef = XLSX.utils.decode_cell(tab.topLeftCell);
            const headerRow = cellRef.r;
            const startCol = cellRef.c;
            const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1:Z100');

            // Find column indices for Product, Qty, Unit, Remark
            const productCol = this.findColumnIndex(worksheet, headerRow, startCol, tab.product);
            const qtyCol = this.findColumnIndex(worksheet, headerRow, startCol, tab.qty);
            const unitCol = this.findColumnIndex(worksheet, headerRow, startCol, tab.unit);
            const remarkCol = this.findColumnIndex(worksheet, headerRow, startCol, tab.remark);

            // Extract data rows
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
                    break; // Stop at first empty row
                }
            }

            // Create new workbook with HI Marine template using ExcelJS
            const newWorkbook = new ExcelJS.Workbook();
            const newWorksheet = newWorkbook.addWorksheet('Sheet1');

            // Remove grid lines for cleaner look (same as invoice component)
            newWorksheet.properties.showGridLines = false;
            newWorksheet.views = [{ showGridLines: false }];

            // Helper to format date as "November 03, 2025"
            const formatDateAsText = (dateString: string): string => {
                if (!dateString) return '';
                const date = new Date(dateString);
                const months = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
                const month = months[date.getMonth()];
                const day = date.getDate().toString().padStart(2, '0');
                const year = date.getFullYear();
                return `${month} ${day}, ${year}`;
            };

            // Set up HI Marine template structure based on image 2
            // Logo/Header (A2-B5)
            newWorksheet.getCell('A2').value = 'hi marine co';

            // Company Information (left side)
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

            // Vessel Details (right side, column E)
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

            // Bank Details (A16-B21)
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

            // Order/Invoice Details (Right side, D15-G21)
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

            // Table headers (Row 26)
            newWorksheet.getCell('A26').value = 'Pos.';
            newWorksheet.getCell('B26').value = 'Description';
            newWorksheet.getCell('C26').value = 'Remark';
            newWorksheet.getCell('D26').value = 'Unit';
            newWorksheet.getCell('E26').value = 'Qty';
            newWorksheet.getCell('F26').value = 'Price';
            newWorksheet.getCell('G26').value = 'Total';

            // Add data rows
            dataRows.forEach((row, index) => {
                const rowNum = 27 + index;
                newWorksheet.getCell(`A${rowNum}`).value = index + 1;
                newWorksheet.getCell(`B${rowNum}`).value = row.description;
                newWorksheet.getCell(`C${rowNum}`).value = row.remark;
                newWorksheet.getCell(`D${rowNum}`).value = row.unit;
                newWorksheet.getCell(`E${rowNum}`).value = row.qty;
                newWorksheet.getCell(`F${rowNum}`).value = ''; // Price empty
                newWorksheet.getCell(`G${rowNum}`).value = '$ -'; // Total with $ - sign
            });

            // Set print area (Print Active Sheets)
            const maxRow = 26 + dataRows.length;
            newWorksheet.pageSetup.printArea = `A1:G${maxRow}`;

            // Set view to page break preview (ExcelJS may not fully support this, but grid lines are removed)
            newWorksheet.views = [{
                showGridLines: false
            }];

            // Set column widths for HI US/HI UK (ExcelJS uses character width)
            // Excel default: 8.43 chars = 64px, so 1 char ≈ 7.59px
            // Using more accurate conversion: pixels / 7.59
            newWorksheet.getColumn('A').width = 56 / 7.59;   // A: 56px
            newWorksheet.getColumn('B').width = 295 / 7.59;  // B: 295px
            newWorksheet.getColumn('C').width = 56 / 7.59;   // C: 56px
            newWorksheet.getColumn('D').width = 80 / 7.59;   // D: 80px
            newWorksheet.getColumn('E').width = 82 / 7.59;   // E: 82px
            newWorksheet.getColumn('F').width = 131 / 7.59;  // F: 131px
            newWorksheet.getColumn('G').width = 121 / 7.59;  // G: 121px

            // Generate filename and save
            const fileName = `${analysis.fileName}_${tab.tabName}_${this.selectedCompany}_RFQ.xlsx`;
            const buffer = await newWorkbook.xlsx.writeBuffer();
            const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
            saveAs(blob, fileName);

            this.loggingService.logUserAction('himarine_workbook_created', {
                fileName: fileName,
                tabName: tab.tabName,
                company: this.selectedCompany,
                rowCount: dataRows.length
            }, 'RfqComponent');

        } catch (error) {
            console.error('Error creating HI Marine workbook:', error);
            this.loggingService.logError(
                error as Error,
                'himarine_workbook_creation_error',
                'RfqComponent',
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
        return startCol; // Fallback to start column
    }

    private setCellValue(worksheet: XLSX.WorkSheet, cellAddress: string, value: any): void {
        worksheet[cellAddress] = { v: value, t: typeof value === 'number' ? 'n' : 's' };
    }

    onCompanyChange(company: 'HI US' | 'HI UK' | 'EOS'): void {
        this.loggingService.logButtonClick(`company_selected_${company}`, 'RfqComponent', {
            selectedCompany: company
        });
        this.selectedCompany = company;
        this.onCompanySelectionChange();
    }

    onCompanySelectionChange(): void {
        // Clear all bank details first
        this.clearBankDetails();

        // Populate bank details based on selection
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
                // Keep fields blank
                break;
        }
    }

    private clearBankDetails(): void {
        // Clear Our Company Details
        this.rfqData.ourCompanyName = '';
        this.rfqData.ourCompanyAddress = '';
        this.rfqData.ourCompanyAddress2 = '';
        this.rfqData.ourCompanyCity = '';
        this.rfqData.ourCompanyCountry = '';
        this.rfqData.ourCompanyPhone = '';
        this.rfqData.ourCompanyEmail = '';

        // Clear Vessel Details
        this.rfqData.vesselName = '';
        this.rfqData.vesselName2 = '';
        this.rfqData.vesselAddress = '';
        this.rfqData.vesselAddress2 = '';
        this.rfqData.vesselCity = '';
        this.rfqData.vesselCountry = '';

        // Clear Bank Details
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
        // Our Company Details
        this.rfqData.ourCompanyName = 'HI MARINE COMPANY LIMITED';
        this.rfqData.ourCompanyAddress = '167-169 Great Portland Street';
        this.rfqData.ourCompanyAddress2 = '';
        this.rfqData.ourCompanyCity = 'London, London, W1W 5PF';
        this.rfqData.ourCompanyCountry = 'United Kingdom';
        this.rfqData.ourCompanyPhone = '';
        this.rfqData.ourCompanyEmail = 'office@himarinecompany.com';

        // Vessel Details (do not auto-populate; keep blank by default)
        this.rfqData.vesselName = '';
        this.rfqData.vesselName2 = '';
        this.rfqData.vesselAddress = '';
        this.rfqData.vesselAddress2 = '';
        this.rfqData.vesselCity = '';
        this.rfqData.vesselCountry = '';

        // Bank Details
        this.rfqData.bankName = 'Lloyds Bank plc';
        this.rfqData.bankAddress = '6 Market Place, Oldham, OL11JG, United Kingdom';
        this.rfqData.iban = 'GB84LOYD30962678553260';
        this.rfqData.swiftCode = 'LOYDGB21446';
        this.rfqData.accountTitle = 'HI MARINE COMPANY LIMITED';
        this.rfqData.accountNumber = '78553260';
        this.rfqData.sortCode = '30-96-26';
    }

    private populateUSBankDetails(): void {
        // Our Company Details
        this.rfqData.ourCompanyName = 'HI MARINE COMPANY INC.';
        this.rfqData.ourCompanyAddress = '9407 N.E. Vancouver Mall Drive, Suite 104';
        this.rfqData.ourCompanyAddress2 = '';
        this.rfqData.ourCompanyCity = 'Vancouver, WA  98662';
        this.rfqData.ourCompanyCountry = 'USA';
        this.rfqData.ourCompanyPhone = '+1 857 2045786';
        this.rfqData.ourCompanyEmail = 'office@himarinecompany.com';

        // Vessel Details (do not auto-populate; keep blank by default)
        this.rfqData.vesselName = '';
        this.rfqData.vesselName2 = '';
        this.rfqData.vesselAddress = '';
        this.rfqData.vesselAddress2 = '';
        this.rfqData.vesselCity = '';
        this.rfqData.vesselCountry = '';

        // Bank Details
        this.rfqData.bankName = 'Bank of America';
        this.rfqData.bankAddress = '100 West 33d Street New York, New York 10001';
        this.rfqData.accountNumber = '466002755612';
        this.rfqData.swiftCode = 'BofAUS3N';
        this.rfqData.achRouting = '011000138';
        this.rfqData.accountTitle = 'Hi Marine Company Inc.';
    }

    private populateEOSBankDetails(): void {
        // Our Company Details
        this.rfqData.ourCompanyName = 'EOS SUPPLY LTD';
        this.rfqData.ourCompanyAddress = '85 Great Portland Street, First Floor';
        this.rfqData.ourCompanyAddress2 = '';
        this.rfqData.ourCompanyCity = 'London, England, W1W 7LT';
        this.rfqData.ourCompanyCountry = 'United Kingdom';
        this.rfqData.ourCompanyPhone = '';
        this.rfqData.ourCompanyEmail = '';

        // Vessel Details (do not auto-populate; keep blank by default)
        this.rfqData.vesselName = '';
        this.rfqData.vesselName2 = '';
        this.rfqData.vesselAddress = '';
        this.rfqData.vesselAddress2 = '';
        this.rfqData.vesselCity = '';
        this.rfqData.vesselCountry = '';

        // Bank Details
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
        // Reset port when country changes
        this.rfqData.port = '';

        // Update available ports based on selected country
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
        // Trigger change detection to update button state
        this.cdr.detectChanges();
    }

    getTopLeftCellOptions(): string[] {
        const options: string[] = [];
        // Generate A1 through A25
        for (let i = 1; i <= 25; i++) {
            options.push('A' + i);
        }
        return options;
    }

    getTopLeftCellOptionsLimited(): string[] {
        // Return only first 10 options to limit visible items in dropdown (A1-A10)
        // Users can still type any cell reference freely (e.g., A13, A25, B5, etc.)
        return this.getTopLeftCellOptions().slice(0, 10);
    }

    onTopLeftCellChange(event: Event, tab: TabInfo): void {
        const input = event.target as HTMLInputElement;
        let value = input.value.toUpperCase();

        // Only allow one letter followed by 1-2 digits
        const regex = /^[A-Z][0-9]{1,2}$/;
        if (value && !regex.test(value)) {
            // Remove invalid characters
            value = value.replace(/[^A-Z0-9]/g, '');
            // Ensure it starts with a letter
            if (value && !/^[A-Z]/.test(value)) {
                value = '';
            }
            // Limit to one letter + max 2 digits
            const match = value.match(/^([A-Z])([0-9]{0,2})/);
            if (match) {
                value = match[0];
            } else if (value.length > 0 && /^[A-Z]/.test(value)) {
                // If it starts with a letter but has no digits yet, keep just the letter
                value = value.charAt(0);
            }
        }

        input.value = value;
        tab.topLeftCell = value;

        // Remove datalist when there's a value
        if (value && value.trim() !== '') {
            input.removeAttribute('list');

            // Re-analyze the tab with the new topLeftCell value
            // Find the file analysis that contains this tab
            for (const analysis of this.fileAnalyses) {
                if (analysis.tabs.includes(tab)) {
                    // Call async method and handle any errors
                    this.reanalyzeTabWithTopLeft(analysis, tab, value).catch(error => {
                        console.error('Error re-analyzing tab:', error);
                        this.loggingService.logError(
                            error as Error,
                            'tab_reanalysis_error',
                            'RfqComponent',
                            {
                                fileName: analysis.fileName,
                                tabName: tab.tabName,
                                topLeftCell: value
                            }
                        );
                    });
                    break;
                }
            }
        } else {
            // If value is cleared, reset column headers and dropdowns
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
        }, 'RfqComponent');
        // Trigger change detection to update button state
        this.cdr.detectChanges();
    }

    private updateExcludedState(tab: TabInfo): void {
        // Auto-exclude if rowCount is 0 or topLeftCell is blank
        tab.excluded = tab.rowCount === 0 || !tab.topLeftCell || tab.topLeftCell.trim() === '';
    }

    onTopLeftCellFocus(event: Event, fileIndex: number, tabIndex: number): void {
        const input = event.target as HTMLInputElement;
        // Only show datalist if field is empty
        if (!input.value || input.value.trim() === '') {
            input.setAttribute('list', `topLeftList-${fileIndex}-${tabIndex}`);
        }
    }

    onTopLeftCellSelectChange(event: Event, tab: TabInfo): void {
        const select = event.target as HTMLSelectElement;
        const value = select.value;

        this.loggingService.logButtonClick('top_left_cell_select_change', 'RfqComponent', {
            value: value
        });

        // If user selects "custom", clear the value to show the input field
        if (value === 'custom') {
            tab.topLeftCell = '';
            // Use setTimeout to ensure the input appears before focusing
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
            // If invalid, try to fix common issues
            if (tab.topLeftCell.trim() !== '') {
                const match = tab.topLeftCell.match(/^([A-Z])([0-9]{1,2})/);
                if (match) {
                    tab.topLeftCell = match[0];
                } else {
                    // If can't be fixed, clear it
                    tab.topLeftCell = '';
                }
            }
        }
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


