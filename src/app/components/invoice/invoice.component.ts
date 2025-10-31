import { Component, OnInit } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { DataService, ProcessedDataRow, ExcelProcessedData, ExcelItemData } from '../../services/data.service';
import { LoggingService } from '../../services/logging.service';
import * as XLSX from 'xlsx';
import * as ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import jsPDF from 'jspdf';

interface InvoiceItem {
    pos: number;
    description: string;
    remark: string;
    unit: string;
    qty: number;
    price: number;
    total: number;
    tabName: string;
    currency: string;
}

interface InvoiceData {
    invoiceNumber: string;
    invoiceDate: string;
    vessel: string;
    country: string;
    port: string;
    category: string;
    invoiceDue: string;
    items: InvoiceItem[];
    totalGBP: number;
    deliveryFee: number;
    grandTotal: number;
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
    // Fees
    portFee: number;
    agencyFee: number;
    transportCustomsLaunchFees: number;
    launchFee: number;
    discountPercent: number;
}

@Component({
    selector: 'app-invoice',
    standalone: true,
    imports: [CommonModule, FormsModule],
    templateUrl: './invoice.component.html',
    styleUrls: ['./invoice.component.scss']
})
export class InvoiceComponent implements OnInit {
    processedData: ProcessedDataRow[] = [];
    hasDataToInvoice = false;
    selectedBank: string = ''; // Default to blank
    primaryCurrency: string = '£'; // Default to GBP, will be updated from Excel file

    // Toggle switch for Split File / One Invoice
    splitFileMode: boolean = false; // false = "Split file" (default), true = "One invoice"

    // Country dropdown options
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

    // Category dropdown options
    categories = ['Provisions', 'Bonds'];

    // Available ports for selected country
    availablePorts: string[] = [];

    // Country to ports mapping
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

    invoiceData: InvoiceData = {
        invoiceNumber: '',
        invoiceDate: this.getTodayDate(),
        vessel: '',
        country: '',
        port: '',
        category: '',
        invoiceDue: '',
        items: [],
        totalGBP: 0,
        deliveryFee: 0,
        grandTotal: 0,
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
        // Fees
        portFee: 0,
        agencyFee: 0,
        transportCustomsLaunchFees: 0,
        launchFee: 0,
        discountPercent: 0
    };

    constructor(private dataService: DataService, private loggingService: LoggingService) { }

    private getTodayDate(): string {
        const today = new Date();
        return today.toISOString().split('T')[0]; // Returns YYYY-MM-DD format
    }

    onCountryChange(): void {
        // Reset port when country changes
        this.invoiceData.port = '';

        // Update available ports based on selected country
        if (this.invoiceData.country && this.countryPorts[this.invoiceData.country]) {
            this.availablePorts = this.countryPorts[this.invoiceData.country];
        } else {
            this.availablePorts = [];
        }
    }

    onBankSelectionChange(): void {
        // Clear all bank details first
        this.clearBankDetails();

        // Populate bank details based on selection
        switch (this.selectedBank) {
            case 'UK':
                this.populateUKBankDetails();
                break;
            case 'US':
                this.populateUSBankDetails();
                break;
            case 'EOS':
                this.populateEOSBankDetails();
                break;
            case '':
            default:
                // Keep fields blank
                break;
        }
    }

    onToggleChange(): void {
        // Handle toggle change logic here
        // splitFileMode = false means "Split file" (default)
        // splitFileMode = true means "One invoice"
        console.log('Toggle changed:', this.splitFileMode ? 'One invoice' : 'Split file');
        // Add any additional logic you need when the toggle changes
    }

    private clearBankDetails(): void {
        // Clear Our Company Details
        this.invoiceData.ourCompanyName = '';
        this.invoiceData.ourCompanyAddress = '';
        this.invoiceData.ourCompanyAddress2 = '';
        this.invoiceData.ourCompanyCity = '';
        this.invoiceData.ourCompanyCountry = '';
        this.invoiceData.ourCompanyPhone = '';
        this.invoiceData.ourCompanyEmail = '';

        // Clear Vessel Details
        this.invoiceData.vesselName = '';
        this.invoiceData.vesselName2 = '';
        this.invoiceData.vesselAddress = '';
        this.invoiceData.vesselAddress2 = '';
        this.invoiceData.vesselCity = '';
        this.invoiceData.vesselCountry = '';

        // Clear Bank Details
        this.invoiceData.bankName = '';
        this.invoiceData.bankAddress = '';
        this.invoiceData.iban = '';
        this.invoiceData.swiftCode = '';
        this.invoiceData.accountTitle = '';
        this.invoiceData.accountNumber = '';
        this.invoiceData.sortCode = '';
        this.invoiceData.achRouting = '';
        this.invoiceData.intermediaryBic = '';
    }

    private populateUKBankDetails(): void {
        // Our Company Details
        this.invoiceData.ourCompanyName = 'HI MARINE COMPANY LIMITED';
        this.invoiceData.ourCompanyAddress = '167-169 Great Portland Street';
        this.invoiceData.ourCompanyAddress2 = '';
        this.invoiceData.ourCompanyCity = 'London, London, W1W 5PF';
        this.invoiceData.ourCompanyCountry = 'United Kingdom';
        this.invoiceData.ourCompanyPhone = '';
        this.invoiceData.ourCompanyEmail = 'office@himarinecompany.com';

        // Vessel Details (do not auto-populate; keep blank by default)
        this.invoiceData.vesselName = '';
        this.invoiceData.vesselName2 = '';
        this.invoiceData.vesselAddress = '';
        this.invoiceData.vesselAddress2 = '';
        this.invoiceData.vesselCity = '';
        this.invoiceData.vesselCountry = '';

        // Bank Details
        this.invoiceData.bankName = 'Lloyds Bank plc';
        this.invoiceData.bankAddress = '6 Market Place, Oldham, OL11JG, United Kingdom';
        this.invoiceData.iban = 'GB84LOYD30962678553260';
        this.invoiceData.swiftCode = 'LOYDGB21446';
        this.invoiceData.accountTitle = 'HI MARINE COMPANY LIMITED';
        this.invoiceData.accountNumber = '78553260';
        this.invoiceData.sortCode = '30-96-26';
    }

    private populateUSBankDetails(): void {
        // Our Company Details
        this.invoiceData.ourCompanyName = 'HI MARINE COMPANY INC.';
        this.invoiceData.ourCompanyAddress = '9407 N.E. Vancouver Mall Drive, Suite 104';
        this.invoiceData.ourCompanyAddress2 = '';
        this.invoiceData.ourCompanyCity = 'Vancouver, WA  98662';
        this.invoiceData.ourCompanyCountry = 'USA';
        this.invoiceData.ourCompanyPhone = '+1 857 2045786';
        this.invoiceData.ourCompanyEmail = 'office@himarinecompany.com';

        // Vessel Details (do not auto-populate; keep blank by default)
        this.invoiceData.vesselName = '';
        this.invoiceData.vesselName2 = '';
        this.invoiceData.vesselAddress = '';
        this.invoiceData.vesselAddress2 = '';
        this.invoiceData.vesselCity = '';
        this.invoiceData.vesselCountry = '';

        // Bank Details
        this.invoiceData.bankName = 'Bank of America';
        this.invoiceData.bankAddress = '100 West 33d Street New York, New York 10001';
        this.invoiceData.accountNumber = '466002755612';
        this.invoiceData.swiftCode = 'BofAUS3N';
        this.invoiceData.achRouting = '011000138';
        this.invoiceData.accountTitle = 'Hi Marine Company Inc.';
    }

    private populateEOSBankDetails(): void {
        // Our Company Details
        this.invoiceData.ourCompanyName = 'EOS SUPPLY LTD';
        this.invoiceData.ourCompanyAddress = 'Wearfield, Enterprise Park East';
        this.invoiceData.ourCompanyAddress2 = 'Sunderland, Tyne and Wear, SR5 2TA';
        this.invoiceData.ourCompanyCity = '';
        this.invoiceData.ourCompanyCountry = 'United Kingdom';
        this.invoiceData.ourCompanyPhone = '';
        this.invoiceData.ourCompanyEmail = 'office@eos-supply.co.uk';

        // Vessel Details (do not auto-populate; keep blank by default)
        this.invoiceData.vesselName = '';
        this.invoiceData.vesselName2 = '';
        this.invoiceData.vesselAddress = '';
        this.invoiceData.vesselAddress2 = '';
        this.invoiceData.vesselCity = '';
        this.invoiceData.vesselCountry = '';

        // Bank Details
        this.invoiceData.bankName = 'Revolut Ltd';
        this.invoiceData.bankAddress = '7 Westferry Circus, Canary Wharf, London, England, E14 4HD';
        this.invoiceData.iban = 'GB64REVO00996912321885';
        this.invoiceData.swiftCode = 'REVOGB21XXX';
        this.invoiceData.intermediaryBic = 'CHASGB2L';
        this.invoiceData.accountTitle = 'EOS SUPPLY LTD';
        this.invoiceData.accountNumber = '69340501';
        this.invoiceData.sortCode = '04-00-75';
    }

    ngOnInit(): void {
        // Subscribe to Excel data from captains-request component
        this.dataService.excelData$.subscribe(data => {
            if (data) {
                this.convertExcelDataToInvoiceItems(data);
                this.hasDataToInvoice = this.invoiceData.items.length > 0;

                this.loggingService.logDataProcessing('invoice_data_updated', {
                    totalItems: this.invoiceData.items.length,
                    totalAmount: this.invoiceData.totalGBP,
                    finalAmount: this.invoiceData.grandTotal
                }, 'InvoiceComponent');
            }
        });

        // Also subscribe to processed data for backward compatibility
        this.dataService.processedData$.subscribe(data => {
            this.processedData = data;
            if (!this.dataService.getExcelData()) {
                this.convertToInvoiceItems(data);
                this.hasDataToInvoice = data.some(row => row.count > 0);
            }
        });
    }

    private convertExcelDataToInvoiceItems(data: ExcelProcessedData): void {
        const allItems: ExcelItemData[] = [];

        // Combine items from all tabs (PROVISIONS, FRESH PROVISIONS, BOND)
        const tabs = ['PROVISIONS', 'FRESH PROVISIONS', 'BOND'];
        tabs.forEach(tabName => {
            if (data[tabName] && data[tabName].items) {
                allItems.push(...data[tabName].items);
            }
        });

        // Convert to invoice items
        this.invoiceData.items = allItems.map((item, index) => ({
            pos: index + 1,
            description: item.description,
            remark: item.remark,
            unit: item.unit,
            qty: item.qty,
            price: item.price,
            total: item.total,
            tabName: item.tabName,
            currency: item.currency
        }));

        // Detect primary currency from the items (most common currency)
        this.detectPrimaryCurrency();

        this.calculateTotals();
    }

    private detectPrimaryCurrency(): void {
        if (this.invoiceData.items.length === 0) {
            this.primaryCurrency = '£';
            return;
        }

        // Count occurrences of each currency
        const currencyCount: { [key: string]: number } = {};
        this.invoiceData.items.forEach(item => {
            const currency = item.currency || '£';
            currencyCount[currency] = (currencyCount[currency] || 0) + 1;
        });

        // Find the most common currency
        let maxCount = 0;
        let mostCommonCurrency = '£';
        for (const [currency, count] of Object.entries(currencyCount)) {
            if (count > maxCount) {
                maxCount = count;
                mostCommonCurrency = currency;
            }
        }

        this.primaryCurrency = mostCommonCurrency;
    }

    getCurrencyLabel(currency: string): string {
        switch (currency) {
            case 'NZ$':
                return 'NZD';
            case 'A$':
                return 'AUD';
            case 'C$':
                return 'CAD';
            case '€':
                return 'EUR';
            case '$':
                return 'USD';
            case '£':
                return 'GBP';
            default:
                return 'GBP';
        }
    }

    getCurrencyExcelFormat(currency: string): string {
        switch (currency) {
            case 'NZ$':
                return '"NZ$"#,##0.00';
            case 'A$':
                return '"A$"#,##0.00';
            case 'C$':
                return '"C$"#,##0.00';
            case '€':
                return '€#,##0.00';
            case '$':
                return '$#,##0.00';
            case '£':
                return '£#,##0.00';
            default:
                return '£#,##0.00';
        }
    }

    private convertToInvoiceItems(data: ProcessedDataRow[]): void {
        const includedItems = data.filter(row => row.count > 0);
        this.invoiceData.items = includedItems.map((row, index) => ({
            pos: index + 1,
            description: row.description,
            remark: row.remarks || '',
            unit: row.unit || 'EACH',
            qty: row.count,
            price: row.price,
            total: row.count * row.price,
            tabName: 'LEGACY', // Default value for legacy data
            currency: '£' // Default to GBP for legacy data
        }));

        // For legacy data, use GBP as primary currency
        this.primaryCurrency = '£';

        this.calculateTotals();
    }

    calculateTotals(): void {
        // Sum of all item totals (pre-discount)
        this.invoiceData.totalGBP = this.invoiceData.items.reduce((sum, item) => sum + item.total, 0);

        // Discount is applied only to the items total
        const discountFactor = 1 - (this.invoiceData.discountPercent || 0) / 100;
        const discountedItemsTotal = this.invoiceData.totalGBP * discountFactor;

        // Sum all fee components
        const feesTotal =
            (this.invoiceData.deliveryFee || 0) +
            (this.invoiceData.portFee || 0) +
            (this.invoiceData.agencyFee || 0) +
            (this.invoiceData.transportCustomsLaunchFees || 0) +
            (this.invoiceData.launchFee || 0);

        // Final amount = discounted items total + fees
        this.invoiceData.grandTotal = discountedItemsTotal + feesTotal;
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
        this.loggingService.logButtonClick('generate_invoice_excel', 'InvoiceComponent', {
            totalItems: this.processedData.length,
            itemsToInvoice: this.includedItems.length,
            totalAmount: this.totalAmount,
            finalAmount: this.finalTotalAmount
        });

        const includedData = this.processedData.filter(row => row.count > 0);

        if (includedData.length === 0) {
            this.loggingService.logUserAction('invoice_generation_failed', {
                reason: 'no_items_selected'
            }, 'InvoiceComponent');
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

        this.loggingService.logExport('excel_invoice_generated', {
            fileName,
            fileSize: blob.size,
            itemsIncluded: includedData.length,
            totalAmount: this.finalTotalAmount
        }, 'InvoiceComponent');
    }

    generateInvoicePDF(): void {
        this.loggingService.logButtonClick('generate_invoice_pdf', 'InvoiceComponent', {
            totalItems: this.processedData.length,
            itemsToInvoice: this.includedItems.length,
            totalAmount: this.totalAmount,
            finalAmount: this.finalTotalAmount
        });

        const includedData = this.processedData.filter(row => row.count > 0);

        if (includedData.length === 0) {
            this.loggingService.logUserAction('pdf_invoice_generation_failed', {
                reason: 'no_items_selected'
            }, 'InvoiceComponent');
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

        this.loggingService.logExport('pdf_invoice_generated', {
            fileName,
            itemsIncluded: includedData.length,
            totalAmount: this.finalTotalAmount
        }, 'InvoiceComponent');
    }

    async exportInvoiceToExcel(): Promise<void> {
        this.loggingService.logButtonClick('export_invoice_excel', 'InvoiceComponent', {
            totalItems: this.invoiceData.items.length,
            totalAmount: this.invoiceData.totalGBP,
            finalAmount: this.invoiceData.grandTotal
        });

        if (this.invoiceData.items.length === 0) {
            this.loggingService.logUserAction('excel_export_failed', {
                reason: 'no_items_in_invoice'
            }, 'InvoiceComponent');
            alert('No items in the invoice to export.');
            return;
        }

        try {
            // Create workbook and worksheet using ExcelJS
            const workbook = new ExcelJS.Workbook();
            const worksheet = workbook.addWorksheet('Invoice');

            // Remove grid lines for cleaner look
            worksheet.properties.showGridLines = false;
            worksheet.views = [{ showGridLines: false }];

            // Top header rendering (image for HI Marine, custom header for EOS)
            try {
                if (this.selectedBank === 'EOS') {
                    // Left title: EOS SUPPLY LTD
                    worksheet.mergeCells('A2:D2');
                    const eosTitle = worksheet.getCell('A2');
                    eosTitle.value = 'EOS SUPPLY LTD';
                    eosTitle.font = { name: 'Arial', size: 18, bold: true, italic: true, color: { argb: 'FF0B2E66' } } as any;
                    eosTitle.alignment = { horizontal: 'left', vertical: 'middle' } as any;

                    // Right contact details
                    worksheet.mergeCells('E2:H2');
                    const eosPhoneUk = worksheet.getCell('E2');
                    eosPhoneUk.value = 'Phone: +44 730 7988228';
                    eosPhoneUk.font = { name: 'Arial', size: 11, color: { argb: 'FF0B2E66' } } as any;
                    eosPhoneUk.alignment = { horizontal: 'right', vertical: 'middle' } as any;

                    worksheet.mergeCells('E3:H3');
                    const eosPhoneUs = worksheet.getCell('E3');
                    eosPhoneUs.value = 'Phone: +1 857 204-5786';
                    eosPhoneUs.font = { name: 'Arial', size: 11, color: { argb: 'FF0B2E66' } } as any;
                    eosPhoneUs.alignment = { horizontal: 'right', vertical: 'middle' } as any;

                    worksheet.mergeCells('E4:H4');
                    const eosEmail = worksheet.getCell('E4');
                    eosEmail.value = 'office@eos-supply.co.uk';
                    eosEmail.font = { name: 'Arial', size: 11, color: { argb: 'FF0B2E66' } } as any;
                    eosEmail.alignment = { horizontal: 'right', vertical: 'middle' } as any;

                    // Navy bar across the sheet
                    worksheet.mergeCells('A6:H6');
                    const eosBar = worksheet.getCell('A6');
                    eosBar.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0B2E66' } } as any;
                    worksheet.getRow(6).height = 18;
                } else {
                    // Default: add HI Marine top logo image
                    const topImageResponse = await fetch('assets/images/HIMarineTopImage.png');
                    const topImageBuffer = await topImageResponse.arrayBuffer();
                    const topImageId = workbook.addImage({
                        buffer: topImageBuffer,
                        extension: 'png',
                    });
                    worksheet.addImage(topImageId, {
                        tl: { col: 1.5, row: 0.5 },
                        ext: { width: 300, height: 180 }
                    });
                }

                // Bottom border image (used for all variants)
                const bottomImageResponse = await fetch('assets/images/HIMarineBottomBorder.png');
                const bottomImageBuffer = await bottomImageResponse.arrayBuffer();
                const bottomImageId = workbook.addImage({
                    buffer: bottomImageBuffer,
                    extension: 'png',
                });
                (worksheet as any)._bottomImageId = bottomImageId;
            } catch (imageError) {
                console.warn('Could not render header/footer for Excel export:', imageError);
            }

            // Set column widths
            worksheet.getColumn('A').width = 6;   // Pos
            worksheet.getColumn('B').width = 15;  // Tab
            worksheet.getColumn('C').width = 35;  // Description  
            worksheet.getColumn('D').width = 15;  // Remark
            worksheet.getColumn('E').width = 8;   // Unit
            worksheet.getColumn('F').width = 6;   // Qty
            worksheet.getColumn('G').width = 12;  // Price
            worksheet.getColumn('H').width = 12;  // Total

            // Our Company Details (Top-Left) - write concatenated text to column A only, no merges, no wrap
            const companyHeader = worksheet.getCell('A9');
            companyHeader.value = `Name: ${this.invoiceData.ourCompanyName || ''}`;
            companyHeader.font = { size: 11, name: 'Arial', bold: true };
            companyHeader.alignment = { horizontal: 'left', vertical: 'middle', wrapText: false } as any;

            const address1 = worksheet.getCell('A10');
            address1.value = `Address: ${this.invoiceData.ourCompanyAddress || ''}`;
            address1.font = { size: 11, name: 'Arial', bold: true };
            address1.alignment = { horizontal: 'left', vertical: 'middle', wrapText: false } as any;

            const address2 = worksheet.getCell('A11');
            address2.value = `Address2: ${this.invoiceData.ourCompanyAddress2 || ''}`;
            address2.font = { size: 11, name: 'Arial', bold: true };
            address2.alignment = { horizontal: 'left', vertical: 'middle', wrapText: false } as any;

            const country = worksheet.getCell('A12');
            const leftCityLine = [this.invoiceData.ourCompanyCity, this.invoiceData.ourCompanyCountry]
                .filter(Boolean)
                .join(', ');
            country.value = `City: ${leftCityLine}`;
            country.font = { size: 11, name: 'Arial', bold: true };
            country.alignment = { horizontal: 'left', vertical: 'middle', wrapText: false } as any;

            const email = worksheet.getCell('A13');
            email.value = `Email: ${this.invoiceData.ourCompanyEmail || ''}`;
            email.font = { size: 11, name: 'Arial', bold: true };
            email.alignment = { horizontal: 'left', vertical: 'middle', wrapText: false } as any;

            // Vessel Details (Top-Right) under the logo - label + value in one cell
            const vesselLabelStyle = { font: { bold: true, size: 11, name: 'Arial' } };
            const vesselValueStyle = { font: { size: 11, name: 'Arial', bold: true } };
            worksheet.getCell('E9').value = `Name: ${this.invoiceData.vesselName || ''}`;
            worksheet.getCell('E9').font = vesselValueStyle.font;
            worksheet.getCell('F9').value = null as any;
            worksheet.getCell('G9').value = null as any;

            worksheet.getCell('E10').value = `Name2: ${this.invoiceData.vesselName2 || ''}`;
            worksheet.getCell('E10').font = vesselValueStyle.font;
            worksheet.getCell('F10').value = null as any;
            worksheet.getCell('G10').value = null as any;

            worksheet.getCell('E11').value = `Address: ${this.invoiceData.vesselAddress || ''}`;
            worksheet.getCell('E11').font = vesselValueStyle.font;
            worksheet.getCell('F11').value = null as any;
            worksheet.getCell('G11').value = null as any;

            worksheet.getCell('E12').value = `Address2: ${this.invoiceData.vesselAddress2 || ''}`;
            worksheet.getCell('E12').font = vesselValueStyle.font;
            worksheet.getCell('F12').value = null as any;
            worksheet.getCell('G12').value = null as any;

            worksheet.getCell('E13').value = `City/Country: ${[this.invoiceData.vesselCity, this.invoiceData.vesselCountry].filter(Boolean).join(', ')}`;
            worksheet.getCell('E13').font = vesselValueStyle.font;
            worksheet.getCell('F13').value = null as any;
            worksheet.getCell('G13').value = null as any;

            // Bank Details Section (Left side - Row 15-23)
            const bankDetailsStart = 15;
            const bankLabelStyle = { font: { bold: true, size: 11, name: 'Arial' } };
            const bankValueStyle = { font: { size: 11, name: 'Arial' } };
            // Merge A:D and write rich text "Label: value" (label bold)
            const writeBankLine = (row: number, label: string, value: string) => {
                worksheet.mergeCells(`A${row}:D${row}`);
                const cell = worksheet.getCell(`A${row}`);
                cell.value = {
                    richText: [
                        { text: `${label}: `, font: { bold: true, size: 11, name: 'Arial' } },
                        { text: `${value || ''}`, font: { size: 11, name: 'Arial', bold: true } }
                    ]
                } as any;
            };

            writeBankLine(bankDetailsStart, 'Bank Name', this.invoiceData.bankName);
            writeBankLine(bankDetailsStart + 1, 'Bank Address', this.invoiceData.bankAddress);
            writeBankLine(bankDetailsStart + 2, 'IBAN', this.invoiceData.iban);
            writeBankLine(bankDetailsStart + 3, 'Swift Code', this.invoiceData.swiftCode);
            writeBankLine(bankDetailsStart + 4, 'Title on Account', this.invoiceData.accountTitle);

            // Conditional bank extras
            if (this.selectedBank === 'UK') {
                worksheet.mergeCells(`A${bankDetailsStart + 6}:D${bankDetailsStart + 6}`);
                worksheet.getCell(`A${bankDetailsStart + 6}`).value = 'UK DOMESTIC WIRES:';
                worksheet.getCell(`A${bankDetailsStart + 6}`).font = { bold: true, size: 11, name: 'Arial' };

                writeBankLine(bankDetailsStart + 7, 'Account number', this.invoiceData.accountNumber);
                writeBankLine(bankDetailsStart + 8, 'Sort code', this.invoiceData.sortCode);
            }
            if (this.selectedBank === 'US') {
                writeBankLine(bankDetailsStart + 6, 'ACH Routing', this.invoiceData.achRouting || '');
            }
            if (this.selectedBank === 'EOS') {
                writeBankLine(bankDetailsStart + 6, 'Intermediary BIC', this.invoiceData.intermediaryBic || '');
            }

            // Invoice Details Section (Right side - Row 15-21)
            const invoiceDetailsStart = 15;
            const invoiceLabelStyle = { font: { bold: true, size: 11, name: 'Arial' } };
            const invoiceValueStyle = { font: { size: 11, name: 'Arial' } };
            // Helper to write invoice detail in merged E:G cell with bold label
            const writeInvoiceDetail = (offset: number, label: string, value: string) => {
                const row = invoiceDetailsStart + offset;
                const cell = worksheet.getCell(`E${row}`);
                cell.value = {
                    richText: [
                        { text: `${label}: `, font: { size: 11, name: 'Arial', bold: false } },
                        { text: `${value || ''}`, font: { size: 11, name: 'Arial', bold: false } }
                    ]
                } as any;
                cell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: false } as any;
                worksheet.getCell(`F${row}`).value = null as any;
                worksheet.getCell(`G${row}`).value = null as any;
            };

            writeInvoiceDetail(0, 'No', this.invoiceData.invoiceNumber);
            writeInvoiceDetail(1, 'Invoice Date', this.invoiceData.invoiceDate);
            writeInvoiceDetail(2, 'Vessel', this.invoiceData.vessel);
            writeInvoiceDetail(3, 'Country', this.invoiceData.country);
            writeInvoiceDetail(4, 'Port', this.invoiceData.port);
            writeInvoiceDetail(5, 'Category', this.invoiceData.category);
            writeInvoiceDetail(6, 'Invoice Due', this.invoiceData.invoiceDue);

            // Items Table (Starting from row 25)
            const tableStartRow = 25;

            // Table Headers
            const headers = ['Pos', 'Tab', 'Description', 'Remark', 'Unit', 'Qty', 'Price', 'Total'];
            headers.forEach((header, index) => {
                const cell = worksheet.getCell(tableStartRow, index + 1);
                cell.value = header;
                cell.font = { bold: true, size: 11, name: 'Arial', color: { argb: 'FFFFFFFF' } };
                const headerFillColor = this.selectedBank === 'EOS' ? 'FF0B2E66' : 'FF808080';
                cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: headerFillColor } } as any;
                cell.alignment = { horizontal: 'center', vertical: 'middle' };
            });

            // Table Data
            this.invoiceData.items.forEach((item, index) => {
                const rowIndex = tableStartRow + 1 + index;

                // Pos
                const posCell = worksheet.getCell(rowIndex, 1);
                posCell.value = item.pos;
                posCell.font = { size: 10, name: 'Arial' };
                posCell.alignment = { horizontal: 'center', vertical: 'middle' };

                // Tab
                const tabCell = worksheet.getCell(rowIndex, 2);
                tabCell.value = item.tabName;
                tabCell.font = { size: 10, name: 'Arial' };
                tabCell.alignment = { horizontal: 'center', vertical: 'middle' };

                // Description
                const descCell = worksheet.getCell(rowIndex, 3);
                descCell.value = item.description;
                descCell.font = { size: 10, name: 'Arial' };
                descCell.alignment = { horizontal: 'left', vertical: 'middle' };

                // Remark
                const remarkCell = worksheet.getCell(rowIndex, 4);
                remarkCell.value = item.remark;
                remarkCell.font = { size: 10, name: 'Arial' };
                remarkCell.alignment = { horizontal: 'left', vertical: 'middle' };

                // Unit
                const unitCell = worksheet.getCell(rowIndex, 5);
                unitCell.value = item.unit;
                unitCell.font = { size: 10, name: 'Arial' };
                unitCell.alignment = { horizontal: 'center', vertical: 'middle' };

                // Qty (default to 0)
                const qtyCell = worksheet.getCell(rowIndex, 6);
                qtyCell.value = (item.qty ?? 0);
                qtyCell.font = { size: 10, name: 'Arial' };
                qtyCell.alignment = { horizontal: 'center', vertical: 'middle' };

                // Price (rounded to nearest penny)
                const priceCell = worksheet.getCell(rowIndex, 7);
                priceCell.value = Math.round(item.price * 100) / 100;
                priceCell.font = { size: 10, name: 'Arial' };
                priceCell.alignment = { horizontal: 'right', vertical: 'middle' };
                // Use dynamic currency formatting based on item currency
                const currencyFormat = this.getCurrencyExcelFormat(item.currency);
                priceCell.numFmt = currencyFormat;

                // Total (formula = F * G)
                const totalCell = worksheet.getCell(rowIndex, 8);
                totalCell.value = { formula: `F${rowIndex}*G${rowIndex}` } as any;
                totalCell.font = { size: 10, name: 'Arial' };
                totalCell.alignment = { horizontal: 'right', vertical: 'middle' };
                totalCell.numFmt = currencyFormat;
            });

            // Totals and Fees Section
            let totalsStartRow = tableStartRow + this.invoiceData.items.length + 2; // slightly tighter spacing

            // List discount (amount) and non-zero fees just above the Total
            const feeLines: { label: string; value?: number; includeInSum?: boolean }[] = [];
            const discountAmount = this.invoiceData.totalGBP * (this.invoiceData.discountPercent || 0) / 100;
            if (discountAmount > 0) feeLines.push({ label: 'Discount:', value: -discountAmount, includeInSum: false });
            if (this.invoiceData.deliveryFee) feeLines.push({ label: 'Delivery fee:', value: this.invoiceData.deliveryFee, includeInSum: true });
            if (this.invoiceData.portFee) feeLines.push({ label: 'Port fee:', value: this.invoiceData.portFee, includeInSum: true });
            if (this.invoiceData.agencyFee) feeLines.push({ label: 'Agency fee:', value: this.invoiceData.agencyFee, includeInSum: true });
            if (this.invoiceData.transportCustomsLaunchFees) feeLines.push({ label: 'Transport, Customs, Launch fees:', value: this.invoiceData.transportCustomsLaunchFees, includeInSum: true });
            if (this.invoiceData.launchFee) feeLines.push({ label: 'Launch:', value: this.invoiceData.launchFee, includeInSum: true });

            // Get currency format for totals
            const primaryCurrencyFormat = this.getCurrencyExcelFormat(this.primaryCurrency);
            const totalLabel = `TOTAL ${this.getCurrencyLabel(this.primaryCurrency)}`;

            // Track fee rows that contain numeric amounts for later SUM
            const feeAmountRowRefs: string[] = [];
            feeLines.forEach((fee, idx) => {
                const rowIndex = totalsStartRow + idx;
                const labelCell = worksheet.getCell(`G${rowIndex}`);
                labelCell.value = fee.label;
                labelCell.font = { bold: true, size: 11, name: 'Arial' };
                labelCell.alignment = { horizontal: 'right', vertical: 'middle' };

                const valueCell = worksheet.getCell(`H${rowIndex}`);
                valueCell.value = fee.value as number;
                valueCell.numFmt = primaryCurrencyFormat;
                if (fee.includeInSum) feeAmountRowRefs.push(`H${rowIndex}`);
                valueCell.font = { bold: true, size: 11, name: 'Arial' };
                valueCell.alignment = { horizontal: 'right', vertical: 'middle' };

                // Use standard line height like the rest of the page
                worksheet.getRow(rowIndex).height = 15; // typical default
            });

            totalsStartRow += feeLines.length;

            // Draw a line above the total row across columns F-H
            for (let col = 6; col <= 8; col++) { // Column F=6, G=7, H=8
                const cell = worksheet.getCell(totalsStartRow, col);
                cell.border = {
                    top: { style: 'thin', color: { argb: 'FF000000' } }
                } as any;
            }

            // TOTAL (formula: sum of H column item totals + all monetary fee cells)
            worksheet.getCell(`G${totalsStartRow}`).value = totalLabel;
            worksheet.getCell(`G${totalsStartRow}`).font = { bold: true, size: 11, name: 'Arial' };
            worksheet.getCell(`G${totalsStartRow}`).alignment = { horizontal: 'right', vertical: 'middle' };

            const firstDataRow = tableStartRow + 1;
            const lastDataRow = tableStartRow + this.invoiceData.items.length;
            const itemsSumFormula = `SUM(H${firstDataRow}:H${lastDataRow})`;
            const feeSumPart = feeAmountRowRefs.length ? `+${feeAmountRowRefs.join('+')}` : '';
            const discountFactor = this.invoiceData.discountPercent ? `(1-${this.invoiceData.discountPercent}/100)` : '1';
            const totalFormula = `(${itemsSumFormula}*${discountFactor})${feeSumPart}`;
            worksheet.getCell(`H${totalsStartRow}`).value = { formula: totalFormula } as any;
            worksheet.getCell(`H${totalsStartRow}`).font = { bold: true, size: 11, name: 'Arial' };
            worksheet.getCell(`H${totalsStartRow}`).alignment = { horizontal: 'right', vertical: 'middle' };
            worksheet.getCell(`H${totalsStartRow}`).numFmt = primaryCurrencyFormat;

            // Removed separate grand total line; the total row represents the final amount

            // Terms and Conditions
            const termsStartRow = totalsStartRow + 4;
            worksheet.mergeCells(`A${termsStartRow}:G${termsStartRow}`);
            worksheet.getCell(`A${termsStartRow}`).value = 'By placing the order according to the above quotation you are accepting the following terms:';
            worksheet.getCell(`A${termsStartRow}`).font = { bold: true, size: 11, name: 'Arial' };

            const terms = [
                'I.    Credit days: 30 calendar days.',
                'II.   Accounts not paid in this time frame will be charged 10% interest rate per month, any discount given will be null and void.',
                'III.  Should collection or legal action be required to collect past dues, fees for such action will be added to your account.',
                'IV.   Subject to unsold. Final weights are subject to vendor packing standards.',
                'V.    If the transaction is canceled after the order is authorized, we have the right to collect the invoice without claim.'
            ];

            terms.forEach((term, index) => {
                worksheet.mergeCells(`A${termsStartRow + 1 + index}:G${termsStartRow + 1 + index}`);
                worksheet.getCell(`A${termsStartRow + 1 + index}`).value = term;
                worksheet.getCell(`A${termsStartRow + 1 + index}`).font = { size: 10, name: 'Arial' };
            });

            // If EOS category selected, print company and bank details at bottom (left and right blocks)
            if (this.selectedBank === 'EOS') {
                const eosFont = { name: 'Arial', size: 11, bold: true, color: { argb: 'FF0B2E66' } } as any;
                const bottomSectionStart = termsStartRow + terms.length + 4;

                // Left block: Our Company Details in column A
                const leftLines: string[] = [
                    this.invoiceData.ourCompanyName || '',
                    this.invoiceData.ourCompanyAddress || '',
                    this.invoiceData.ourCompanyAddress2 || '',
                    [this.invoiceData.ourCompanyCity, this.invoiceData.ourCompanyCountry].filter(Boolean).join(', ')
                ].filter(Boolean);

                leftLines.forEach((text, idx) => {
                    const rowIndex = bottomSectionStart + idx;
                    const cell = worksheet.getCell(`A${rowIndex}`);
                    cell.value = text;
                    cell.font = eosFont;
                    cell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: false } as any;
                });

                // Right block: Bank Details starting at column E; merge to allow long text on one line
                const rightStartRow = bottomSectionStart;
                const writeRight = (offset: number, text: string) => {
                    const row = rightStartRow + offset;
                    worksheet.mergeCells(`E${row}:H${row}`);
                    const cell = worksheet.getCell(`E${row}`);
                    cell.value = text;
                    cell.font = eosFont;
                    cell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: false } as any;
                };

                writeRight(0, `Bank Name: ${this.invoiceData.bankName || ''}`);
                writeRight(1, `Bank Address: ${this.invoiceData.bankAddress || ''}`);
                writeRight(2, `IBAN: ${this.invoiceData.iban || ''}`);
                writeRight(3, `SWIFTBIC: ${this.invoiceData.swiftCode || ''}`);
            }

            // Add bottom image
            const bottomImageRowPosition = termsStartRow + terms.length + 8; // Leave extra spacing after shifting sections down
            if ((worksheet as any)._bottomImageId) {
                worksheet.addImage((worksheet as any)._bottomImageId, {
                    tl: { col: 0, row: bottomImageRowPosition }, // Position at bottom
                    ext: { width: 667, height: 80 } // 2/3 of previous width: 667x80 pixels
                });
            }

            // Set print area: columns A-H starting from row 1, ending 3 rows below the bottom image
            // bottomImageRowPosition is 0-based, so add 1 to convert to 1-based, then add 3 more rows
            const printAreaEndRow = bottomImageRowPosition + 1 + 3;
            worksheet.pageSetup.printArea = `A1:H${printAreaEndRow}`;

            // Generate Excel file
            const buffer = await workbook.xlsx.writeBuffer();
            const blob = new Blob([buffer], {
                type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            });

            // Download file
            const filePrefix = this.selectedBank === 'EOS' ? 'EOS_Invoice' : 'HIMarine_Invoice';
            const fileName = `${filePrefix}_${this.invoiceData.invoiceNumber}_${new Date().toISOString().split('T')[0]}.xlsx`;
            saveAs(blob, fileName);

            this.loggingService.logExport('excel_invoice_exported', {
                fileName,
                fileSize: blob.size,
                itemsIncluded: this.invoiceData.items.length,
                totalAmount: this.invoiceData.grandTotal
            }, 'InvoiceComponent');

        } catch (error) {
            this.loggingService.logError(error as Error, 'excel_export', 'InvoiceComponent');
            alert('An error occurred while exporting to Excel. Please try again.');
        }
    }

    exportInvoiceToPDF(): void {
        this.loggingService.logButtonClick('export_invoice_pdf', 'InvoiceComponent', {
            totalItems: this.invoiceData.items.length,
            totalAmount: this.invoiceData.totalGBP,
            finalAmount: this.invoiceData.grandTotal
        });

        if (this.invoiceData.items.length === 0) {
            this.loggingService.logUserAction('pdf_export_failed', {
                reason: 'no_items_in_invoice'
            }, 'InvoiceComponent');
            alert('No items in the invoice to export.');
            return;
        }

        // Create new PDF document
        const doc = new jsPDF();
        const pageWidth = doc.internal.pageSize.getWidth();
        const margin = 15;

        // Add company header
        doc.setFontSize(20);
        doc.setFont('helvetica', 'bold');
        doc.text('HI MARINE COMPANY LIMITED', pageWidth / 2, 20, { align: 'center' });

        doc.setFontSize(10);
        doc.setFont('helvetica', 'normal');
        doc.text('Wearfield, Enterprise Park East, Sunderland, Tyne and Wear, SR5 2TA', pageWidth / 2, 30, { align: 'center' });
        doc.text('United Kingdom', pageWidth / 2, 35, { align: 'center' });
        doc.text('office@himarinecompany.com', pageWidth / 2, 40, { align: 'center' });

        // Invoice details - Left side
        let yPos = 55;
        doc.setFontSize(12);
        doc.setFont('helvetica', 'bold');
        doc.text('Invoice Details:', margin, yPos);

        yPos += 10;
        doc.setFontSize(10);
        doc.setFont('helvetica', 'normal');
        doc.text(`Invoice No: ${this.invoiceData.invoiceNumber}`, margin, yPos);
        yPos += 7;
        doc.text(`Invoice Date: ${this.invoiceData.invoiceDate}`, margin, yPos);
        yPos += 7;
        doc.text(`Vessel: ${this.invoiceData.vessel}`, margin, yPos);
        yPos += 7;
        doc.text(`Country: ${this.invoiceData.country}`, margin, yPos);
        yPos += 7;
        doc.text(`Port: ${this.invoiceData.port}`, margin, yPos);
        yPos += 7;
        doc.text(`Category: ${this.invoiceData.category}`, margin, yPos);
        yPos += 7;
        doc.text(`Invoice Due: ${this.invoiceData.invoiceDue}`, margin, yPos);

        // Bank details - Right side
        yPos = 55;
        const rightMargin = pageWidth / 2 + 10;
        doc.setFontSize(12);
        doc.setFont('helvetica', 'bold');
        doc.text('Bank Details:', rightMargin, yPos);

        yPos += 10;
        doc.setFontSize(10);
        doc.setFont('helvetica', 'normal');
        doc.text(`Bank Name: ${this.invoiceData.bankName}`, rightMargin, yPos);
        yPos += 7;
        doc.text(`Bank Address: ${this.invoiceData.bankAddress}`, rightMargin, yPos, { maxWidth: pageWidth / 2 - 20 });
        yPos += 14;
        doc.text(`IBAN: ${this.invoiceData.iban}`, rightMargin, yPos);
        yPos += 7;
        doc.text(`Swift Code: ${this.invoiceData.swiftCode}`, rightMargin, yPos);
        yPos += 7;
        doc.text(`Account Title: ${this.invoiceData.accountTitle}`, rightMargin, yPos, { maxWidth: pageWidth / 2 - 20 });
        yPos += 14;
        doc.text(`Account Number: ${this.invoiceData.accountNumber}`, rightMargin, yPos);
        yPos += 7;
        doc.text(`Sort Code: ${this.invoiceData.sortCode}`, rightMargin, yPos);

        // Items table
        yPos = 140;
        doc.setFontSize(12);
        doc.setFont('helvetica', 'bold');
        doc.text('Items:', margin, yPos);

        yPos += 10;
        // Table header
        const colWidths = [15, 20, 50, 25, 15, 15, 20, 20];
        const colPositions = [margin];
        for (let i = 1; i < colWidths.length; i++) {
            colPositions.push(colPositions[i - 1] + colWidths[i - 1]);
        }

        doc.setFontSize(10);
        doc.setFont('helvetica', 'bold');
        const headers = ['Pos', 'Tab', 'Description', 'Remark', 'Unit', 'Qty', 'Price', 'Total'];
        headers.forEach((header, index) => {
            doc.text(header, colPositions[index], yPos);
        });

        // Draw header line
        doc.line(margin, yPos + 3, pageWidth - margin, yPos + 3);
        yPos += 10;

        // Table rows
        doc.setFont('helvetica', 'normal');
        this.invoiceData.items.forEach((item) => {
            if (yPos > doc.internal.pageSize.getHeight() - 50) {
                doc.addPage();
                yPos = 20;
            }

            doc.text(item.pos.toString(), colPositions[0], yPos);
            doc.text(item.tabName.substring(0, 8) + (item.tabName.length > 8 ? '...' : ''), colPositions[1], yPos);
            doc.text(item.description.substring(0, 20) + (item.description.length > 20 ? '...' : ''), colPositions[2], yPos);
            doc.text(item.remark.substring(0, 10) + (item.remark.length > 10 ? '...' : ''), colPositions[3], yPos);
            doc.text(item.unit, colPositions[4], yPos);
            doc.text(item.qty.toString(), colPositions[5], yPos);
            doc.text(`${item.currency}${item.price.toFixed(2)}`, colPositions[6], yPos);
            doc.text(`${item.currency}${item.total.toFixed(2)}`, colPositions[7], yPos);

            yPos += 8;
        });

        // Totals and Fees
        const totalCurrencyLabel = this.getCurrencyLabel(this.primaryCurrency);
        yPos += 10;
        doc.line(margin, yPos, pageWidth - margin, yPos);
        yPos += 10;

        doc.setFont('helvetica', 'bold');
        doc.text(`TOTAL ${totalCurrencyLabel}: ${this.primaryCurrency}${this.invoiceData.totalGBP.toFixed(2)}`, pageWidth - 100, yPos);
        yPos += 8;
        doc.setFont('helvetica', 'normal');
        const discountAmountForPdf = this.invoiceData.totalGBP * (this.invoiceData.discountPercent || 0) / 100;
        if (discountAmountForPdf > 0) {
            doc.text(`Discount: -${this.primaryCurrency}${discountAmountForPdf.toFixed(2)}`, pageWidth - 100, yPos);
            yPos += 8;
        }
        if ((this.invoiceData.deliveryFee || 0) > 0) { doc.text(`Delivery fee: ${this.primaryCurrency}${(this.invoiceData.deliveryFee || 0).toFixed(2)}`, pageWidth - 100, yPos); yPos += 8; }
        if ((this.invoiceData.portFee || 0) > 0) { doc.text(`Port fee: ${this.primaryCurrency}${(this.invoiceData.portFee || 0).toFixed(2)}`, pageWidth - 100, yPos); yPos += 8; }
        if ((this.invoiceData.agencyFee || 0) > 0) { doc.text(`Agency fee: ${this.primaryCurrency}${(this.invoiceData.agencyFee || 0).toFixed(2)}`, pageWidth - 100, yPos); yPos += 8; }
        if ((this.invoiceData.transportCustomsLaunchFees || 0) > 0) { doc.text(`Transport/Customs/Launch: ${this.primaryCurrency}${(this.invoiceData.transportCustomsLaunchFees || 0).toFixed(2)}`, pageWidth - 100, yPos); yPos += 8; }
        if ((this.invoiceData.launchFee || 0) > 0) { doc.text(`Launch: ${this.primaryCurrency}${(this.invoiceData.launchFee || 0).toFixed(2)}`, pageWidth - 100, yPos); yPos += 8; }
        yPos += 10;
        doc.line(margin, yPos, pageWidth - margin, yPos);
        yPos += 10;
        doc.setFont('helvetica', 'bold');
        doc.setFontSize(12);
        doc.text(`GRAND TOTAL: ${this.primaryCurrency}${this.invoiceData.grandTotal.toFixed(2)}`, pageWidth - 110, yPos);

        // Save the PDF
        const fileName = `HIMarine_Invoice_${this.invoiceData.invoiceNumber}_${new Date().toISOString().split('T')[0]}.pdf`;
        doc.save(fileName);

        this.loggingService.logExport('pdf_invoice_exported', {
            fileName,
            itemsIncluded: this.invoiceData.items.length,
            totalAmount: this.invoiceData.grandTotal
        }, 'InvoiceComponent');
    }
}

