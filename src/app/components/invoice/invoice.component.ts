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
    // Export
    exportFileName?: string;
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
    selectedBank: string = ''; // Default to empty (no selection)
    primaryCurrency: string = '£'; // Default to GBP, will be updated from Excel file

    // Toggle switch for Split Invoices / One Invoice
    splitFileMode: boolean = false; // false = "Split invoices" (default), true = "One invoice"

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
        discountPercent: 0,
        // Export
        exportFileName: ''
    };

    constructor(private dataService: DataService, private loggingService: LoggingService) { }

    private getTodayDate(): string {
        const today = new Date();
        return today.toISOString().split('T')[0]; // Returns YYYY-MM-DD format
    }

    private generateAutoFileName(): void {
        const parts: string[] = [];

        // Invoice prefix
        parts.push('Invoice');

        // Invoice Number
        if (this.invoiceData.invoiceNumber) {
            parts.push(this.invoiceData.invoiceNumber);
        }

        // Today's Date
        if (this.invoiceData.invoiceDate) {
            const date = new Date(this.invoiceData.invoiceDate);
            const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sept', 'Oct', 'Nov', 'Dec'];
            const month = months[date.getMonth()];
            const day = date.getDate();
            const year = date.getFullYear();
            parts.push(`${day}-${month}-${year}`);
        }

        // Vessel (from Invoice Details)
        if (this.invoiceData.vessel) {
            parts.push(this.invoiceData.vessel);
        }

        // Country
        if (this.invoiceData.country) {
            parts.push(this.invoiceData.country);
        }

        // Port
        if (this.invoiceData.port) {
            parts.push(this.invoiceData.port);
        }

        // Category
        if (this.invoiceData.category) {
            // If category is 'Bonds and Provisions', use placeholder
            if (this.invoiceData.category === 'Bonds and Provisions') {
                parts.push('<Category>');
            } else {
                parts.push(this.invoiceData.category);
            }
        }

        // Invoice Due
        if (this.invoiceData.invoiceDue) {
            parts.push(this.invoiceData.invoiceDue);
        }

        this.invoiceData.exportFileName = parts.join(' ');
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

        // Update filename
        this.generateAutoFileName();
    }

    onInvoiceDetailChange(): void {
        // Update filename when invoice details change
        this.generateAutoFileName();
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
            default:
                // Keep fields blank
                break;
        }
    }

    onToggleChange(): void {
        // Handle toggle change logic here
        // splitFileMode = false means "Split invoices" (default)
        // splitFileMode = true means "One invoice"
        console.log('Toggle changed:', this.splitFileMode ? 'One invoice' : 'Split invoices');
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
        this.invoiceData.ourCompanyAddress = '85 Great Portland Street, First Floor';
        this.invoiceData.ourCompanyAddress2 = '';
        this.invoiceData.ourCompanyCity = 'London, England, W1W 7LT';
        this.invoiceData.ourCompanyCountry = 'United Kingdom';
        this.invoiceData.ourCompanyPhone = '';
        this.invoiceData.ourCompanyEmail = '';

        // Vessel Details (do not auto-populate; keep blank by default)
        this.invoiceData.vesselName = '';
        this.invoiceData.vesselName2 = '';
        this.invoiceData.vesselAddress = '';
        this.invoiceData.vesselAddress2 = '';
        this.invoiceData.vesselCity = '';
        this.invoiceData.vesselCountry = '';

        // Bank Details
        this.invoiceData.bankName = 'Revolut Ltd';
        this.invoiceData.bankAddress = '7 Westferry Circus, London, England, E14 4HD';
        this.invoiceData.iban = 'GB64REVO00996912321885';
        this.invoiceData.swiftCode = 'REVOGB21XXX';
        this.invoiceData.intermediaryBic = 'CHASGB2L';
        this.invoiceData.accountTitle = 'EOS SUPPLY LTD';
        this.invoiceData.accountNumber = '69340501';
        this.invoiceData.sortCode = '04-00-75';
    }

    ngOnInit(): void {
        // Subscribe to Excel data from captains-order component
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

        // Auto-determine category based on tabName values
        this.determineCategoryFromItems();

        // Detect primary currency from the items (most common currency)
        this.detectPrimaryCurrency();

        this.calculateTotals();
    }

    private determineCategoryFromItems(): void {
        if (this.invoiceData.items.length === 0) {
            this.invoiceData.category = '';
            this.generateAutoFileName();
            return;
        }

        // Collect unique tabName values
        const uniqueTabs = new Set(this.invoiceData.items.map(item => item.tabName));

        // Determine category based on tabName values
        if (uniqueTabs.has('BOND') && !uniqueTabs.has('PROVISIONS') && !uniqueTabs.has('FRESH PROVISIONS')) {
            // Only BOND
            this.invoiceData.category = 'Bond';
        } else if ((uniqueTabs.has('PROVISIONS') || uniqueTabs.has('FRESH PROVISIONS')) && !uniqueTabs.has('BOND')) {
            // Only PROVISIONS or FRESH PROVISIONS
            this.invoiceData.category = 'Provisions';
        } else if (uniqueTabs.has('BOND') && (uniqueTabs.has('PROVISIONS') || uniqueTabs.has('FRESH PROVISIONS'))) {
            // Both BOND and PROVISIONS/FRESH PROVISIONS
            this.invoiceData.category = 'Bonds and Provisions';
        } else {
            // Fallback for any other cases
            this.invoiceData.category = '';
        }

        this.generateAutoFileName();
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

        // Auto-determine category based on tabName values
        this.determineCategoryFromItems();

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

    get isExportDisabled(): boolean {
        // Disable if no items
        if (!this.invoiceData.items || this.invoiceData.items.length === 0) {
            return true;
        }
        // Disable if category is empty, null, undefined, or 'Blank'
        if (!this.invoiceData.category || this.invoiceData.category === '' || this.invoiceData.category === 'Blank') {
            return true;
        }
        return false;
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
            // Check if split file mode is enabled
            if (!this.splitFileMode) {
                // Split file mode: Create separate files for Bonded and Provisions
                await this.exportSplitFiles();
            } else {
                // One invoice mode: Create single file (existing behavior)
                await this.exportSingleInvoiceFile(this.invoiceData.items, this.invoiceData.category, false, true);
            }
        } catch (error) {
            this.loggingService.logError(error as Error, 'excel_export', 'InvoiceComponent');
            alert('An error occurred while exporting to Excel. Please try again.');
        }
    }

    private async exportSplitFiles(): Promise<void> {
        // Group items by category
        const bondedItems = this.invoiceData.items.filter(item =>
            item.tabName === 'BOND'
        );
        const provisionsItems = this.invoiceData.items.filter(item =>
            item.tabName === 'PROVISIONS' || item.tabName === 'FRESH PROVISIONS'
        );

        // Check if both categories exist
        const hasBothCategories = bondedItems.length > 0 && provisionsItems.length > 0;

        // Determine which category should receive fees:
        // - If both exist: fees go to Provisions only
        // - If only Provisions: fees go to Provisions
        // - If only Bond: fees go to Bond
        const provisionsShouldGetFees = provisionsItems.length > 0;
        const bondShouldGetFees = bondedItems.length > 0 && provisionsItems.length === 0;

        // Export each category if it has items
        if (bondedItems.length > 0) {
            await this.exportSingleInvoiceFile(bondedItems, 'Bond', false, bondShouldGetFees);
        }
        if (provisionsItems.length > 0) {
            // Append "A" to invoice number for provisions if both categories exist
            await this.exportSingleInvoiceFile(provisionsItems, 'Provisions', hasBothCategories, provisionsShouldGetFees);
        }

        // Log export
        const filesCreated = [bondedItems.length > 0, provisionsItems.length > 0].filter(Boolean).length;
        this.loggingService.logExport('excel_invoice_exported_split', {
            filesCreated,
            bondedItems: bondedItems.length,
            provisionsItems: provisionsItems.length
        }, 'InvoiceComponent');
    }

    private async exportSingleInvoiceFile(items: InvoiceItem[], categoryOverride?: string, appendAtoInvoiceNumber?: boolean, includeFees: boolean = true): Promise<void> {
        try {
            // Create workbook and worksheet using ExcelJS
            const workbook = new ExcelJS.Workbook();
            const worksheet = workbook.addWorksheet('Invoice');

            // Remove grid lines for cleaner look
            worksheet.properties.showGridLines = false;
            worksheet.views = [{ showGridLines: false }];

            // Top header rendering (image for Hi Marine, custom header for EOS)
            try {
                if (this.selectedBank === 'EOS') {
                    // Left title: EOS SUPPLY LTD
                    worksheet.mergeCells('A2:D2');
                    const eosTitle = worksheet.getCell('A2');
                    eosTitle.value = 'EOS SUPPLY LTD';
                    eosTitle.font = { name: 'Calibri', size: 18, bold: true, italic: true, color: { argb: 'FF0B2E66' } } as any;
                    eosTitle.alignment = { horizontal: 'left', vertical: 'middle' } as any;

                    // Right contact details
                    worksheet.mergeCells('E2:G2');
                    const eosPhoneUk = worksheet.getCell('E2');
                    eosPhoneUk.value = 'Phone: +44 730 7988228';
                    eosPhoneUk.font = { name: 'Calibri', size: 11, color: { argb: 'FF0B2E66' } } as any;
                    eosPhoneUk.alignment = { horizontal: 'right', vertical: 'middle' } as any;

                    worksheet.mergeCells('E3:G3');
                    const eosPhoneUs = worksheet.getCell('E3');
                    eosPhoneUs.value = 'Phone: +1 857 204-5786';
                    eosPhoneUs.font = { name: 'Calibri', size: 11, color: { argb: 'FF0B2E66' } } as any;
                    eosPhoneUs.alignment = { horizontal: 'right', vertical: 'middle' } as any;

                    worksheet.mergeCells('E4:G4');
                    const eosEmail = worksheet.getCell('E4');
                    eosEmail.value = 'office@eos-supply.co.uk';
                    eosEmail.font = { name: 'Calibri', size: 11, color: { argb: 'FF0B2E66' } } as any;
                    eosEmail.alignment = { horizontal: 'right', vertical: 'middle' } as any;

                    // Navy bar across the sheet
                    worksheet.mergeCells('A6:G6');
                    const eosBar = worksheet.getCell('A6');
                    eosBar.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0B2E66' } } as any;
                    worksheet.getRow(6).height = 18;
                } else {
                    // Default: add Hi Marine top logo image
                    const topImageResponse = await fetch('assets/images/HIMarineTopImage_sm.png');
                    const topImageBuffer = await topImageResponse.arrayBuffer();

                    // Get original image dimensions to prevent stretching
                    const imageBlob = new Blob([topImageBuffer], { type: 'image/png' });
                    const imageUrl = URL.createObjectURL(imageBlob);
                    const img = new Image();

                    // Wait for image to load to get natural dimensions
                    await new Promise<void>((resolve, reject) => {
                        img.onload = () => {
                            URL.revokeObjectURL(imageUrl);
                            resolve();
                        };
                        img.onerror = () => {
                            URL.revokeObjectURL(imageUrl);
                            reject(new Error('Failed to load image'));
                        };
                        img.src = imageUrl;
                    });

                    const topImageId = workbook.addImage({
                        buffer: topImageBuffer,
                        extension: 'png',
                    });
                    // Use original image dimensions to prevent resizing or stretching
                    // Position image 45 pixels from left edge (approximately 0.75 column units)
                    worksheet.addImage(topImageId, {
                        tl: { col: 0.75, row: 0.5 },
                        ext: { width: img.naturalWidth, height: img.naturalHeight }
                    });
                }

                // Bottom border image (used for all variants)
                const bottomImagePath = this.selectedBank === 'EOS'
                    ? 'assets/images/EosSupplyLtdBottomBorder.png'
                    : 'assets/images/HIMarineBottomBorder.png';
                const bottomImageResponse = await fetch(bottomImagePath);
                const bottomImageBuffer = await bottomImageResponse.arrayBuffer();

                // Get original image dimensions to maintain aspect ratio
                const imageBlob = new Blob([bottomImageBuffer], { type: 'image/png' });
                const imageUrl = URL.createObjectURL(imageBlob);
                const img = new Image();

                // Wait for image to load to get natural dimensions
                await new Promise<void>((resolve, reject) => {
                    img.onload = () => {
                        URL.revokeObjectURL(imageUrl);
                        resolve();
                    };
                    img.onerror = () => {
                        URL.revokeObjectURL(imageUrl);
                        reject(new Error('Failed to load bottom image'));
                    };
                    img.src = imageUrl;
                });

                const bottomImageId = workbook.addImage({
                    buffer: bottomImageBuffer,
                    extension: 'png',
                });
                // Store image ID and dimensions for later use
                (worksheet as any)._bottomImageId = bottomImageId;
                (worksheet as any)._bottomImageWidth = img.naturalWidth;
                (worksheet as any)._bottomImageHeight = img.naturalHeight;
            } catch (imageError) {
                console.warn('Could not render header/footer for Excel export:', imageError);
            }

            // Set column widths (converted from pixels to Excel character units: ~7 pixels per unit)
            worksheet.getColumn('A').width = 56 / 7;   // Pos: 56 pixels
            worksheet.getColumn('B').width = 374 / 7;   // Description: 374 pixels
            worksheet.getColumn('C').width = 254 / 7;   // Remark: 254 pixels
            worksheet.getColumn('D').width = 80 / 7;    // Unit: 80 pixels
            worksheet.getColumn('E').width = 82 / 7;   // Qty: 82 pixels
            worksheet.getColumn('F').width = 131 / 7;  // Price: 131 pixels
            worksheet.getColumn('G').width = 120 / 7;  // Total: 120 pixels

            // Our Company Details (Top-Left) - write concatenated text to column A only, no merges, no wrap
            // Only write rows that have data to avoid blank lines
            // Skip Email field when 'US' or 'UK' is selected
            // For EOS, start at row 8 instead of 9 to remove the empty row 7
            let companyRow = this.selectedBank === 'EOS' ? 8 : 9;
            const companyDetails = [
                { label: 'Name', value: this.invoiceData.ourCompanyName },
                { label: 'Address', value: this.invoiceData.ourCompanyAddress },
                { label: 'Address2', value: this.invoiceData.ourCompanyAddress2 },
                { label: 'City', value: [this.invoiceData.ourCompanyCity, this.invoiceData.ourCompanyCountry].filter(Boolean).join(', ') },
                { label: 'Phone', value: this.invoiceData.ourCompanyPhone },
                { label: 'Email', value: this.invoiceData.ourCompanyEmail, skipForUSUK: true }
            ];
            companyDetails.forEach(detail => {
                // Skip Email field when 'US' or 'UK' is selected
                if (detail.skipForUSUK && (this.selectedBank === 'US' || this.selectedBank === 'UK')) {
                    return;
                }
                if (detail.value && detail.value.trim()) {
                    const cell = worksheet.getCell(`A${companyRow}`);
                    cell.value = detail.value;
                    cell.font = { size: 11, name: 'Calibri', bold: true };
                    cell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: false } as any;
                    companyRow++;
                }
            });

            // Vessel Details (Top-Right) under the logo - values only, no labels
            // Only write rows that have data to avoid blank lines
            // For EOS, start at row 8 instead of 9 to match companyRow
            let vesselRow = this.selectedBank === 'EOS' ? 8 : 9;
            const vesselDetails = [
                this.invoiceData.vesselName,
                this.invoiceData.vesselName2,
                this.invoiceData.vesselAddress,
                this.invoiceData.vesselAddress2,
                [this.invoiceData.vesselCity, this.invoiceData.vesselCountry].filter(Boolean).join(', ')
            ];
            const vesselValueStyle = { font: { size: 11, name: 'Calibri', bold: true } };
            vesselDetails.forEach(value => {
                if (value && value.trim()) {
                    worksheet.getCell(`E${vesselRow}`).value = value;
                    worksheet.getCell(`E${vesselRow}`).font = vesselValueStyle.font;
                    worksheet.getCell(`F${vesselRow}`).value = null as any;
                    worksheet.getCell(`G${vesselRow}`).value = null as any;
                    vesselRow++;
                }
            });

            // Calculate the end of company/vessel details section and start bank/invoice details right after
            // Use the maximum row from company or vessel details, then add 1 row gap
            const topSectionEndRow = Math.max(companyRow, vesselRow);
            const bankDetailsStartRow = topSectionEndRow + 1;

            // Bank Details Section (Left side)
            // Include bank details for all companies (US, UK, and EOS)
            let bankRow = bankDetailsStartRow;
            // Only write rows that have data to avoid blank lines
            const bankLabelStyle = { font: { bold: true, size: 11, name: 'Calibri' } };
            const bankValueStyle = { font: { size: 11, name: 'Calibri' } };
            // Merge A:D and write rich text "Label: value" (no bold)
            const writeBankLine = (row: number, label: string, value: string) => {
                worksheet.mergeCells(`A${row}:D${row}`);
                const cell = worksheet.getCell(`A${row}`);
                cell.value = {
                    richText: [
                        { text: `${label}: `, font: { bold: false, size: 11, name: 'Calibri' } },
                        { text: `${value || ''}`, font: { size: 11, name: 'Calibri', bold: false } }
                    ]
                } as any;
                cell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true } as any;
            };

            const standardBankDetails = [
                { label: 'Bank Name', value: this.invoiceData.bankName },
                { label: 'Bank Address', value: this.invoiceData.bankAddress },
                { label: 'IBAN', value: this.invoiceData.iban },
                { label: 'Swift Code', value: this.invoiceData.swiftCode },
                { label: 'Title on Account', value: this.invoiceData.accountTitle }
            ];
            standardBankDetails.forEach(detail => {
                if (detail.value && detail.value.trim()) {
                    writeBankLine(bankRow, detail.label, detail.value);
                    bankRow++;
                }
            });

            // Conditional bank extras based on company
            if (this.selectedBank === 'US') {
                if (this.invoiceData.achRouting && this.invoiceData.achRouting.trim()) {
                    writeBankLine(bankRow, 'ACH Routing', this.invoiceData.achRouting);
                    bankRow++;
                }
            } else if (this.selectedBank === 'UK') {
                if (this.invoiceData.accountNumber && this.invoiceData.accountNumber.trim()) {
                    writeBankLine(bankRow, 'Account Number', this.invoiceData.accountNumber);
                    bankRow++;
                }
                if (this.invoiceData.sortCode && this.invoiceData.sortCode.trim()) {
                    writeBankLine(bankRow, 'Sort Code', this.invoiceData.sortCode);
                    bankRow++;
                }

                // UK DOMESTIC WIRES section
                bankRow++; // Add space above
                worksheet.mergeCells(`A${bankRow}:D${bankRow}`);
                const ukDomesticHeader = worksheet.getCell(`A${bankRow}`);
                ukDomesticHeader.value = 'UK DOMESTIC WIRES:';
                ukDomesticHeader.font = { bold: false, size: 11, name: 'Calibri' };
                ukDomesticHeader.alignment = { horizontal: 'left', vertical: 'middle', wrapText: false } as any;
                bankRow++;

                if (this.invoiceData.accountNumber && this.invoiceData.accountNumber.trim()) {
                    writeBankLine(bankRow, 'Account number', this.invoiceData.accountNumber);
                    bankRow++;
                }
                if (this.invoiceData.sortCode && this.invoiceData.sortCode.trim()) {
                    writeBankLine(bankRow, 'Sort code', this.invoiceData.sortCode);
                    bankRow++;
                }
            } else if (this.selectedBank === 'EOS') {
                if (this.invoiceData.intermediaryBic && this.invoiceData.intermediaryBic.trim()) {
                    writeBankLine(bankRow, 'Intermediary BIC', this.invoiceData.intermediaryBic);
                    bankRow++;
                }
            }

            // Invoice Details Section (Right side - Row 15-21)
            // Only write rows that have data to avoid blank lines
            const invoiceLabelStyle = { font: { bold: true, size: 11, name: 'Calibri' } };
            const invoiceValueStyle = { font: { size: 11, name: 'Calibri' } };
            // Helper to format date as "November 03, 2025"
            const formatDateAsText = (dateString: string): string => {
                const date = new Date(dateString);
                const months = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
                const month = months[date.getMonth()];
                const day = date.getDate().toString().padStart(2, '0');
                const year = date.getFullYear();
                return `${month} ${day}, ${year}`;
            };

            // Helper to write invoice detail with label and value in separate cells
            const writeInvoiceDetail = (row: number, label: string, value: string, isDate: boolean = false) => {
                // Write label in column E with colon
                const labelCell = worksheet.getCell(`E${row}`);
                labelCell.value = `${label}:`;
                labelCell.font = { size: 11, name: 'Calibri', bold: false };
                labelCell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: false } as any;

                // Leave column F empty
                worksheet.getCell(`F${row}`).value = null as any;

                // Write value in column G
                const valueCell = worksheet.getCell(`G${row}`);
                // Format date as text if needed
                if (isDate && value) {
                    valueCell.value = formatDateAsText(value);
                } else {
                    valueCell.value = value || '';
                }
                valueCell.font = { size: 11, name: 'Calibri', bold: false };
                valueCell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: false } as any;
            };

            const categoryToUse = categoryOverride || this.invoiceData.category;
            // Append "A" to invoice number if requested (for provisions when both categories exist)
            const invoiceNumberToUse = appendAtoInvoiceNumber ? `${this.invoiceData.invoiceNumber}A` : this.invoiceData.invoiceNumber;
            let invoiceRow = bankDetailsStartRow;
            const invoiceDetails = [
                { label: 'No', value: invoiceNumberToUse, isDate: false },
                { label: 'Invoice Date', value: this.invoiceData.invoiceDate, isDate: true },
                { label: 'Vessel', value: this.invoiceData.vessel, isDate: false },
                { label: 'Country', value: this.invoiceData.country, isDate: false },
                { label: 'Port', value: this.invoiceData.port, isDate: false },
                { label: 'Category', value: categoryToUse, isDate: false },
                { label: 'Invoice Due', value: this.invoiceData.invoiceDue, isDate: false }
            ];
            invoiceDetails.forEach(detail => {
                writeInvoiceDetail(invoiceRow, detail.label, detail.value || '', detail.isDate);
                invoiceRow++;
            });

            // Calculate the end of bank/invoice details section and start table right after
            // Use the maximum row from bank or invoice details, then add 1 row gap
            // Move table down one additional row
            const middleSectionEndRow = Math.max(bankRow, invoiceRow);
            const tableStartRow = middleSectionEndRow + 2;

            // Table Headers
            const headers = ['Pos', 'Description', 'Remark', 'Unit', 'Qty', 'Price', 'Total'];
            headers.forEach((header, index) => {
                const cell = worksheet.getCell(tableStartRow, index + 1);
                cell.value = header;
                cell.font = { bold: true, size: 11, name: 'Calibri', color: { argb: 'FFFFFFFF' } };
                const headerFillColor = this.selectedBank === 'EOS' ? 'FF0B2E66' : 'FF808080';
                cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: headerFillColor } } as any;
                cell.alignment = { horizontal: 'center', vertical: 'middle' };
            });

            // Calculate totals for these items
            const itemsSubtotal = items.reduce((sum, item) => sum + (item.total || 0), 0);
            const discountAmount = itemsSubtotal * (this.invoiceData.discountPercent || 0) / 100;
            const feesTotal = includeFees ?
                ((this.invoiceData.deliveryFee || 0) +
                    (this.invoiceData.portFee || 0) +
                    (this.invoiceData.agencyFee || 0) +
                    (this.invoiceData.transportCustomsLaunchFees || 0) +
                    (this.invoiceData.launchFee || 0)) : 0;
            const categoryTotal = (itemsSubtotal - discountAmount) + feesTotal;

            // Table Data
            items.forEach((item, index) => {
                const rowIndex = tableStartRow + 1 + index;

                // Pos (renumber from 1 for this file)
                const posCell = worksheet.getCell(rowIndex, 1);
                posCell.value = index + 1;
                posCell.font = { size: 10, name: 'Calibri' };
                posCell.alignment = { horizontal: 'center', vertical: 'middle' };

                // Description
                const descCell = worksheet.getCell(rowIndex, 2);
                descCell.value = item.description;
                descCell.font = { size: 10, name: 'Calibri' };
                descCell.alignment = { horizontal: 'left', vertical: 'middle' };

                // Remark
                const remarkCell = worksheet.getCell(rowIndex, 3);
                remarkCell.value = item.remark;
                remarkCell.font = { size: 10, name: 'Calibri' };
                remarkCell.alignment = { horizontal: 'left', vertical: 'middle' };

                // Unit
                const unitCell = worksheet.getCell(rowIndex, 4);
                unitCell.value = item.unit;
                unitCell.font = { size: 10, name: 'Calibri' };
                unitCell.alignment = { horizontal: 'center', vertical: 'middle' };

                // Qty (default to 0)
                const qtyCell = worksheet.getCell(rowIndex, 5);
                qtyCell.value = (item.qty ?? 0);
                qtyCell.font = { size: 10, name: 'Calibri' };
                qtyCell.alignment = { horizontal: 'center', vertical: 'middle' };

                // Price (rounded to nearest penny)
                const priceCell = worksheet.getCell(rowIndex, 6);
                priceCell.value = Math.round(item.price * 100) / 100;
                priceCell.font = { size: 10, name: 'Calibri' };
                priceCell.alignment = { horizontal: 'right', vertical: 'middle' };
                // Use dynamic currency formatting based on item currency
                const currencyFormat = this.getCurrencyExcelFormat(item.currency);
                priceCell.numFmt = currencyFormat;

                // Total (formula = E * F)
                const totalCell = worksheet.getCell(rowIndex, 7);
                totalCell.value = { formula: `E${rowIndex}*F${rowIndex}` } as any;
                totalCell.font = { size: 10, name: 'Calibri' };
                totalCell.alignment = { horizontal: 'right', vertical: 'middle' };
                totalCell.numFmt = currencyFormat;

                // Add grid lines to data cells
                posCell.border = { top: { style: 'thin' }, bottom: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } } as any;
                descCell.border = { top: { style: 'thin' }, bottom: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } } as any;
                remarkCell.border = { top: { style: 'thin' }, bottom: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } } as any;
                unitCell.border = { top: { style: 'thin' }, bottom: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } } as any;
                qtyCell.border = { top: { style: 'thin' }, bottom: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } } as any;
                priceCell.border = { top: { style: 'thin' }, bottom: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } } as any;
                totalCell.border = { top: { style: 'thin' }, bottom: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } } as any;
            });

            // Get currency format for totals
            const primaryCurrencyFormat = this.getCurrencyExcelFormat(this.primaryCurrency);

            // Add subtotal row for Total column at the bottom of datatable - moved down 1 row
            const subtotalRow = tableStartRow + items.length + 2;
            const firstDataRow = tableStartRow + 1;
            const lastDataRow = tableStartRow + items.length;

            // Add label in column F
            const subtotalLabelCell = worksheet.getCell(`F${subtotalRow}`);
            subtotalLabelCell.value = `TOTAL ${this.getCurrencyLabel(this.primaryCurrency)}`;
            subtotalLabelCell.font = { size: 11, name: 'Calibri', bold: true };
            subtotalLabelCell.alignment = { horizontal: 'right', vertical: 'middle' };

            // Add formula in column G (no border, no shading)
            const subtotalCell = worksheet.getCell(`G${subtotalRow}`);
            subtotalCell.value = { formula: `SUM(G${firstDataRow}:G${lastDataRow})` } as any;
            subtotalCell.font = { size: 11, name: 'Calibri', bold: true };
            subtotalCell.alignment = { horizontal: 'right', vertical: 'middle' };
            subtotalCell.numFmt = primaryCurrencyFormat;

            // Totals and Fees Section
            let totalsStartRow = tableStartRow + items.length + 4; // moved down by 2 rows now

            // List discount (amount) and non-zero fees just above the Total
            // Discount is always included (applied to items), fees are conditional
            const feeLines: { label: string; value?: number; includeInSum?: boolean }[] = [];
            if (discountAmount > 0) feeLines.push({ label: 'Discount:', value: -discountAmount, includeInSum: false });

            // Only include fees if includeFees is true
            if (includeFees) {
                if (this.invoiceData.deliveryFee) feeLines.push({ label: 'Delivery fee:', value: this.invoiceData.deliveryFee, includeInSum: true });
                if (this.invoiceData.portFee) feeLines.push({ label: 'Port fee:', value: this.invoiceData.portFee, includeInSum: true });
                if (this.invoiceData.agencyFee) feeLines.push({ label: 'Agency fee:', value: this.invoiceData.agencyFee, includeInSum: true });
                if (this.invoiceData.transportCustomsLaunchFees) feeLines.push({ label: 'Transport, Customs, Launch fees:', value: this.invoiceData.transportCustomsLaunchFees, includeInSum: true });
                if (this.invoiceData.launchFee) feeLines.push({ label: 'Launch:', value: this.invoiceData.launchFee, includeInSum: true });
            }

            // Get currency format for totals (already defined above)
            const totalLabel = `TOTAL ${this.getCurrencyLabel(this.primaryCurrency)}`;

            // Track fee rows that contain numeric amounts for later SUM
            const feeAmountRowRefs: string[] = [];
            feeLines.forEach((fee, idx) => {
                const rowIndex = totalsStartRow + idx;
                const labelCell = worksheet.getCell(`F${rowIndex}`);
                labelCell.value = fee.label;
                labelCell.font = { bold: true, size: 11, name: 'Calibri' };
                labelCell.alignment = { horizontal: 'right', vertical: 'middle' };

                const valueCell = worksheet.getCell(`G${rowIndex}`);
                valueCell.value = fee.value as number;
                valueCell.numFmt = primaryCurrencyFormat;
                if (fee.includeInSum) feeAmountRowRefs.push(`G${rowIndex}`);
                valueCell.font = { bold: true, size: 11, name: 'Calibri' };
                valueCell.alignment = { horizontal: 'right', vertical: 'middle' };

                // Use standard line height like the rest of the page
                worksheet.getRow(rowIndex).height = 15; // typical default
            });

            totalsStartRow += feeLines.length;

            // Final TOTAL (formula: sum of G column item totals + all monetary fee cells) - no label
            worksheet.getCell(`F${totalsStartRow}`).value = '';
            worksheet.getCell(`F${totalsStartRow}`).alignment = { horizontal: 'right', vertical: 'middle' };

            // firstDataRow and lastDataRow already defined above
            const itemsSumFormula = `SUM(G${firstDataRow}:G${lastDataRow})`;
            const feeSumPart = feeAmountRowRefs.length ? `+${feeAmountRowRefs.join('+')}` : '';
            const discountFactor = this.invoiceData.discountPercent ? `(1-${this.invoiceData.discountPercent}/100)` : '1';
            const totalFormula = `(${itemsSumFormula}*${discountFactor})${feeSumPart}`;
            worksheet.getCell(`G${totalsStartRow}`).value = { formula: totalFormula } as any;
            worksheet.getCell(`G${totalsStartRow}`).font = { bold: true, size: 11, name: 'Calibri' };
            worksheet.getCell(`G${totalsStartRow}`).alignment = { horizontal: 'right', vertical: 'middle' };
            worksheet.getCell(`G${totalsStartRow}`).numFmt = primaryCurrencyFormat;

            // Removed separate grand total line; the total row represents the final amount

            // Terms and Conditions
            // For EOS, reduce gap from 4 rows to 2 rows (remove 2 empty rows above terms)
            const termsStartRow = this.selectedBank === 'EOS' ? totalsStartRow + 2 : totalsStartRow + 4;
            worksheet.mergeCells(`A${termsStartRow}:G${termsStartRow}`);
            const termsHeader = worksheet.getCell(`A${termsStartRow}`);
            termsHeader.value = 'By placing the order according to the above quotation you are accepting the following terms:';
            termsHeader.font = { bold: true, size: 11, name: 'Calibri' };
            termsHeader.alignment = { horizontal: 'left', vertical: 'middle', wrapText: false } as any;

            const terms = [
                { roman: 'I.', text: 'Credit days: 30 calendar days.' },
                { roman: 'II.', text: 'Accounts not paid in this time frame will be charged 10% interest rate per month, any discount given will be null and void.' },
                { roman: 'III.', text: 'Should collection or legal action be required to collect past dues, fees for such action will be added to your account.' },
                { roman: 'IV.', text: 'Subject to unsold. Final weights are subject to vendor packing standards.' },
                { roman: 'V.', text: 'If the transaction is canceled after the order is authorized, we have the right to collect the invoice without claim.' }
            ];

            terms.forEach((term, index) => {
                const row = termsStartRow + 1 + index;
                // Roman numeral in column A
                const romanCell = worksheet.getCell(`A${row}`);
                romanCell.value = term.roman;
                romanCell.font = { size: 10, name: 'Calibri' };
                romanCell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: false } as any;

                // Text in column B
                const textCell = worksheet.getCell(`B${row}`);
                textCell.value = term.text;
                textCell.font = { size: 10, name: 'Calibri' };
                textCell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: false } as any;
            });

            // Print company and bank details at bottom (left and right blocks) - only for EOS
            let maxBottomOffset = 0;
            if (this.selectedBank === 'EOS') {
                const fontColor = this.selectedBank === 'EOS' ? 'FF0B2E66' : 'FF000000';
                const bottomFont = { name: 'Calibri', size: 11, bold: true, color: { argb: fontColor } } as any;
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
                    cell.font = bottomFont;
                    cell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: false } as any;
                });

                // Right block: Bank Details starting at column D; merge to allow long text on one line
                const rightStartRow = bottomSectionStart;
                const writeRight = (offset: number, text: string) => {
                    const row = rightStartRow + offset;
                    worksheet.mergeCells(`D${row}:G${row}`);
                    const cell = worksheet.getCell(`D${row}`);
                    cell.value = text;
                    cell.font = bottomFont;
                    cell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: false } as any;
                };

                // Add "Bank Details" header (underlined)
                const headerRow = rightStartRow;
                worksheet.mergeCells(`D${headerRow}:G${headerRow}`);
                const headerCell = worksheet.getCell(`D${headerRow}`);
                headerCell.value = 'Bank Details';
                headerCell.font = { ...bottomFont, underline: true } as any;
                headerCell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: false } as any;

                // Bank details - standard fields
                let currentOffset = 1;
                writeRight(currentOffset++, this.invoiceData.bankName || '');
                writeRight(currentOffset++, this.invoiceData.bankAddress || '');

                // EOS bank details
                writeRight(currentOffset++, `IBAN: ${this.invoiceData.iban || ''}`);
                writeRight(currentOffset++, `SWIFTBIC: ${this.invoiceData.swiftCode || ''}`);

                // Track the maximum offset used (currentOffset is already incremented past the last write)
                maxBottomOffset = currentOffset - 1; // Subtract 1 because currentOffset was incremented after the last write
            }

            // Add bottom image
            // Calculate the last row of text content
            // maxBottomOffset is only calculated for EOS now
            const lastTextRow = (this.selectedBank === 'EOS')
                ? (termsStartRow + terms.length + 4 + maxBottomOffset)  // bottomSectionStart + maxBottomOffset
                : (termsStartRow + terms.length);  // Last term row
            const bottomImageRowPosition = lastTextRow + 2; // Leave exactly 2 blank rows
            if ((worksheet as any)._bottomImageId) {
                const originalWidth = (worksheet as any)._bottomImageWidth || 667;
                const originalHeight = (worksheet as any)._bottomImageHeight || 80;
                // Set image width to 1000 pixels, maintain aspect ratio
                const newWidth = 1000;
                const newHeight = (originalHeight * newWidth) / originalWidth; // Maintain aspect ratio
                worksheet.addImage((worksheet as any)._bottomImageId, {
                    tl: { col: 0, row: bottomImageRowPosition }, // Position at bottom
                    ext: { width: newWidth, height: newHeight }
                });
            }

            // Set print area: columns A-G starting from row 1, ending 3 rows below the bottom image
            // bottomImageRowPosition is 0-based, so add 1 to convert to 1-based, then add 3 more rows
            const printAreaEndRow = bottomImageRowPosition + 1 + 3;
            worksheet.pageSetup.printArea = `A1:G${printAreaEndRow}`;

            // Set page setup to fit to page width (1 page wide), allow multiple pages for height
            worksheet.pageSetup.fitToPage = true;
            worksheet.pageSetup.fitToWidth = 1;
            // Don't set fitToHeight to allow multiple pages for height

            // Generate Excel file
            const buffer = await workbook.xlsx.writeBuffer();
            const blob = new Blob([buffer], {
                type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            });

            // Download file
            let fileName: string;
            if (this.invoiceData.exportFileName && this.invoiceData.exportFileName.trim()) {
                fileName = this.invoiceData.exportFileName.trim();
                // Replace <Category> placeholder with actual category if categoryOverride is provided
                if (categoryOverride && fileName.includes('<Category>')) {
                    fileName = fileName.replace('<Category>', categoryOverride);
                }
                if (!fileName.endsWith('.xlsx')) {
                    fileName = `${fileName}.xlsx`;
                }
            } else {
                const filePrefix = this.selectedBank === 'EOS' ? 'EOS_Invoice' : 'HIMarine_Invoice';
                const categorySuffix = categoryOverride ? `_${categoryOverride}` : '';
                fileName = `${filePrefix}_${this.invoiceData.invoiceNumber}${categorySuffix}_${new Date().toISOString().split('T')[0]}.xlsx`;
            }
            saveAs(blob, fileName);

            this.loggingService.logExport('excel_invoice_exported', {
                fileName,
                fileSize: blob.size,
                itemsIncluded: items.length,
                totalAmount: categoryTotal,
                category: categoryToUse
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

