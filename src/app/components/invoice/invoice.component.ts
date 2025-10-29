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
        deliveryFee: 225.00,
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

        // Vessel Details
        this.invoiceData.vesselName = 'Ludogorets Maritime Ltd., Marshall Islands';
        this.invoiceData.vesselName2 = 'c/o Navigation Maritime Bulgare';
        this.invoiceData.vesselAddress = '1 Primorski Blvd,';
        this.invoiceData.vesselAddress2 = '9000 Varna,';
        this.invoiceData.vesselCity = '';
        this.invoiceData.vesselCountry = 'Bulgaria';

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

        // Vessel Details
        this.invoiceData.vesselName = 'TOUGH JENS MARITIME S.A.,';
        this.invoiceData.vesselName2 = 'M/V ZALIV';
        this.invoiceData.vesselAddress = 'VIA ESPANA 122, DELTA TOWER FL 14';
        this.invoiceData.vesselAddress2 = 'DILIGIANNI TH. 9, KIFISSIA, ATHENS';
        this.invoiceData.vesselCity = '';
        this.invoiceData.vesselCountry = 'Greece';

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

        // Vessel Details
        this.invoiceData.vesselName = 'GLOBAL PACIFIC SHIPPING JOINT STOCK COMPANY';
        this.invoiceData.vesselName2 = '';
        this.invoiceData.vesselAddress = '10th Floor, Tower 1 of the Room Commercial Service -';
        this.invoiceData.vesselAddress2 = 'Hotel area project (The Nexus), 3A-3B Ton Duc Thang Street,';
        this.invoiceData.vesselCity = 'Sai Gon Ward, Ho Chi Minh City';
        this.invoiceData.vesselCountry = 'Vietnam';

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

        this.calculateTotals();
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

        this.calculateTotals();
    }

    private calculateTotals(): void {
        this.invoiceData.totalGBP = this.invoiceData.items.reduce((sum, item) => sum + item.total, 0);
        this.invoiceData.grandTotal = this.invoiceData.totalGBP + this.invoiceData.deliveryFee;
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

            // Load and add images
            try {
                // Add top logo image
                const topImageResponse = await fetch('assets/images/HIMarineTopImage.png');
                const topImageBuffer = await topImageResponse.arrayBuffer();
                const topImageId = workbook.addImage({
                    buffer: topImageBuffer,
                    extension: 'png',
                });
                worksheet.addImage(topImageId, {
                    tl: { col: 1.5, row: 0.5 }, // Moved down slightly from the very top
                    ext: { width: 300, height: 180 } // 150% of current size: 300x180 pixels
                });

                // Add bottom border image at the end (we'll set this position later after we know the final row)
                const bottomImageResponse = await fetch('assets/images/HIMarineBottomBorder.png');
                const bottomImageBuffer = await bottomImageResponse.arrayBuffer();
                const bottomImageId = workbook.addImage({
                    buffer: bottomImageBuffer,
                    extension: 'png',
                });

                // We'll add the bottom image after calculating all content
                // Store the ID for later use
                (worksheet as any)._bottomImageId = bottomImageId;
            } catch (imageError) {
                console.warn('Could not load images for Excel export:', imageError);
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

            // Company Header - Left aligned in column A starting at row 9 (ALL BOLD)
            const companyHeader = worksheet.getCell('A9');
            companyHeader.value = 'HI MARINE COMPANY LIMITED';
            companyHeader.font = { bold: true, size: 18, name: 'Arial' };
            companyHeader.alignment = { horizontal: 'left' };

            // Company Address - Left aligned (ALL BOLD)
            const address1 = worksheet.getCell('A10');
            address1.value = 'Wearfield, Enterprise Park East,';
            address1.font = { bold: true, size: 12, name: 'Arial' };
            address1.alignment = { horizontal: 'left' };

            const address2 = worksheet.getCell('A11');
            address2.value = 'Sunderland, Tyne and Wear, SR5 2TA';
            address2.font = { bold: true, size: 12, name: 'Arial' };
            address2.alignment = { horizontal: 'left' };

            const country = worksheet.getCell('A12');
            country.value = 'United Kingdom';
            country.font = { bold: true, size: 12, name: 'Arial' };
            country.alignment = { horizontal: 'left' };

            const email = worksheet.getCell('A13');
            email.value = 'office@himarinecompany.com';
            email.font = { bold: true, size: 12, name: 'Arial' };
            email.alignment = { horizontal: 'left' };

            // Bank Details Section (Left side - Row 15-23)
            const bankDetailsStart = 15;
            const bankLabelStyle = { font: { bold: true, size: 11, name: 'Arial' } };
            const bankValueStyle = { font: { size: 11, name: 'Arial' } };

            worksheet.getCell(`A${bankDetailsStart}`).value = 'Bank Name:';
            worksheet.getCell(`A${bankDetailsStart}`).font = bankLabelStyle.font;
            worksheet.getCell(`B${bankDetailsStart}`).value = this.invoiceData.bankName;
            worksheet.getCell(`B${bankDetailsStart}`).font = bankValueStyle.font;

            worksheet.getCell(`A${bankDetailsStart + 1}`).value = 'Bank Address:';
            worksheet.getCell(`A${bankDetailsStart + 1}`).font = bankLabelStyle.font;
            worksheet.mergeCells(`B${bankDetailsStart + 1}:D${bankDetailsStart + 1}`);
            worksheet.getCell(`B${bankDetailsStart + 1}`).value = this.invoiceData.bankAddress;
            worksheet.getCell(`B${bankDetailsStart + 1}`).font = bankValueStyle.font;

            worksheet.getCell(`A${bankDetailsStart + 2}`).value = 'IBAN:';
            worksheet.getCell(`A${bankDetailsStart + 2}`).font = bankLabelStyle.font;
            worksheet.mergeCells(`B${bankDetailsStart + 2}:D${bankDetailsStart + 2}`);
            worksheet.getCell(`B${bankDetailsStart + 2}`).value = this.invoiceData.iban;
            worksheet.getCell(`B${bankDetailsStart + 2}`).font = bankValueStyle.font;

            worksheet.getCell(`A${bankDetailsStart + 3}`).value = 'Swift Code:';
            worksheet.getCell(`A${bankDetailsStart + 3}`).font = bankLabelStyle.font;
            worksheet.getCell(`B${bankDetailsStart + 3}`).value = this.invoiceData.swiftCode;
            worksheet.getCell(`B${bankDetailsStart + 3}`).font = bankValueStyle.font;

            worksheet.getCell(`A${bankDetailsStart + 4}`).value = 'Title on Account:';
            worksheet.getCell(`A${bankDetailsStart + 4}`).font = bankLabelStyle.font;
            worksheet.mergeCells(`B${bankDetailsStart + 4}:D${bankDetailsStart + 4}`);
            worksheet.getCell(`B${bankDetailsStart + 4}`).value = this.invoiceData.accountTitle;
            worksheet.getCell(`B${bankDetailsStart + 4}`).font = bankValueStyle.font;

            worksheet.getCell(`A${bankDetailsStart + 6}`).value = 'UK DOMESTIC WIRES:';
            worksheet.getCell(`A${bankDetailsStart + 6}`).font = { bold: true, size: 11, name: 'Arial' };

            worksheet.getCell(`A${bankDetailsStart + 7}`).value = 'Account number:';
            worksheet.getCell(`A${bankDetailsStart + 7}`).font = bankLabelStyle.font;
            worksheet.getCell(`B${bankDetailsStart + 7}`).value = this.invoiceData.accountNumber;
            worksheet.getCell(`B${bankDetailsStart + 7}`).font = bankValueStyle.font;

            worksheet.getCell(`A${bankDetailsStart + 8}`).value = 'Sort code:';
            worksheet.getCell(`A${bankDetailsStart + 8}`).font = bankLabelStyle.font;
            worksheet.getCell(`B${bankDetailsStart + 8}`).value = this.invoiceData.sortCode;
            worksheet.getCell(`B${bankDetailsStart + 8}`).font = bankValueStyle.font;

            // Invoice Details Section (Right side - Row 15-21)
            const invoiceDetailsStart = 15;
            const invoiceLabelStyle = { font: { bold: true, size: 11, name: 'Arial' } };
            const invoiceValueStyle = { font: { size: 11, name: 'Arial' } };

            worksheet.getCell(`E${invoiceDetailsStart}`).value = 'No';
            worksheet.getCell(`E${invoiceDetailsStart}`).font = invoiceLabelStyle.font;
            worksheet.getCell(`F${invoiceDetailsStart}`).value = this.invoiceData.invoiceNumber;
            worksheet.getCell(`F${invoiceDetailsStart}`).font = invoiceValueStyle.font;

            worksheet.getCell(`E${invoiceDetailsStart + 1}`).value = 'Invoice Date';
            worksheet.getCell(`E${invoiceDetailsStart + 1}`).font = invoiceLabelStyle.font;
            worksheet.mergeCells(`F${invoiceDetailsStart + 1}:G${invoiceDetailsStart + 1}`);
            worksheet.getCell(`F${invoiceDetailsStart + 1}`).value = this.invoiceData.invoiceDate;
            worksheet.getCell(`F${invoiceDetailsStart + 1}`).font = invoiceValueStyle.font;

            worksheet.getCell(`E${invoiceDetailsStart + 2}`).value = 'Vessel';
            worksheet.getCell(`E${invoiceDetailsStart + 2}`).font = invoiceLabelStyle.font;
            worksheet.getCell(`F${invoiceDetailsStart + 2}`).value = this.invoiceData.vessel;
            worksheet.getCell(`F${invoiceDetailsStart + 2}`).font = invoiceValueStyle.font;

            worksheet.getCell(`E${invoiceDetailsStart + 3}`).value = 'Country';
            worksheet.getCell(`E${invoiceDetailsStart + 3}`).font = invoiceLabelStyle.font;
            worksheet.getCell(`F${invoiceDetailsStart + 3}`).value = this.invoiceData.country;
            worksheet.getCell(`F${invoiceDetailsStart + 3}`).font = invoiceValueStyle.font;

            worksheet.getCell(`E${invoiceDetailsStart + 4}`).value = 'Port';
            worksheet.getCell(`E${invoiceDetailsStart + 4}`).font = invoiceLabelStyle.font;
            worksheet.getCell(`F${invoiceDetailsStart + 4}`).value = this.invoiceData.port;
            worksheet.getCell(`F${invoiceDetailsStart + 4}`).font = invoiceValueStyle.font;

            worksheet.getCell(`E${invoiceDetailsStart + 5}`).value = 'Category';
            worksheet.getCell(`E${invoiceDetailsStart + 5}`).font = invoiceLabelStyle.font;
            worksheet.getCell(`F${invoiceDetailsStart + 5}`).value = this.invoiceData.category;
            worksheet.getCell(`F${invoiceDetailsStart + 5}`).font = invoiceValueStyle.font;

            worksheet.getCell(`E${invoiceDetailsStart + 6}`).value = 'Invoice Due';
            worksheet.getCell(`E${invoiceDetailsStart + 6}`).font = invoiceLabelStyle.font;
            worksheet.getCell(`F${invoiceDetailsStart + 6}`).value = this.invoiceData.invoiceDue;
            worksheet.getCell(`F${invoiceDetailsStart + 6}`).font = invoiceValueStyle.font;

            // Items Table (Starting from row 25)
            const tableStartRow = 25;

            // Table Headers
            const headers = ['Pos', 'Tab', 'Description', 'Remark', 'Unit', 'Qty', 'Price', 'Total'];
            headers.forEach((header, index) => {
                const cell = worksheet.getCell(tableStartRow, index + 1);
                cell.value = header;
                cell.font = { bold: true, size: 11, name: 'Arial', color: { argb: 'FFFFFFFF' } };
                cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF808080' } } as any;
                cell.alignment = { horizontal: 'center', vertical: 'middle' };
                cell.border = {
                    top: { style: 'thin', color: { argb: 'FF000000' } },
                    bottom: { style: 'thin', color: { argb: 'FF000000' } },
                    left: { style: 'thin', color: { argb: 'FF000000' } },
                    right: { style: 'thin', color: { argb: 'FF000000' } }
                } as any;
            });

            // Table Data
            this.invoiceData.items.forEach((item, index) => {
                const rowIndex = tableStartRow + 1 + index;

                // Pos
                const posCell = worksheet.getCell(rowIndex, 1);
                posCell.value = item.pos;
                posCell.font = { size: 10, name: 'Arial' };
                posCell.alignment = { horizontal: 'center', vertical: 'middle' };
                posCell.border = {
                    top: { style: 'thin', color: { argb: 'FF000000' } },
                    bottom: { style: 'thin', color: { argb: 'FF000000' } },
                    left: { style: 'thin', color: { argb: 'FF000000' } },
                    right: { style: 'thin', color: { argb: 'FF000000' } }
                } as any;

                // Tab
                const tabCell = worksheet.getCell(rowIndex, 2);
                tabCell.value = item.tabName;
                tabCell.font = { size: 10, name: 'Arial' };
                tabCell.alignment = { horizontal: 'center', vertical: 'middle' };
                tabCell.border = {
                    top: { style: 'thin', color: { argb: 'FF000000' } },
                    bottom: { style: 'thin', color: { argb: 'FF000000' } },
                    left: { style: 'thin', color: { argb: 'FF000000' } },
                    right: { style: 'thin', color: { argb: 'FF000000' } }
                } as any;

                // Description
                const descCell = worksheet.getCell(rowIndex, 3);
                descCell.value = item.description;
                descCell.font = { size: 10, name: 'Arial' };
                descCell.alignment = { horizontal: 'left', vertical: 'middle' };
                descCell.border = {
                    top: { style: 'thin', color: { argb: 'FF000000' } },
                    bottom: { style: 'thin', color: { argb: 'FF000000' } },
                    left: { style: 'thin', color: { argb: 'FF000000' } },
                    right: { style: 'thin', color: { argb: 'FF000000' } }
                } as any;

                // Remark
                const remarkCell = worksheet.getCell(rowIndex, 4);
                remarkCell.value = item.remark;
                remarkCell.font = { size: 10, name: 'Arial' };
                remarkCell.alignment = { horizontal: 'left', vertical: 'middle' };
                remarkCell.border = {
                    top: { style: 'thin', color: { argb: 'FF000000' } },
                    bottom: { style: 'thin', color: { argb: 'FF000000' } },
                    left: { style: 'thin', color: { argb: 'FF000000' } },
                    right: { style: 'thin', color: { argb: 'FF000000' } }
                } as any;

                // Unit
                const unitCell = worksheet.getCell(rowIndex, 5);
                unitCell.value = item.unit;
                unitCell.font = { size: 10, name: 'Arial' };
                unitCell.alignment = { horizontal: 'center', vertical: 'middle' };
                unitCell.border = {
                    top: { style: 'thin', color: { argb: 'FF000000' } },
                    bottom: { style: 'thin', color: { argb: 'FF000000' } },
                    left: { style: 'thin', color: { argb: 'FF000000' } },
                    right: { style: 'thin', color: { argb: 'FF000000' } }
                } as any;

                // Qty
                const qtyCell = worksheet.getCell(rowIndex, 6);
                qtyCell.value = item.qty;
                qtyCell.font = { size: 10, name: 'Arial' };
                qtyCell.alignment = { horizontal: 'center', vertical: 'middle' };
                qtyCell.border = {
                    top: { style: 'thin', color: { argb: 'FF000000' } },
                    bottom: { style: 'thin', color: { argb: 'FF000000' } },
                    left: { style: 'thin', color: { argb: 'FF000000' } },
                    right: { style: 'thin', color: { argb: 'FF000000' } }
                } as any;

                // Price
                const priceCell = worksheet.getCell(rowIndex, 7);
                priceCell.value = item.price;
                priceCell.font = { size: 10, name: 'Arial' };
                priceCell.alignment = { horizontal: 'right', vertical: 'middle' };
                // Use dynamic currency formatting based on item currency
                const currencyFormat = item.currency === '$' ? '$#,##0.00' :
                    item.currency === '€' ? '€#,##0.00' :
                        '£#,##0.00';
                priceCell.numFmt = currencyFormat;
                priceCell.border = {
                    top: { style: 'thin', color: { argb: 'FF000000' } },
                    bottom: { style: 'thin', color: { argb: 'FF000000' } },
                    left: { style: 'thin', color: { argb: 'FF000000' } },
                    right: { style: 'thin', color: { argb: 'FF000000' } }
                } as any;

                // Total
                const totalCell = worksheet.getCell(rowIndex, 8);
                totalCell.value = item.total;
                totalCell.font = { size: 10, name: 'Arial' };
                totalCell.alignment = { horizontal: 'right', vertical: 'middle' };
                totalCell.numFmt = currencyFormat;
                totalCell.border = {
                    top: { style: 'thin', color: { argb: 'FF000000' } },
                    bottom: { style: 'thin', color: { argb: 'FF000000' } },
                    left: { style: 'thin', color: { argb: 'FF000000' } },
                    right: { style: 'thin', color: { argb: 'FF000000' } }
                } as any;
            });

            // Totals Section
            const totalsStartRow = tableStartRow + this.invoiceData.items.length + 3;

            // TOTAL GBP
            worksheet.getCell(`F${totalsStartRow}`).value = 'TOTAL GBP';
            worksheet.getCell(`F${totalsStartRow}`).font = { bold: true, size: 11, name: 'Arial' };
            worksheet.getCell(`F${totalsStartRow}`).alignment = { horizontal: 'right', vertical: 'middle' };
            worksheet.getCell(`G${totalsStartRow}`).value = this.invoiceData.totalGBP;
            worksheet.getCell(`G${totalsStartRow}`).font = { bold: true, size: 11, name: 'Arial' };
            worksheet.getCell(`G${totalsStartRow}`).alignment = { horizontal: 'right', vertical: 'middle' };
            worksheet.getCell(`G${totalsStartRow}`).numFmt = '£#,##0.00';

            // Delivery fee
            worksheet.getCell(`F${totalsStartRow + 1}`).value = 'Delivery fee';
            worksheet.getCell(`F${totalsStartRow + 1}`).font = { bold: true, size: 11, name: 'Arial' };
            worksheet.getCell(`F${totalsStartRow + 1}`).alignment = { horizontal: 'right', vertical: 'middle' };
            worksheet.getCell(`G${totalsStartRow + 1}`).value = this.invoiceData.deliveryFee;
            worksheet.getCell(`G${totalsStartRow + 1}`).font = { bold: true, size: 11, name: 'Arial' };
            worksheet.getCell(`G${totalsStartRow + 1}`).alignment = { horizontal: 'right', vertical: 'middle' };
            worksheet.getCell(`G${totalsStartRow + 1}`).numFmt = '£#,##0.00';

            // Grand Total
            worksheet.getCell(`G${totalsStartRow + 2}`).value = this.invoiceData.grandTotal;
            worksheet.getCell(`G${totalsStartRow + 2}`).font = { bold: true, size: 12, name: 'Arial' };
            worksheet.getCell(`G${totalsStartRow + 2}`).alignment = { horizontal: 'right', vertical: 'middle' };
            worksheet.getCell(`G${totalsStartRow + 2}`).numFmt = '£#,##0.00';

            // Terms and Conditions
            const termsStartRow = totalsStartRow + 5;
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

            // Add bottom image
            const bottomImageRowPosition = termsStartRow + terms.length + 3; // Add some spacing
            if ((worksheet as any)._bottomImageId) {
                worksheet.addImage((worksheet as any)._bottomImageId, {
                    tl: { col: 0, row: bottomImageRowPosition }, // Position at bottom
                    ext: { width: 667, height: 80 } // 2/3 of previous width: 667x80 pixels
                });
            }

            // Generate Excel file
            const buffer = await workbook.xlsx.writeBuffer();
            const blob = new Blob([buffer], {
                type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            });

            // Download file
            const fileName = `HIMarine_Invoice_${this.invoiceData.invoiceNumber}_${new Date().toISOString().split('T')[0]}.xlsx`;
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

        // Totals
        yPos += 10;
        doc.line(margin, yPos, pageWidth - margin, yPos);
        yPos += 10;

        doc.setFont('helvetica', 'bold');
        doc.text(`TOTAL GBP: £${this.invoiceData.totalGBP.toFixed(2)}`, pageWidth - 80, yPos);
        yPos += 8;
        doc.text(`Delivery fee: £${this.invoiceData.deliveryFee.toFixed(2)}`, pageWidth - 80, yPos);
        yPos += 8;
        doc.setFontSize(12);
        doc.text(`GRAND TOTAL: £${this.invoiceData.grandTotal.toFixed(2)}`, pageWidth - 90, yPos);

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

