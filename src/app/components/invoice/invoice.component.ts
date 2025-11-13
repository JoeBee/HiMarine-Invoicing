import { Component, OnInit } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { DataService, ProcessedDataRow, ExcelProcessedData, ExcelItemData } from '../../services/data.service';
import { LoggingService } from '../../services/logging.service';
import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';
import jsPDF from 'jspdf';
import { buildInvoiceStyleWorkbook, InvoiceWorkbookOptions } from '../../utils/invoice-workbook-builder';
import { COUNTRIES, COUNTRY_PORTS } from '../../constants/countries.constants';

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
    countries = COUNTRIES;

    // Available ports for selected country
    availablePorts: string[] = [];

    // Country to ports mapping
    countryPorts: { [key: string]: string[] } = COUNTRY_PORTS;

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
        // Billing Information
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

    onCompanySelectionChange(): void {
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

        // Clear Billing Information
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

        // Billing Information (do not auto-populate; keep blank by default)
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

        // Billing Information (do not auto-populate; keep blank by default)
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

        // Billing Information (do not auto-populate; keep blank by default)
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
                const description = row.description ? String(row.description).toUpperCase() : '';
                const remarks = row.remarks ? String(row.remarks).toUpperCase() : '';
                return [
                    row.fileName,
                    description,
                    row.count,
                    remarks,
                    adjustedPrice,
                    row.count * adjustedPrice
                ];
            }),
            ['', '', '', '', 'TOTAL AMOUNT:', this.finalTotalAmount]
        ];

        // Create workbook and worksheet
        const worksheet = XLSX.utils.aoa_to_sheet(worksheetData);

        const headerAddresses = ['A1', 'B1', 'C1', 'D1', 'E1', 'F1'];
        headerAddresses.forEach(address => {
            const cell = worksheet[address] as any;
            if (!cell) {
                return;
            }
            const baseStyle = cell.s ? { ...cell.s } : {};
            const baseFont = baseStyle.font ? { ...baseStyle.font } : {};
            cell.s = {
                ...baseStyle,
                font: {
                    ...baseFont,
                    bold: true,
                    color: { rgb: 'FFFFFFFF' }
                },
                fill: {
                    patternType: 'solid',
                    fgColor: { rgb: 'FF4472C4' }
                },
                alignment: {
                    horizontal: 'center',
                    vertical: 'center'
                }
            };
        });

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

        this.applyCambriaFontToSheet(worksheet);

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

    private applyCambriaFontToSheet(worksheet: XLSX.WorkSheet): void {
        const reference = worksheet['!ref'];
        if (!reference) {
            return;
        }

        const range = XLSX.utils.decode_range(reference);
        for (let row = range.s.r; row <= range.e.r; row++) {
            for (let col = range.s.c; col <= range.e.c; col++) {
                const address = XLSX.utils.encode_cell({ r: row, c: col });
                const cell = worksheet[address] as any;
                if (!cell) {
                    continue;
                }

                const style = cell.s ? { ...cell.s } : {};
                const font = style.font ? { ...style.font } : {};
                cell.s = {
                    ...style,
                    font: {
                        ...font,
                        name: 'Cambria',
                        sz: 11
                    }
                };
            }
        }
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
        // Disable if no company is selected
        if (!this.selectedBank || this.selectedBank === '') {
            return true;
        }
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
        const bank = this.selectedBank;
        if (bank !== 'US' && bank !== 'UK' && bank !== 'EOS') {
            this.loggingService.logError(new Error(`Invalid bank selection: ${bank}`), 'excel_export', 'InvoiceComponent');
            alert('Please select a company before exporting.');
            return;
        }

        const itemsSubtotal = items.reduce((sum, item) => sum + (item.total || 0), 0);
        const discountAmount = itemsSubtotal * (this.invoiceData.discountPercent || 0) / 100;
        const feesTotal = includeFees
            ? ((this.invoiceData.deliveryFee || 0) +
                (this.invoiceData.portFee || 0) +
                (this.invoiceData.agencyFee || 0) +
                (this.invoiceData.transportCustomsLaunchFees || 0) +
                (this.invoiceData.launchFee || 0))
            : 0;
        const categoryTotal = (itemsSubtotal - discountAmount) + feesTotal;
        const categoryToUse = categoryOverride || this.invoiceData.category;

        const workbookData = { ...this.invoiceData, items };

        const options: InvoiceWorkbookOptions = {
            data: workbookData,
            selectedBank: bank,
            primaryCurrency: this.primaryCurrency,
            categoryOverride,
            appendAtoInvoiceNumber,
            includeFees
        };

        try {
            const { blob, fileName } = await buildInvoiceStyleWorkbook(options);
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

