import * as ExcelJS from 'exceljs';
import JSZip from 'jszip';

export type InvoiceWorkbookBank = 'US' | 'UK' | 'EOS';

export interface InvoiceWorkbookItem {
    pos: number;
    description: string;
    remark: string;
    unit: string;
    qty: number | string;
    price: number;
    total: number;
    currency: string;
    blankNumericColumns?: boolean;
}

export interface InvoiceWorkbookData {
    items: InvoiceWorkbookItem[];
    discountPercent: number;
    deliveryFee: number;
    portFee: number;
    agencyFee: number;
    transportCustomsLaunchFees: number;
    launchFee: number;
    ourCompanyName: string;
    ourCompanyAddress: string;
    ourCompanyAddress2: string;
    ourCompanyCity: string;
    ourCompanyCountry: string;
    ourCompanyPhone: string;
    ourCompanyEmail: string;
    vesselName: string;
    vesselName2: string;
    vesselAddress: string;
    vesselAddress2: string;
    vesselCity: string;
    vesselCountry: string;
    bankName: string;
    bankAddress: string;
    iban: string;
    swiftCode: string;
    accountTitle: string;
    accountNumber: string;
    sortCode: string;
    achRouting?: string;
    intermediaryBic?: string;
    invoiceNumber: string;
    invoiceDate: string;
    vessel: string;
    country: string;
    port: string;
    category: string;
    invoiceDue: string;
    exportFileName?: string;
}

export interface InvoiceWorkbookOptions {
    data: InvoiceWorkbookData;
    selectedBank: InvoiceWorkbookBank;
    primaryCurrency: string;
    categoryOverride?: string;
    appendAtoInvoiceNumber?: boolean;
    includeFees?: boolean;
    fileNameOverride?: string;
    showUSD?: boolean;
}

export interface InvoiceWorkbookResult {
    blob: Blob;
    fileName: string;
}

const currencyLabelMap: Record<string, string> = {
    'NZ$': 'NZD',
    'A$': 'AUD',
    'C$': 'CAD',
    '€': 'EUR',
    '$': 'USD',
    '£': 'GBP'
};

function getCurrencyLabel(currency: string): string {
    return currencyLabelMap[currency] || 'GBP';
}

function getCurrencyExcelFormat(currency: string): string {
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

function sanitizeFileName(name: string): string {
    return name.replace(/[<>:"/\\|?*]/g, '_');
}

function applyCambriaFontToWorkbook(workbook: ExcelJS.Workbook): void {
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

function toUpperCaseText(value: unknown): string {
    if (typeof value === 'string') {
        return value.toUpperCase();
    }
    if (value === undefined || value === null) {
        return '';
    }
    return String(value).toUpperCase();
}

function formatDateAsText(dateString: string): string {
    if (!dateString) {
        return '';
    }
    const date = new Date(dateString);
    if (Number.isNaN(date.getTime())) {
        return '';
    }
    const months = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
    const month = months[date.getMonth()];
    const day = date.getDate().toString().padStart(2, '0');
    const year = date.getFullYear();
    return `${month} ${day}, ${year}`;
}

export async function buildInvoiceStyleWorkbook(options: InvoiceWorkbookOptions): Promise<InvoiceWorkbookResult> {
    const {
        data,
        selectedBank,
        primaryCurrency,
        categoryOverride,
        appendAtoInvoiceNumber,
        includeFees = true,
        fileNameOverride,
        showUSD = false
    } = options;

    const items = data.items;

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Invoice');

    worksheet.properties.showGridLines = false;
    worksheet.views = [{ showGridLines: false }];

    const fetchImage = async (path: string): Promise<{ buffer: ArrayBuffer; width: number; height: number }> => {
        const response = await fetch(path);
        const buffer = await response.arrayBuffer();
        const blob = new Blob([buffer], { type: 'image/png' });
        const url = URL.createObjectURL(blob);
        const img = new Image();
        await new Promise<void>((resolve, reject) => {
            img.onload = () => {
                URL.revokeObjectURL(url);
                resolve();
            };
            img.onerror = () => {
                URL.revokeObjectURL(url);
                reject(new Error(`Failed to load image at ${path}`));
            };
            img.src = url;
        });

        return { buffer, width: img.naturalWidth, height: img.naturalHeight };
    };

    try {
        if (selectedBank === 'EOS') {
            worksheet.mergeCells('A2:D2');
            const eosTitle = worksheet.getCell('A2');
            eosTitle.value = 'EOS SUPPLY LTD';
            eosTitle.font = { name: 'Calibri', size: 16, bold: true, italic: true, color: { argb: 'FF0B2E66' } } as any;
            eosTitle.alignment = { horizontal: 'left', vertical: 'middle' } as any;

            worksheet.mergeCells('E2:G2');
            const eosPhoneUk = worksheet.getCell('E2');
            eosPhoneUk.value = 'Phone: +44 730 7988228';
            eosPhoneUk.font = { name: 'Calibri', size: 11, bold: true, color: { argb: 'FF0B2E66' } } as any;
            eosPhoneUk.alignment = { horizontal: 'right', vertical: 'middle' } as any;

            worksheet.mergeCells('E3:G3');
            const eosPhoneUs = worksheet.getCell('E3');
            eosPhoneUs.value = 'Phone: +1 857 204-5786';
            eosPhoneUs.font = { name: 'Calibri', size: 11, bold: true, color: { argb: 'FF0B2E66' } } as any;
            eosPhoneUs.alignment = { horizontal: 'right', vertical: 'middle' } as any;

            worksheet.mergeCells('E4:G4');
            const eosEmail = worksheet.getCell('E4');
            eosEmail.value = 'office@eos-supply.co.uk';
            eosEmail.font = { name: 'Calibri', size: 11, bold: true, color: { argb: 'FF0B2E66' } } as any;
            eosEmail.alignment = { horizontal: 'right', vertical: 'middle' } as any;

            worksheet.mergeCells('A6:G6');
            const eosBar = worksheet.getCell('A6');
            eosBar.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0B2E66' } } as any;
            worksheet.getRow(6).height = 18;
        } else {
            try {
                const topImage = await fetchImage('assets/images/HIMarineTopImage_sm.png');
                const topImageId = workbook.addImage({
                    buffer: topImage.buffer,
                    extension: 'png'
                });
                worksheet.addImage(topImageId, {
                    tl: { col: 0.75, row: 0.5 },
                    ext: { width: topImage.width, height: topImage.height }
                });
            } catch (error) {
                console.warn('Failed to load top image for workbook export:', error);
            }
        }

        try {
            const bottomImagePath = selectedBank === 'EOS'
                ? 'assets/images/EosSupplyLtdBottomBorder.png'
                : 'assets/images/HIMarineBottomBorder.png';
            const bottomImage = await fetchImage(bottomImagePath);
            const bottomImageId = workbook.addImage({
                buffer: bottomImage.buffer,
                extension: 'png'
            });
            (worksheet as any)._bottomImageId = bottomImageId;
            (worksheet as any)._bottomImageWidth = bottomImage.width;
            (worksheet as any)._bottomImageHeight = bottomImage.height;
        } catch (error) {
            console.warn('Failed to load bottom image for workbook export:', error);
        }
    } catch (imageError) {
        console.warn('Could not render header/footer for Excel export:', imageError);
    }

    worksheet.getColumn('A').width = 56 / 7;
    worksheet.getColumn('B').width = 374 / 7;
    worksheet.getColumn('C').width = 254 / 7;
    worksheet.getColumn('D').width = 80 / 7;
    worksheet.getColumn('E').width = 82 / 7;
    worksheet.getColumn('F').width = 131 / 7;
    worksheet.getColumn('G').width = 120 / 7;

    const companyDetails = [
        data.ourCompanyName,
        data.ourCompanyAddress,
        data.ourCompanyAddress2,
        data.ourCompanyCity,
        data.ourCompanyCountry,
        data.ourCompanyPhone,
        data.ourCompanyEmail
    ];
    let companyRow = selectedBank === 'EOS' ? 8 : 9;
    companyDetails.forEach(detail => {
        if (detail && detail.trim()) {
            const cell = worksheet.getCell(`A${companyRow}`);
            cell.value = detail;
            cell.font = { size: 11, name: 'Calibri', bold: true };
            cell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: false } as any;
            companyRow++;
        }
    });
    companyRow++;

    const vesselDetails = [
        data.vesselName,
        data.vesselName2,
        data.vesselAddress,
        data.vesselAddress2,
        [data.vesselCity, data.vesselCountry].filter(Boolean).join(', ')
    ];
    let vesselRow: number;
    if (selectedBank === 'US' || selectedBank === 'UK') {
        vesselRow = 9;
    } else if (selectedBank === 'EOS') {
        vesselRow = 8;
    } else {
        vesselRow = 6;
    }
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

    const bankDetailsStartRow = Math.max(companyRow, vesselRow);
    let bankRow = bankDetailsStartRow;
    const writeBankLine = (row: number, label: string, value: string) => {
        worksheet.mergeCells(`A${row}:D${row}`);
        const cell = worksheet.getCell(`A${row}`);
        cell.value = {
            richText: [
                { text: `${label}: `, font: { bold: true, size: 11, name: 'Calibri' } },
                { text: `${value || ''}`, font: { size: 11, name: 'Calibri', bold: true } }
            ]
        } as any;
        cell.font = { size: 11, name: 'Calibri', bold: true };
        cell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true } as any;
    };

    const standardBankDetails = [
        { label: 'Bank Name', value: data.bankName },
        { label: 'Bank Address', value: data.bankAddress },
        { label: 'IBAN', value: data.iban },
        { label: 'Swift Code', value: data.swiftCode },
        { label: 'Title on Account', value: data.accountTitle }
    ];
    standardBankDetails.forEach(detail => {
        if (detail.value && detail.value.trim()) {
            writeBankLine(bankRow, detail.label, detail.value);
            bankRow++;
        }
    });

    if (selectedBank === 'US') {
        if (data.achRouting && data.achRouting.trim()) {
            writeBankLine(bankRow, 'ACH Routing', data.achRouting);
            bankRow++;
        }
    } else if (selectedBank === 'UK') {
        if (data.accountNumber && data.accountNumber.trim()) {
            writeBankLine(bankRow, 'Account Number', data.accountNumber);
            bankRow++;
        }
        if (data.sortCode && data.sortCode.trim()) {
            writeBankLine(bankRow, 'Sort Code', data.sortCode);
            bankRow++;
        }

        bankRow++;
        worksheet.mergeCells(`A${bankRow}:D${bankRow}`);
        const ukDomesticHeader = worksheet.getCell(`A${bankRow}`);
        ukDomesticHeader.value = 'UK DOMESTIC WIRES:';
        ukDomesticHeader.font = { bold: true, size: 11, name: 'Calibri' };
        ukDomesticHeader.alignment = { horizontal: 'left', vertical: 'middle', wrapText: false } as any;
        bankRow++;

        if (data.accountNumber && data.accountNumber.trim()) {
            writeBankLine(bankRow, 'Account number', data.accountNumber);
            bankRow++;
        }
        if (data.sortCode && data.sortCode.trim()) {
            writeBankLine(bankRow, 'Sort code', data.sortCode);
            bankRow++;
        }
    } else if (selectedBank === 'EOS') {
        if (data.intermediaryBic && data.intermediaryBic.trim()) {
            writeBankLine(bankRow, 'Intermediary BIC', data.intermediaryBic);
            bankRow++;
        }
    }

    const invoiceDetailsStartRow = vesselRow + 1;
    const writeInvoiceDetail = (row: number, label: string, value: string, isDate: boolean = false) => {
        const labelCell = worksheet.getCell(`E${row}`);
        labelCell.value = `${label}:`;
        labelCell.font = { size: 11, name: 'Calibri', bold: true };
        labelCell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: false } as any;

        const valueCell = worksheet.getCell(`F${row}`);
        valueCell.value = isDate && value ? formatDateAsText(value) : (value || '');
        valueCell.font = { size: 11, name: 'Calibri', bold: true };
        valueCell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: false } as any;
    };

    const categoryToUse = categoryOverride || data.category;
    const invoiceNumberToUse = appendAtoInvoiceNumber ? `${data.invoiceNumber}a` : data.invoiceNumber;
    let invoiceRow = invoiceDetailsStartRow;
    const invoiceDetails = [
        { label: 'No', value: invoiceNumberToUse, isDate: false },
        { label: 'Invoice Date', value: data.invoiceDate, isDate: true },
        { label: 'Vessel', value: data.vessel, isDate: false },
        { label: 'Country', value: data.country, isDate: false },
        { label: 'Port', value: data.port, isDate: false },
        { label: 'Category', value: categoryToUse, isDate: false },
        { label: 'Invoice Due', value: data.invoiceDue, isDate: false }
    ];
    invoiceDetails.forEach(detail => {
        writeInvoiceDetail(invoiceRow, detail.label, detail.value || '', detail.isDate);
        invoiceRow++;
    });

    const tableStartRow = Math.max(bankRow, invoiceRow) + 2;
    const headers = ['Pos', 'Description', 'Remark', 'Unit', 'Qty', 'Price', 'Total'];
    headers.forEach((header, index) => {
        const cell = worksheet.getCell(tableStartRow, index + 1);
        cell.value = header;
        cell.font = { bold: true, size: 11, name: 'Calibri', color: { argb: 'FFFFFFFF' } };
        const headerFillColor = selectedBank === 'EOS' ? 'FF0B2E66' : 'FF808080';
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: headerFillColor } } as any;
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
    });

    const itemsSubtotal = items.reduce((sum, item) => sum + (item.total || 0), 0);
    const discountAmount = itemsSubtotal * (data.discountPercent || 0) / 100;
    const feesTotal = includeFees ?
        ((data.deliveryFee || 0) +
            (data.portFee || 0) +
            (data.agencyFee || 0) +
            (data.transportCustomsLaunchFees || 0) +
            (data.launchFee || 0)) : 0;

    items.forEach((item, index) => {
        const rowIndex = tableStartRow + 1 + index;
        const shouldBlankNumericColumns = !!item.blankNumericColumns;

        const posCell = worksheet.getCell(rowIndex, 1);
        posCell.value = index + 1;
        posCell.font = { size: 10, name: 'Calibri' };
        posCell.alignment = { horizontal: 'center', vertical: 'middle' };

        const descCell = worksheet.getCell(rowIndex, 2);
        descCell.value = toUpperCaseText(item.description);
        descCell.font = { size: 10, name: 'Calibri', bold: shouldBlankNumericColumns };
        descCell.alignment = shouldBlankNumericColumns
            ? { horizontal: 'center', vertical: 'middle', wrapText: true }
            : { horizontal: 'left', vertical: 'middle', wrapText: true };

        const remarkCell = worksheet.getCell(rowIndex, 3);
        remarkCell.value = toUpperCaseText(item.remark);
        remarkCell.font = { size: 10, name: 'Calibri' };
        remarkCell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };

        const unitCell = worksheet.getCell(rowIndex, 4);
        unitCell.value = toUpperCaseText(item.unit);
        unitCell.font = { size: 10, name: 'Calibri' };
        unitCell.alignment = { horizontal: 'center', vertical: 'middle' };

        const qtyCell = worksheet.getCell(rowIndex, 5);
        if (shouldBlankNumericColumns) {
            qtyCell.value = null;
        } else {
            const qtyValue = typeof item.qty === 'string' ? Number(item.qty) : item.qty;
            qtyCell.value = Number.isFinite(qtyValue as number) ? qtyValue : (item.qty || 0);
        }
        qtyCell.font = { size: 10, name: 'Calibri' };
        qtyCell.alignment = { horizontal: 'center', vertical: 'middle' };

        const priceCell = worksheet.getCell(rowIndex, 6);
        const currencyFormat = getCurrencyExcelFormat(item.currency || primaryCurrency);
        if (shouldBlankNumericColumns) {
            priceCell.value = null;
            priceCell.numFmt = undefined as any;
        } else {
            priceCell.value = Math.round((item.price || 0) * 100) / 100;
            priceCell.numFmt = currencyFormat;
        }
        priceCell.font = { size: 10, name: 'Calibri' };
        priceCell.alignment = { horizontal: 'right', vertical: 'middle' };

        const totalCell = worksheet.getCell(rowIndex, 7);
        if (shouldBlankNumericColumns) {
            totalCell.value = null;
            totalCell.numFmt = undefined as any;
        } else {
            totalCell.value = { formula: `E${rowIndex}*F${rowIndex}` } as any;
            totalCell.numFmt = currencyFormat;
        }
        totalCell.font = { size: 10, name: 'Calibri' };
        totalCell.alignment = { horizontal: 'right', vertical: 'middle' };

        const border = { top: { style: 'thin' }, bottom: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } } as any;
        posCell.border = border;
        descCell.border = border;
        remarkCell.border = border;
        unitCell.border = border;
        qtyCell.border = border;
        priceCell.border = border;
        totalCell.border = border;
    });

    const firstDataRow = tableStartRow + 1;
    const lastDataRow = tableStartRow + items.length;
    const hasDiscount = discountAmount > 0;
    const hasFees = feesTotal > 0;
    const shouldShowSubtotal = hasDiscount || hasFees;
    const primaryCurrencyFormat = getCurrencyExcelFormat(primaryCurrency);

    let totalsStartRow: number;
    let grandTotalRow: number = 0; // Initialize to avoid TypeScript error
    if (shouldShowSubtotal) {
        const subtotalRow = tableStartRow + items.length + 2;
        const subtotalLabelCell = worksheet.getCell(`F${subtotalRow}`);
        subtotalLabelCell.value = `TOTAL ${getCurrencyLabel(primaryCurrency)}`;
        subtotalLabelCell.font = { size: 11, name: 'Calibri', bold: true };
        subtotalLabelCell.alignment = { horizontal: 'right', vertical: 'middle' };

        const subtotalCell = worksheet.getCell(`G${subtotalRow}`);
        subtotalCell.value = { formula: `SUM(G${firstDataRow}:G${lastDataRow})` } as any;
        subtotalCell.font = { size: 11, name: 'Calibri', bold: true };
        subtotalCell.alignment = { horizontal: 'right', vertical: 'middle' };
        subtotalCell.numFmt = primaryCurrencyFormat;

        totalsStartRow = tableStartRow + items.length + 4;
    } else {
        const totalRow = tableStartRow + items.length + 2;
        const totalLabelCell = worksheet.getCell(`F${totalRow}`);
        totalLabelCell.value = `TOTAL ${getCurrencyLabel(primaryCurrency)}`;
        totalLabelCell.font = { size: 11, name: 'Calibri', bold: true };
        totalLabelCell.alignment = { horizontal: 'right', vertical: 'middle' };

        const totalCell = worksheet.getCell(`G${totalRow}`);
        totalCell.value = { formula: `SUM(G${firstDataRow}:G${lastDataRow})` } as any;
        totalCell.font = { size: 11, name: 'Calibri', bold: true };
        totalCell.alignment = { horizontal: 'right', vertical: 'middle' };
        totalCell.numFmt = primaryCurrencyFormat;

        totalsStartRow = totalRow + 1;
        grandTotalRow = totalRow; // Track the grand total row for when no subtotal
    }

    const feeLines: { label: string; value?: number; includeInSum?: boolean }[] = [];
    if (discountAmount > 0) feeLines.push({ label: 'Discount:', value: -discountAmount, includeInSum: false });
    if (includeFees) {
        if (data.deliveryFee) feeLines.push({ label: 'Delivery fee:', value: data.deliveryFee, includeInSum: true });
        if (data.portFee) feeLines.push({ label: 'Port fee:', value: data.portFee, includeInSum: true });
        if (data.agencyFee) feeLines.push({ label: 'Agency fee:', value: data.agencyFee, includeInSum: true });
        if (data.transportCustomsLaunchFees) feeLines.push({ label: 'Transport, Customs, Launch fees:', value: data.transportCustomsLaunchFees, includeInSum: true });
        if (data.launchFee) feeLines.push({ label: 'Launch:', value: data.launchFee, includeInSum: true });
    }

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
        worksheet.getRow(rowIndex).height = 15;
    });

    totalsStartRow += feeLines.length;
    if (shouldShowSubtotal) {
        worksheet.getCell(`F${totalsStartRow}`).value = '';
        worksheet.getCell(`F${totalsStartRow}`).alignment = { horizontal: 'right', vertical: 'middle' };

        const itemsSumFormula = `SUM(G${firstDataRow}:G${lastDataRow})`;
        const feeSumPart = feeAmountRowRefs.length ? `+${feeAmountRowRefs.join('+')}` : '';
        const discountFactor = data.discountPercent ? `(1-${data.discountPercent}/100)` : '1';
        const totalFormula = `(${itemsSumFormula}*${discountFactor})${feeSumPart}`;
        worksheet.getCell(`G${totalsStartRow}`).value = { formula: totalFormula } as any;
        worksheet.getCell(`G${totalsStartRow}`).font = { bold: true, size: 11, name: 'Calibri' };
        worksheet.getCell(`G${totalsStartRow}`).alignment = { horizontal: 'right', vertical: 'middle' };
        worksheet.getCell(`G${totalsStartRow}`).numFmt = primaryCurrencyFormat;
        grandTotalRow = totalsStartRow; // Track the grand total row for when subtotal exists
    }

    // Add USD total row if showUSD is enabled and currency is not already USD
    if (showUSD && primaryCurrency && primaryCurrency !== '$') {
        const usdRow = grandTotalRow + 1;
        const usdLabelCell = worksheet.getCell(`F${usdRow}`);
        usdLabelCell.value = 'TOTAL USD';
        usdLabelCell.font = { bold: true, size: 11, name: 'Calibri' };
        usdLabelCell.alignment = { horizontal: 'right', vertical: 'middle' };

        // Approximate exchange rates
        const exchangeRates: { [key: string]: number } = {
            '£': 1.27,
            '€': 1.08,
            'A$': 0.66,
            'NZ$': 0.61,
            'C$': 0.73
        };
        const rate = exchangeRates[primaryCurrency] || 1;
        const usdFormula = `G${grandTotalRow}*${rate}`;
        const usdValueCell = worksheet.getCell(`G${usdRow}`);
        usdValueCell.value = { formula: usdFormula } as any;
        usdValueCell.font = { bold: true, size: 11, name: 'Calibri' };
        usdValueCell.alignment = { horizontal: 'right', vertical: 'middle' };
        usdValueCell.numFmt = '$#,##0.00';
        worksheet.getRow(usdRow).height = 15;
        totalsStartRow = usdRow;
    }

    const termsStartRow = totalsStartRow + 1;
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
        const romanCell = worksheet.getCell(`A${row}`);
        romanCell.value = term.roman;
        romanCell.font = { size: 10, name: 'Calibri' };
        romanCell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: false } as any;

        const textCell = worksheet.getCell(`B${row}`);
        textCell.value = term.text;
        textCell.font = { size: 10, name: 'Calibri' };
        textCell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: false } as any;
    });

    let maxBottomOffset = 0;
    if (selectedBank === 'EOS') {
        const fontColor = 'FF0B2E66';
        const bottomFont = { name: 'Calibri', size: 11, bold: true, color: { argb: fontColor } } as any;
        const bottomSectionStart = termsStartRow + terms.length + 4;

        const leftLines: string[] = [
            data.ourCompanyName || '',
            data.ourCompanyAddress || '',
            data.ourCompanyAddress2 || '',
            [data.ourCompanyCity, data.ourCompanyCountry].filter(Boolean).join(', ')
        ].filter(Boolean);

        leftLines.forEach((text, idx) => {
            const rowIndex = bottomSectionStart + idx;
            const cell = worksheet.getCell(`A${rowIndex}`);
            cell.value = text;
            cell.font = bottomFont;
            cell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: false } as any;
        });

        const rightStartRow = bottomSectionStart;
        const writeRight = (offset: number, text: string) => {
            const row = rightStartRow + offset;
            worksheet.mergeCells(`D${row}:G${row}`);
            const cell = worksheet.getCell(`D${row}`);
            cell.value = text;
            cell.font = bottomFont;
            cell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: false } as any;
        };

        let currentOffset = 0;
        writeRight(currentOffset++, `Bank Name: ${data.bankName || ''}`);
        writeRight(currentOffset++, `Bank Address: ${data.bankAddress || ''}`);
        writeRight(currentOffset++, `IBAN: ${data.iban || ''}`);
        writeRight(currentOffset++, `SWIFTBIC: ${data.swiftCode || ''}`);

        maxBottomOffset = currentOffset - 1;
    }

    const lastTextRow = (selectedBank === 'EOS')
        ? (termsStartRow + terms.length + 4 + maxBottomOffset)
        : (termsStartRow + terms.length);

    const bottomImageRowPosition = lastTextRow + 2;
    if ((worksheet as any)._bottomImageId) {
        const originalWidth = (worksheet as any)._bottomImageWidth || 667;
        const originalHeight = (worksheet as any)._bottomImageHeight || 80;
        const newWidth = 1000;
        const newHeight = (originalHeight * newWidth) / originalWidth;
        worksheet.addImage((worksheet as any)._bottomImageId, {
            tl: { col: 0, row: bottomImageRowPosition },
            ext: { width: newWidth, height: newHeight }
        });
    }

    let printAreaEndRow: number;
    if (selectedBank === 'US' || selectedBank === 'UK') {
        printAreaEndRow = lastTextRow + 10;
    } else {
        printAreaEndRow = bottomImageRowPosition + 1 + 3;
    }
    worksheet.pageSetup.printArea = `A1:G${printAreaEndRow}`;
    worksheet.pageSetup.fitToPage = true;
    worksheet.pageSetup.fitToWidth = 1;
    worksheet.pageSetup.orientation = 'portrait';
    worksheet.pageSetup.paperSize = 9;
    worksheet.pageSetup.margins = {
        top: 0.590,
        bottom: 0.590,
        left: 0.393,
        right: 0.393,
        header: 0.314,
        footer: 0.314
    };
    worksheet.pageSetup.horizontalCentered = true;
    worksheet.pageSetup.verticalCentered = false;

    applyCambriaFontToWorkbook(workbook);

    let buffer = await workbook.xlsx.writeBuffer();
    try {
        const zip = await JSZip.loadAsync(buffer);
        const worksheetFiles = Object.keys(zip.files).filter(name =>
            name.startsWith('xl/worksheets/sheet') && name.endsWith('.xml')
        );

        if (worksheetFiles.length > 0) {
            const worksheetXml = await zip.file(worksheetFiles[0])?.async('string');
            if (worksheetXml) {
                let modifiedXml = worksheetXml;
                if (modifiedXml.includes('<sheetViews>')) {
                    modifiedXml = modifiedXml.replace(
                        /<sheetView([^>]*?)(\s*\/?>)/g,
                        (match, attrs, closing) => {
                            let cleanAttrs = attrs.replace(/\s*view="[^"]*"/g, '');
                            if (cleanAttrs && !cleanAttrs.endsWith(' ')) {
                                cleanAttrs += ' ';
                            }
                            return `<sheetView${cleanAttrs}view="pageBreakPreview"${closing}`;
                        }
                    );

                    if (!modifiedXml.includes('view="pageBreakPreview"')) {
                        modifiedXml = modifiedXml.replace(
                            /<sheetViews>(\s*)<sheetView([^>]*?)(\s*\/?>)/g,
                            (match, spacing, attrs, closing) => {
                                let cleanAttrs = attrs.replace(/\s*view="[^"]*"/g, '');
                                if (cleanAttrs && !cleanAttrs.endsWith(' ')) {
                                    cleanAttrs += ' ';
                                }
                                return `<sheetViews>${spacing}<sheetView${cleanAttrs}view="pageBreakPreview"${closing}`;
                            }
                        );

                        if (!modifiedXml.includes('view="pageBreakPreview"')) {
                            modifiedXml = modifiedXml.replace(
                                '<sheetViews>',
                                '<sheetViews><sheetView view="pageBreakPreview"/>'
                            );
                        }
                    }
                } else {
                    modifiedXml = modifiedXml.replace(
                        /(<worksheet[^>]*>)/,
                        '$1<sheetViews><sheetView view="pageBreakPreview"/></sheetViews>'
                    );
                }

                zip.file(worksheetFiles[0], modifiedXml);
            }
        }

        buffer = await zip.generateAsync({ type: 'arraybuffer' });
    } catch (zipError) {
        console.warn('Could not modify Excel file for Page Break Preview view:', zipError);
    }

    const blob = new Blob([buffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });

    let fileName: string | undefined = fileNameOverride;
    if (!fileName) {
        if (data.exportFileName && data.exportFileName.trim()) {
            fileName = data.exportFileName.trim();
            if (categoryOverride && fileName.includes('<Category>')) {
                fileName = fileName.replace('<Category>', categoryOverride);
            }
        } else {
            const filePrefix = selectedBank === 'EOS' ? 'EOS_Invoice' : 'HIMarine_Invoice';
            const categorySuffix = categoryOverride ? `_${categoryOverride}` : '';
            fileName = `${filePrefix}_${data.invoiceNumber}${categorySuffix}_${new Date().toISOString().split('T')[0]}`;
        }
    }

    fileName = sanitizeFileName(fileName);
    if (!fileName.endsWith('.xlsx')) {
        fileName = `${fileName}.xlsx`;
    }

    return { blob, fileName };
}

