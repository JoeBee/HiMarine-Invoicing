import { Component, OnInit } from '@angular/core';
import { CommonModule, CurrencyPipe } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { CurrencyCode, ProposalItem, ProposalTable, RfqData, RfqStateService } from '../../services/rfq-state.service';

@Component({
    selector: 'app-our-quote',
    standalone: true,
    imports: [CommonModule, FormsModule],
    templateUrl: './our-quote.component.html',
    styleUrls: ['./our-quote.component.scss'],
    providers: [CurrencyPipe]
})
export class OurQuoteComponent implements OnInit {
    public rfqData: RfqData;
    exportFileName = '';
    readonly currencyOptions: Array<{ code: CurrencyCode; label: string }> = [
        { code: 'EUR', label: 'EUR (€)' },
        { code: 'AUD', label: 'AUD (A$)' },
        { code: 'NZD', label: 'NZD (NZ$)' },
        { code: 'USD', label: 'USD ($)' },
        { code: 'CAD', label: 'CAD (C$)' }
    ];
    private selectedCompanyInternal: ('HI US' | 'HI UK' | 'EOS') | null = null;
    private selectedCurrencyInternal: CurrencyCode | null = null;

    constructor(public rfqState: RfqStateService, private currencyPipe: CurrencyPipe) {
        this.rfqData = this.rfqState.rfqData;
    }

    ngOnInit(): void {
        this.updateExportFileName();
    }

    get selectedCompany(): ('HI US' | 'HI UK' | 'EOS') | null {
        return this.selectedCompanyInternal;
    }

    set selectedCompany(value: ('HI US' | 'HI UK' | 'EOS') | null) {
        if (this.selectedCompanyInternal === value) {
            return;
        }
        this.selectedCompanyInternal = value;
        if (value) {
            this.onCompanyChange(value);
        }
    }

    get selectedCurrency(): CurrencyCode | null {
        return this.selectedCurrencyInternal;
    }

    set selectedCurrency(value: CurrencyCode | null) {
        if (this.selectedCurrencyInternal === value) {
            return;
        }
        this.selectedCurrencyInternal = value;
        if (value) {
            this.rfqState.setSelectedCurrency(value);
        }
    }

    get isProcessing(): boolean {
        return this.rfqState.isProcessing;
    }

    get errorMessage(): string {
        return this.rfqState.errorMessage;
    }

    get countries(): string[] {
        return this.rfqState.countries;
    }

    get availablePorts(): string[] {
        return this.rfqState.availablePorts;
    }

    onExportFileNameChange(): void {
        // Allow manual override but keep value in sync for potential future use.
        this.exportFileName = (this.exportFileName || '').trim();
    }

    onCompanyChange(company: 'HI US' | 'HI UK' | 'EOS'): void {
        this.rfqState.onCompanyChange(company);
        this.updateExportFileName();
    }

    onCountryChange(): void {
        this.rfqState.onCountryChange();
        this.updateExportFileName();
    }

    canCreateRFQs(): boolean {
        return this.rfqState.canCreateRFQs();
    }

    createRFQs(): void {
        this.rfqState.createRFQs();
    }

    onProposalDetailChange(): void {
        this.updateExportFileName();
    }

    get proposalItems(): ProposalItem[] {
        return this.rfqState.proposalItems;
    }

    get proposalTables(): ProposalTable[] {
        return this.rfqState.proposalTables;
    }

    get proposalTableCount(): number {
        return this.proposalTables.length;
    }

    formatCurrency(value?: string | number | null): string {
        const numericValue = this.parseNumericValue(value);
        if (numericValue === null) {
            if (typeof value === 'string') {
                return value;
            }
            return '';
        }

        if (!this.selectedCurrency) {
            return numericValue.toFixed(2);
        }

        return this.currencyPipe.transform(
            numericValue,
            this.selectedCurrency,
            'symbol',
            '1.2-2'
        ) ?? numericValue.toFixed(2);
    }

    private updateExportFileName(): void {
        const parts: string[] = [];

        const companyLabel = this.getCompanyLabel(this.selectedCompany);
        if (companyLabel) {
            parts.push(companyLabel);
        }

        const vesselSegment = this.buildInvoiceDetailSegment('Vessel', this.rfqData.vessel);
        if (vesselSegment) {
            parts.push(vesselSegment);
        }

        const portSegment = this.buildInvoiceDetailSegment('Port', this.rfqData.port);
        if (portSegment) {
            parts.push(portSegment);
        }

        const categorySegment = this.buildInvoiceDetailSegment('Category', this.rfqData.category);
        if (categorySegment) {
            parts.push(categorySegment);
        }

        this.exportFileName = parts.filter(Boolean).join(' ');
    }

    private getCompanyLabel(company: ('HI US' | 'HI UK' | 'EOS') | null): string {
        switch (company) {
            case 'HI US':
            case 'HI UK':
                return 'Hi Marine';
            case 'EOS':
                return 'EOS Supply';
            default:
                return '';
        }
    }

    private buildInvoiceDetailSegment(_label: string, value?: string): string {
        const trimmedValue = value?.trim();
        if (!trimmedValue) {
            return '';
        }
        return trimmedValue;
    }

    private parseNumericValue(value: string | number | null | undefined): number | null {
        if (value === undefined || value === null) {
            return null;
        }

        if (typeof value === 'number') {
            return Number.isFinite(value) ? value : null;
        }

        const trimmed = value.trim();
        if (!trimmed) {
            return null;
        }

        let cleaned = trimmed.replace(/NZ\$/gi, '');
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

}


