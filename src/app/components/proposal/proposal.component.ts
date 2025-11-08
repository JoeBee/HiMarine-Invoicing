import { Component, OnInit } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { FileAnalysis, RfqData, RfqStateService, TabInfo } from '../../services/rfq-state.service';

@Component({
    selector: 'app-proposal',
    standalone: true,
    imports: [CommonModule, FormsModule],
    templateUrl: './proposal.component.html',
    styleUrls: ['./proposal.component.scss']
})
export class ProposalComponent implements OnInit {
    public rfqData: RfqData;
    exportFileName = '';

    constructor(public rfqState: RfqStateService) {
        this.rfqData = this.rfqState.rfqData;
    }

    ngOnInit(): void {
        this.updateExportFileName();
    }

    get selectedCompany(): 'HI US' | 'HI UK' | 'EOS' {
        return this.rfqState.selectedCompany;
    }

    set selectedCompany(value: 'HI US' | 'HI UK' | 'EOS') {
        this.rfqState.selectedCompany = value;
        this.updateExportFileName();
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

    get fileAnalyses(): FileAnalysis[] {
        return this.rfqState.fileAnalyses;
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

    clearAllFiles(): void {
        this.rfqState.clearAllFiles();
    }

    canCreateRFQs(): boolean {
        return this.rfqState.canCreateRFQs();
    }

    createRFQs(): void {
        this.rfqState.createRFQs();
    }

    openFile(analysis: FileAnalysis): void {
        this.rfqState.openFile(analysis);
    }

    onExcludeChange(tab: TabInfo): void {
        this.rfqState.onExcludeChange(tab);
    }

    onColumnChange(analysis: FileAnalysis, tab: TabInfo, columnType: 'product' | 'qty' | 'unit' | 'remark'): void {
        this.rfqState.onColumnChange(analysis, tab, columnType);
    }

    getTopLeftCellOptionsLimited(): string[] {
        return this.rfqState.getTopLeftCellOptionsLimited();
    }

    onTopLeftCellChange(event: Event, tab: TabInfo): void {
        this.rfqState.onTopLeftCellChange(event, tab);
    }

    validateTopLeftCell(tab: TabInfo): void {
        this.rfqState.validateTopLeftCell(tab);
    }

    onTopLeftCellFocus(event: Event, fileIndex: number, tabIndex: number): void {
        this.rfqState.onTopLeftCellFocus(event, fileIndex, tabIndex);
    }

    onProposalDetailChange(): void {
        this.updateExportFileName();
    }

    private updateExportFileName(): void {
        const parts: string[] = [];
        parts.push('Proposal');

        if (this.rfqData.invoiceNumber?.trim()) {
            parts.push(this.rfqData.invoiceNumber.trim());
        }

        const formattedDate = this.formatDateForFileName(this.rfqData.invoiceDate);
        if (formattedDate) {
            parts.push(formattedDate);
        }

        if (this.rfqData.vessel?.trim()) {
            parts.push(this.rfqData.vessel.trim());
        }

        if (this.rfqData.country?.trim()) {
            parts.push(this.rfqData.country.trim());
        }

        if (this.rfqData.port?.trim()) {
            parts.push(this.rfqData.port.trim());
        }

        if (this.rfqData.category?.trim()) {
            parts.push(this.rfqData.category.trim());
        }

        const companyLabel = this.getCompanyLabel(this.selectedCompany);
        if (companyLabel) {
            parts.push(companyLabel);
        }

        this.exportFileName = parts.filter(Boolean).join(' ');
    }

    private getCompanyLabel(company: 'HI US' | 'HI UK' | 'EOS'): string {
        switch (company) {
            case 'HI US':
                return 'HI-US';
            case 'HI UK':
                return 'HI-UK';
            case 'EOS':
                return 'EOS';
            default:
                return '';
        }
    }

    private formatDateForFileName(dateString?: string): string {
        if (!dateString) {
            return '';
        }
        const date = new Date(dateString);
        if (Number.isNaN(date.getTime())) {
            return '';
        }
        const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sept', 'Oct', 'Nov', 'Dec'];
        const month = months[date.getMonth()];
        const day = date.getDate().toString().padStart(2, '0');
        const year = date.getFullYear();
        return `${day}-${month}-${year}`;
    }
}


