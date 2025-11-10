import { Component, OnInit } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { ProposalItem, ProposalTable, RfqData, RfqStateService } from '../../services/rfq-state.service';

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

    private getCompanyLabel(company: 'HI US' | 'HI UK' | 'EOS'): string {
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

}


