import { Component } from '@angular/core';
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
export class ProposalComponent {
    public rfqData: RfqData;

    constructor(public rfqState: RfqStateService) {
        this.rfqData = this.rfqState.rfqData;
    }

    get selectedCompany(): 'HI US' | 'HI UK' | 'EOS' {
        return this.rfqState.selectedCompany;
    }

    set selectedCompany(value: 'HI US' | 'HI UK' | 'EOS') {
        this.rfqState.selectedCompany = value;
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

    onCompanyChange(company: 'HI US' | 'HI UK' | 'EOS'): void {
        this.rfqState.onCompanyChange(company);
    }

    onCountryChange(): void {
        this.rfqState.onCountryChange();
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
}


