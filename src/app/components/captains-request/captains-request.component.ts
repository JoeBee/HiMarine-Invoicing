import { Component } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { RfqStateService } from '../../services/rfq-state.service';

@Component({
    selector: 'app-rfq-captains-request',
    standalone: true,
    imports: [CommonModule, FormsModule],
    templateUrl: './captains-request.component.html',
    styleUrls: ['../our-quote/our-quote.component.scss']
})
export class RfqCaptainsRequestComponent {
    isDragOver = false;
    removeDialogVisible = false;
    fileIndexPendingRemoval: number | null = null;
    pendingRemovalFileName: string | null = null;
    columnInfoDialogVisible = false;
    columnInfoDialogTitle = '';
    columnInfoDialogDescription = '';
    columnInfoDialogItems: string[] = [];
    columnInfoDialogFootnote: string = '';

    private readonly columnInfoConfig: Record<'product' | 'qty' | 'unit' | 'price' | 'remark', {
        title: string;
        description: string;
        items: string[];
        footnote?: string;
    }> = {
        product: {
            title: 'Product Column Auto-Mapping',
            description: 'We auto-select the Product dropdown when an Excel header matches one of these values:',
            items: [
                '"Product Name"',
                '"Description"',
                '"Equipment Description"'
            ],
            footnote: 'Matching is case-insensitive and ignores leading/trailing spaces.'
        },
        qty: {
            title: 'Quantity Column Auto-Mapping',
            description: 'Quantity dropdowns auto-select when we detect one of the following header names:',
            items: [
                '"Requested Qty"',
                '"Quantity"',
                '"Qty"'
            ],
            footnote: 'If your file uses a different label, update the dropdown manually.'
        },
        unit: {
            title: 'Unit Column Auto-Mapping',
            description: 'Units auto-select when the Excel header matches any of these common labels:',
            items: [
                '"Unit Type"',
                '"Unit"',
                '"UOM"',
                '"UN"'
            ],
            footnote: 'Composite selections (primary/secondary/tertiary) are built from multiple matching headers.'
        },
        price: {
            title: 'Price Column Auto-Mapping',
            description: 'We auto-select Price when we find this exact header:',
            items: [
                '"Price"'
            ],
            footnote: 'Matching is case-insensitive and ignores leading/trailing spaces.'
        },
        remark: {
            title: 'Remark Column Auto-Mapping',
            description: 'Remark dropdowns auto-select when headers match any of the following:',
            items: [
                '"Product No"',
                '"Product No."',
                '"Remark"',
                '"Remarks"',
                '"Impa"'
            ],
            footnote: 'Additional remark columns can be combined by selecting secondary/tertiary dropdowns.'
        }
    };

    constructor(public rfqState: RfqStateService) { }

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
            this.rfqState.handleFiles(Array.from(files));
        }
    }

    onFileSelected(event: Event): void {
        const input = event.target as HTMLInputElement;
        if (input.files && input.files.length > 0) {
            this.rfqState.handleFiles(Array.from(input.files));
            input.value = '';
        }
    }

    openRemoveDialog(fileIndex: number): void {
        this.fileIndexPendingRemoval = fileIndex;
        this.pendingRemovalFileName = this.rfqState.fileAnalyses[fileIndex]?.fileName ?? null;
        this.removeDialogVisible = true;
    }

    confirmRemoveFile(): void {
        if (this.fileIndexPendingRemoval !== null) {
            this.rfqState.removeFile(this.fileIndexPendingRemoval);
        }
        this.resetRemoveDialogState();
    }

    cancelRemoveFile(): void {
        this.resetRemoveDialogState();
    }

    private resetRemoveDialogState(): void {
        this.removeDialogVisible = false;
        this.fileIndexPendingRemoval = null;
        this.pendingRemovalFileName = null;
    }

    openColumnInfoDialog(column: 'product' | 'qty' | 'unit' | 'price' | 'remark'): void {
        const config = this.columnInfoConfig[column];
        if (!config) {
            return;
        }

        this.columnInfoDialogTitle = config.title;
        this.columnInfoDialogDescription = config.description;
        this.columnInfoDialogItems = [...config.items];
        this.columnInfoDialogFootnote = config.footnote ?? '';
        this.columnInfoDialogVisible = true;
    }

    closeColumnInfoDialog(): void {
        this.columnInfoDialogVisible = false;
        this.columnInfoDialogTitle = '';
        this.columnInfoDialogDescription = '';
        this.columnInfoDialogItems = [];
        this.columnInfoDialogFootnote = '';
    }
}



