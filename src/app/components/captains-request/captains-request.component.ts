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
}



