import { Component } from '@angular/core';
import { CommonModule } from '@angular/common';
import { RfqStateService } from '../../services/rfq-state.service';

@Component({
    selector: 'app-rfq-captains-request',
    standalone: true,
    imports: [CommonModule],
    templateUrl: './captains-request.component.html',
    styleUrls: ['../proposal/proposal.component.scss']
})
export class RfqCaptainsRequestComponent {
    isDragOver = false;

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
}



