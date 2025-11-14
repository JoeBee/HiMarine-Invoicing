import { Component, EventEmitter, Output } from '@angular/core';
import { CommonModule } from '@angular/common';

@Component({
    selector: 'app-information-modal',
    standalone: true,
    imports: [CommonModule],
    templateUrl: './information-modal.component.html',
    styleUrls: ['./information-modal.component.scss']
})
export class InformationModalComponent {
    @Output() close = new EventEmitter<void>();
    activeTab: 'rfq' | 'suppliers' | 'invoicing' | 'history' = 'rfq';

    closeModal(): void {
        this.close.emit();
    }

    selectTab(tab: 'rfq' | 'suppliers' | 'invoicing' | 'history'): void {
        this.activeTab = tab;
    }
}

