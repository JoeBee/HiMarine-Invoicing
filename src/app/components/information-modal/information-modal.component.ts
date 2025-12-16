import { Component, EventEmitter, Output, OnInit, OnDestroy, HostListener, ElementRef, ViewChild } from '@angular/core';
import { CommonModule } from '@angular/common';

@Component({
    selector: 'app-information-modal',
    standalone: true,
    imports: [CommonModule],
    templateUrl: './information-modal.component.html',
    styleUrls: ['./information-modal.component.scss']
})
export class InformationModalComponent implements OnInit, OnDestroy {
    @Output() close = new EventEmitter<void>();
    activeTab: 'rfq' | 'suppliers' | 'invoicing' | 'supplier-analysis' | 'history' = 'rfq';
    hasHistoryAccess = false;
    private readonly STORAGE_KEY = 'history_access_granted';
    private storageListener?: (event: StorageEvent) => void;
    private historyAccessChangedListener?: () => void;

    @ViewChild('modal', { static: false }) modal?: ElementRef<HTMLElement>;

    ngOnInit(): void {
        this.checkHistoryAccess();
        queueMicrotask(() => this.modal?.nativeElement.focus());
        
        // Listen for storage changes (e.g., when user logs out in another tab)
        this.storageListener = (event: StorageEvent) => {
            if (event.key === this.STORAGE_KEY) {
                this.checkHistoryAccess();
            }
        };
        window.addEventListener('storage', this.storageListener);
        
        // Listen for custom logout event (for same-tab logout)
        this.historyAccessChangedListener = () => {
            this.checkHistoryAccess();
        };
        window.addEventListener('historyAccessChanged', this.historyAccessChangedListener);
    }

    ngOnDestroy(): void {
        if (this.storageListener) {
            window.removeEventListener('storage', this.storageListener);
        }
        if (this.historyAccessChangedListener) {
            window.removeEventListener('historyAccessChanged', this.historyAccessChangedListener);
        }
    }

    private checkHistoryAccess(): void {
        // Check if user has entered the secret word
        const hasAccess = localStorage.getItem(this.STORAGE_KEY);
        const previousAccess = this.hasHistoryAccess;
        this.hasHistoryAccess = hasAccess === 'true';
        
        // If access was revoked and history tab is selected, switch to rfq
        if (previousAccess && !this.hasHistoryAccess && this.activeTab === 'history') {
            this.activeTab = 'rfq';
        }
    }

    closeModal(): void {
        this.close.emit();
    }

    @HostListener('document:keydown.escape')
    onEscapeKey(): void {
        this.closeModal();
    }

    selectTab(tab: 'rfq' | 'suppliers' | 'invoicing' | 'supplier-analysis' | 'history'): void {
        // Only allow selecting history tab if user has access
        if (tab === 'history' && !this.hasHistoryAccess) {
            return;
        }
        this.activeTab = tab;
    }
}

