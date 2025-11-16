import { Component, OnInit } from '@angular/core';
import { RouterModule, Router, NavigationEnd } from '@angular/router';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { filter } from 'rxjs/operators';
import { LoggingService } from './services/logging.service';
import { InformationModalComponent } from './components/information-modal/information-modal.component';

@Component({
    selector: 'app-root',
    standalone: true,
    imports: [RouterModule, CommonModule, FormsModule, InformationModalComponent],
    templateUrl: './app.component.html',
    styleUrls: ['./app.component.scss']
})
export class AppComponent implements OnInit {
    title = 'Hi Marine Invoicing';
    showInfoModal = false;
    activeMainTab = '';
    showHistoryTab = false;
    passwordInput = '';
    showPassword = false;
    private readonly SECRET_WORD = 'drakemaye';
    private readonly STORAGE_KEY = 'history_access_granted';

    constructor(private router: Router, private loggingService: LoggingService) {
        // Listen to route changes to determine active main tab
        this.router.events
            .pipe(filter(event => event instanceof NavigationEnd))
            .subscribe((event) => {
                const navigationEvent = event as NavigationEnd;
                this.updateActiveMainTab(navigationEvent.urlAfterRedirects);
            });

        this.activeMainTab = this.getMainTabFromUrl(this.router.url);
    }

    ngOnInit(): void {
        // Check if user has previously entered the secret word
        const hasAccess = localStorage.getItem(this.STORAGE_KEY);
        if (hasAccess === 'true') {
            this.showHistoryTab = true;
            // Fill password input with masked characters (use text type to show asterisks)
            this.passwordInput = '*'.repeat(this.SECRET_WORD.length);
            this.showPassword = true; // Use text type to show asterisks
        }
    }

    updateActiveMainTab(url: string): void {
        this.activeMainTab = this.getMainTabFromUrl(url);
    }

    onMainTabClick(tab: string): void {
        this.loggingService.logButtonClick(`main_tab_${tab}`, 'AppComponent', {
            previousTab: this.activeMainTab,
            newTab: tab
        });

        if (tab === 'suppliers') {
            this.router.navigate(['/suppliers/supplier-docs']);
        } else if (tab === 'invoicing') {
            this.router.navigate(['/invoicing/captains-order']);
        } else if (tab === 'rfq') {
            this.router.navigate(['/rfq/captains-request']);
        } else if (tab === 'history') {
            this.router.navigate(['/history']);
        }
    }

    private getMainTabFromUrl(url: string): string {
        if (url.startsWith('/suppliers')) {
            return 'suppliers';
        }
        if (url.startsWith('/invoicing')) {
            return 'invoicing';
        }
        if (url.startsWith('/rfq')) {
            return 'rfq';
        }
        if (url.startsWith('/history')) {
            return 'history';
        }
        return '';
    }

    openInfoModal(): void {
        this.loggingService.logButtonClick('info_modal_open', 'AppComponent');
        this.showInfoModal = true;
    }

    closeInfoModal(): void {
        this.loggingService.logButtonClick('info_modal_close', 'AppComponent');
        this.showInfoModal = false;
    }

    onPasswordInput(event: Event): void {
        const input = event.target as HTMLInputElement;
        let value = input.value;

        // If already authenticated, prevent editing
        if (this.showHistoryTab) {
            // Keep the masked value (use text type to show asterisks)
            this.passwordInput = '*'.repeat(this.SECRET_WORD.length);
            input.value = '*'.repeat(this.SECRET_WORD.length);
            this.showPassword = true; // Always show as text with asterisks when authenticated
            return;
        }

        // Check if the entered value matches the secret word (case sensitive)
        if (value === this.SECRET_WORD) {
            this.showHistoryTab = true;
            localStorage.setItem(this.STORAGE_KEY, 'true');
            // Log successful secret word entry
            this.loggingService.logUserAction(
                'secret_word_entered',
                {
                    success: true,
                    timestamp: new Date().toISOString()
                },
                'AppComponent'
            );
            // Switch to text type and show asterisks
            this.showPassword = true; // Use text type to show asterisks
            // Mask the input with asterisks immediately
            this.passwordInput = '*'.repeat(this.SECRET_WORD.length);
            // Also set the input value directly to ensure it's masked
            input.value = '*'.repeat(this.SECRET_WORD.length);
        } else {
            this.passwordInput = value;
        }
    }

    logout(): void {
        // Hide the History tab
        this.showHistoryTab = false;
        // Clear localStorage
        localStorage.removeItem(this.STORAGE_KEY);
        // Clear password input
        this.passwordInput = '';
        this.showPassword = false;
        // Navigate away from history if currently on that page
        if (this.activeMainTab === 'history') {
            this.router.navigate(['/rfq/captains-request']);
        }
        // Log the logout action
        this.loggingService.logUserAction(
            'history_access_logout',
            {
                timestamp: new Date().toISOString()
            },
            'AppComponent'
        );
    }

}

