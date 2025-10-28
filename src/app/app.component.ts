import { Component } from '@angular/core';
import { RouterModule, Router, NavigationEnd } from '@angular/router';
import { CommonModule } from '@angular/common';
import { filter } from 'rxjs/operators';
import { LoggingService } from './services/logging.service';

@Component({
    selector: 'app-root',
    standalone: true,
    imports: [RouterModule, CommonModule],
    templateUrl: './app.component.html',
    styleUrls: ['./app.component.scss']
})
export class AppComponent {
    title = 'HI Marine Invoicing';
    showInfoModal = false;
    activeMainTab = '';

    constructor(private router: Router, private loggingService: LoggingService) {
        // Log application initialization
        this.loggingService.logSystemEvent('application_initialized', {
            timestamp: new Date().toISOString(),
            userAgent: navigator.userAgent,
            url: window.location.href
        }, 'AppComponent');

        // Listen to route changes to determine active main tab
        this.router.events
            .pipe(filter(event => event instanceof NavigationEnd))
            .subscribe((event) => {
                const navigationEvent = event as NavigationEnd;
                this.updateActiveMainTab(navigationEvent.url);
            });
    }

    updateActiveMainTab(url: string): void {
        if (url.startsWith('/suppliers')) {
            this.activeMainTab = 'suppliers';
        } else if (url.startsWith('/invoicing')) {
            this.activeMainTab = 'invoicing';
        } else if (url.startsWith('/history')) {
            this.activeMainTab = 'history';
        } else {
            this.activeMainTab = '';
        }
    }

    onMainTabClick(tab: string): void {
        this.loggingService.logButtonClick(`main_tab_${tab}`, 'AppComponent', {
            previousTab: this.activeMainTab,
            newTab: tab
        });

        if (tab === 'suppliers') {
            this.router.navigate(['/suppliers/supplier-docs']);
        } else if (tab === 'invoicing') {
            this.router.navigate(['/invoicing/captains-request']);
        } else if (tab === 'history') {
            this.router.navigate(['/history']);
        }
    }

    openInfoModal(): void {
        this.loggingService.logButtonClick('info_modal_open', 'AppComponent');
        this.showInfoModal = true;
    }

    closeInfoModal(): void {
        this.loggingService.logButtonClick('info_modal_close', 'AppComponent');
        this.showInfoModal = false;
    }
}

