import { Component } from '@angular/core';
import { RouterModule, Router, NavigationEnd } from '@angular/router';
import { CommonModule } from '@angular/common';
import { filter } from 'rxjs/operators';

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

    constructor(private router: Router) {
        // Listen to route changes to determine active main tab
        this.router.events
            .pipe(filter(event => event instanceof NavigationEnd))
            .subscribe((event) => {
                this.updateActiveMainTab((event as NavigationEnd).url);
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
        if (tab === 'suppliers') {
            this.router.navigate(['/suppliers/supplier-docs']);
        } else if (tab === 'invoicing') {
            this.router.navigate(['/invoicing/captains-request']);
        } else if (tab === 'history') {
            this.router.navigate(['/history']);
        }
    }

    openInfoModal(): void {
        this.showInfoModal = true;
    }

    closeInfoModal(): void {
        this.showInfoModal = false;
    }
}

