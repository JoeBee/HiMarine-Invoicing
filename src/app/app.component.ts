import { Component } from '@angular/core';
import { RouterModule } from '@angular/router';
import { CommonModule } from '@angular/common';

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

    openInfoModal(): void {
        this.showInfoModal = true;
    }

    closeInfoModal(): void {
        this.showInfoModal = false;
    }
}

