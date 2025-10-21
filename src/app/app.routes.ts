import { Routes } from '@angular/router';

export const routes: Routes = [
    {
        path: '',
        redirectTo: '/suppliers',
        pathMatch: 'full'
    },
    {
        path: 'suppliers',
        loadComponent: () => import('./components/suppliers/suppliers.component').then(m => m.SuppliersComponent)
    },
    {
        path: 'process-data',
        loadComponent: () => import('./components/process-data/process-data.component').then(m => m.ProcessDataComponent)
    },
    {
        path: 'invoice',
        loadComponent: () => import('./components/invoice/invoice.component').then(m => m.InvoiceComponent)
    },
    {
        path: 'history',
        loadComponent: () => import('./components/history/history.component').then(m => m.HistoryComponent)
    }
];

