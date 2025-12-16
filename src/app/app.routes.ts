import { Routes } from '@angular/router';

export const routes: Routes = [
    {
        path: '',
        redirectTo: '/rfq/captains-request',
        pathMatch: 'full'
    },
    {
        path: 'suppliers',
        redirectTo: '/suppliers/supplier-docs',
        pathMatch: 'full'
    },
    {
        path: 'invoice',
        redirectTo: '/invoicing/invoice',
        pathMatch: 'full'
    },
    {
        path: 'history',
        loadComponent: () => import('./components/history/history.component').then(m => m.HistoryComponent)
    },
    {
        path: 'invoicing',
        redirectTo: '/invoicing/captains-order',
        pathMatch: 'full'
    },
    {
        path: 'suppliers/supplier-docs',
        loadComponent: () => import('./components/suppliers-docs/suppliers-docs.component').then(m => m.SuppliersDocsComponent)
    },
    {
        path: 'suppliers/price-list',
        loadComponent: () => import('./components/price-list/price-list.component').then(m => m.PriceListComponent)
    },
    {
        path: 'invoicing/captains-order',
        loadComponent: () => import('./components/captains-order/captains-order.component').then(m => m.CaptainsOrderComponent)
    },
    {
        path: 'invoicing/invoice',
        loadComponent: () => import('./components/invoice/invoice.component').then(m => m.InvoiceComponent)
    },
    {
        path: 'rfq',
        redirectTo: '/rfq/captains-request',
        pathMatch: 'full'
    },
    {
        path: 'rfq/captains-request',
        loadComponent: () => import('./components/captains-request/captains-request.component').then(m => m.RfqCaptainsRequestComponent)
    },
    {
        path: 'rfq/proposal',
        loadComponent: () => import('./components/our-quote/our-quote.component').then(m => m.OurQuoteComponent)
    },
    {
        path: 'supplier-analysis',
        redirectTo: '/supplier-analysis/inputs',
        pathMatch: 'full'
    },
    {
        path: 'supplier-analysis/inputs',
        loadComponent: () => import('./components/supplier-analysis-inputs/supplier-analysis-inputs.component').then(m => m.SupplierAnalysisInputsComponent)
    },
    {
        path: 'supplier-analysis/analysis',
        loadComponent: () => import('./components/supplier-analysis-analysis/supplier-analysis-analysis.component').then(m => m.SupplierAnalysisAnalysisComponent)
    }
];

