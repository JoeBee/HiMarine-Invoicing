import { Routes } from '@angular/router';

export const routes: Routes = [
    {
        path: '',
        redirectTo: '/rfq/captains-request',
        pathMatch: 'full'
    },
    {
        path: 'suppliers',
        loadComponent: () => import('./components/suppliers-docs/suppliers-docs.component').then(m => m.SuppliersDocsComponent)
    },
    {
        path: 'invoice',
        loadComponent: () => import('./components/invoice/invoice.component').then(m => m.InvoiceComponent)
    },
    {
        path: 'history',
        loadComponent: () => import('./components/history/history.component').then(m => m.HistoryComponent)
    },
    {
        path: 'suppliers-new',
        loadComponent: () => import('./components/suppliers/suppliers.component').then(m => m.SuppliersComponent)
    },
    {
        path: 'invoicing',
        loadComponent: () => import('./components/invoicing/invoicing.component').then(m => m.InvoicingComponent)
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
        loadComponent: () => import('./components/proposal/proposal.component').then(m => m.ProposalComponent)
    }
];

