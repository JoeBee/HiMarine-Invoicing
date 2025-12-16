# Hi Marine Invoicing

Angular 18 (standalone) web application for processing Excel supplier/captain files and generating professional quotes, price lists, invoices, and audit logs (Firestore).

## Key Tabs

- **RFQ**: Upload captain request files and generate quote workbooks
- **Suppliers**: Upload supplier documents, process them into a selectable price list, export price lists
- **Invoicing**: Upload captain orders and generate invoices (Excel and PDF)
- **Supplier Analysis**: Compare suppliers and export analysis workbooks
- **History**: View/export application logs (requires access)

## Local Development

```bash
npm install
npm start
```

App runs at `http://localhost:4200`.

## Build

```bash
npm run build
```

Output is written to `dist/hi-marine-invoicing/browser/`.

## Deployment (Firebase Hosting)

```bash
npm run build
firebase deploy --only hosting
```

Windows shortcut: `deploy.bat`

Live site: `https://himarine-invoicing.web.app`

See `DEPLOYMENT.md` for details.

## Notes

- **Client-side processing**: Excel parsing and exports run in the browser.
- **Logging**: User actions/errors are written to Firestore and shown in the History tab.

## License

Private - HiMarine Company

