# Hi Marine Invoicing

A modern Angular web application for processing supplier XLSX files and generating invoices.

## Features

- **Suppliers Tab**: Upload multiple XLSX files via drag-and-drop or file browser
- **Captain's Order Tab**: Upload and process Excel files for captain requests
- **Invoice Tab**: Generate Excel invoices from selected items
- **History Tab**: Placeholder for future functionality

## Installation

1. Install dependencies:
```bash
npm install
```

2. Start the development server:
```bash
npm start
```

3. Open your browser to `http://localhost:4200`

## Usage

### Step 1: Upload Supplier Files
- Navigate to the "Suppliers" tab
- Drag and drop Excel files or click to browse
- The system will automatically analyze each file and detect:
  - File name
  - Data table location
  - Description and price columns

### Step 2: Review Data
- Go to the "Price List" tab to review and select items
- Click "Process Supplier Files"
- Review the processed data (maintains original upload order)
- Check the "Include" checkbox for items you want to invoice

### Step 3: Generate Invoice
- Navigate to the "Invoice" tab
- Review the summary and preview
- Click "Generate Invoice" to download the Excel file

## Technology Stack

- Angular 18
- TypeScript
- SCSS
- XLSX library for Excel file processing
- FileSaver for file downloads

## Project Structure

```
src/
├── app/
│   ├── components/
│   │   ├── suppliers/
│   │   ├── invoice/
│   │   └── history/
│   ├── services/
│   │   └── data.service.ts
│   ├── app.component.*
│   └── app.routes.ts
├── index.html
├── main.ts
└── styles.scss
```

## Build

To build the project for production:

```bash
npm run build
```

The build artifacts will be stored in the `dist/` directory.

## Deployment

The application is hosted on Firebase Hosting:

**Live URL:** https://himarine-invoicing.web.app

To deploy updates:

```bash
npm run build
firebase deploy --only hosting
```

Or use the deployment script (Windows):
```bash
deploy.bat
```

See [DEPLOYMENT.md](DEPLOYMENT.md) for detailed deployment instructions.

## License

Private - HiMarine Company

