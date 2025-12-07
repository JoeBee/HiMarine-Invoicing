# EOS/Hi Marine Invoicing - Website Structure Overview

## Application Overview

**EOS/Hi Marine Invoicing** is a modern Angular 18 single-page application designed for processing supplier Excel files and generating professional invoices for ocean vessel shipping operations. The application handles bonded items, provisions, and fresh provisions with automated file processing, data extraction, and multi-format invoice generation.

**Live URL:** https://himarine-invoicing.web.app

## High-Level Architecture

### Technology Stack
- **Framework:** Angular 18 (Standalone Components)
- **Language:** TypeScript 5.4
- **Styling:** SCSS
- **Excel Processing:** XLSX 0.18.5
- **File Operations:** FileSaver 2.0.5
- **Reactive Programming:** RxJS 7.8
- **Backend Services:** Firebase Firestore (for logging)
- **Hosting:** Firebase Hosting

### Architecture Pattern
- **Standalone Components:** No NgModules, all components are standalone
- **Service-Based State Management:** Centralized data via BehaviorSubjects
- **Lazy Loading:** Route-based code splitting for performance
- **Client-Side Processing:** All file operations occur in the browser
- **Reactive Data Flow:** Observable-based state propagation

## Navigation Structure

### Main Navigation Tabs (Top Level)
Located in `app.component.html`:

1. **Request for Quotation (RFQ) Tab** (`/rfq/*`)

            - Sub-tabs:
                    - Captain's Request (`/rfq/captains-request`)
                    - Our Quote (`/rfq/proposal`)

2. **Suppliers Tab** (`/suppliers/*`)

            - Sub-tabs:
                    - Supplier Docs (`/suppliers/supplier-docs`)
                    - Price List (`/suppliers/price-list`)

3. **Invoicing Tab** (`/invoicing/*`)

            - Sub-tabs:
                    - Captain's Order (`/invoicing/captains-order`)
                    - Invoice (`/invoicing/invoice`)

4. **History Tab** (`/history`)

            - No sub-tabs

### Routing Configuration

File: `src/app/app.routes.ts`

Routes use lazy loading with `loadComponent()` for optimal performance:
- Default route redirects to `/rfq/captains-request`
- All routes use standalone component imports
- Routes are organized by main tab sections

## Component Hierarchy

### Root Component

**Location:** `src/app/app.component.*`

**Responsibilities:**
- Main navigation tabs management
- Sub-navigation rendering based on active tab
- Info modal management
- Route tracking and active tab state
- Logging navigation events

**Key Features:**

- Responsive navigation with emoji icons
- Modal overlay with comprehensive user instructions
- Session-based navigation logging

### Core Components

#### 1. RfqCaptainsRequestComponent

**Location:** `src/app/components/captains-request/`
**Route:** `/rfq/captains-request`

**Purpose:** Upload and process captain request Excel files for RFQ generation

**Key Features:**
- Excel file upload interface
- File processing and data extraction
- RFQ data preparation
- Integration with RFQ state service

**Data Flow:**
- Files uploaded → `RfqStateService`
- Processed data stored in RFQ state
- Data flows to Our Quote component

#### 2. OurQuoteComponent

**Location:** `src/app/components/our-quote/`
**Route:** `/rfq/proposal`

**Purpose:** Generate and manage RFQ proposals (quotes)

**Key Features:**
- Company selection (HI US, HI UK, EOS) with theme switching
- Currency selection (EUR, AUD, NZD, USD, CAD)
- Dynamic company branding (top and bottom images) - only shown when proposal tables exist
- Export filename input with separate `.xlsx` label
- Company details and bank details configuration
- Billing Information section (formerly Vessel Details)
- Invoice details configuration
- Proposal tables display with collapsible sections
- Create RFQ functionality
- Theme-based styling (EOS theme: light blue, Hi Marine theme: light blue/red)

**Data Flow:**
- Reads from `RfqStateService`
- Manages proposal data and RFQ generation
- Exports RFQ proposals

**Theme Features:**
- EOS theme: Light blue background (#e6eefc), EOS branding images
- Hi Marine theme: Light blue/red background, Hi Marine branding images
- Top logo and bottom border images switch based on company selection
- Images only display when `proposalTables.length > 0`

#### 3. SuppliersDocsComponent

**Location:** `src/app/components/suppliers-docs/`
**Route:** `/suppliers/supplier-docs`

**Purpose:** Upload and categorize supplier Excel files

**Key Features:**

- Drag-and-drop file upload interface
- Categorized drop zones:
        - Bonded (for restricted items)
        - Provisions (for general and fresh provisions)
- Automatic file analysis on upload
- File metadata display (name, detected columns, row count)
- Real-time processing feedback

**Data Flow:**
- Files uploaded → `DataService.addSupplierFiles()`
- Analysis results stored in `DataService.supplierFiles$`

#### 4. PriceListComponent

**Location:** `src/app/components/price-list/`

**Route:** `/suppliers/price-list`

**Purpose:** Process uploaded files and select items for invoicing

**Key Features:**
- Company selection (EOS, Hi Marine) with theme switching
- Currency selection (EUR, AUD, NZD, USD, CAD)
- Country selection dropdown
- Dynamic company branding (top and bottom images) - only shown when processed data exists
- Export filename input with separate `.xlsx` label (flexible width to fit container)
- "Process Supplier Files" button (enabled when files exist)
- Data extraction and display
- Item selection via checkboxes
- Price calculation with divider
- Export processed data to Excel format
- Maintains original file upload order
- Filtering by file name, description, and text search
- Sortable columns
- Expandable rows showing original data

**Data Flow:**
- Reads from `DataService.supplierFiles$`
- Processes files → `DataService.processSupplierFiles()`
- Stores processed data in `DataService.processedData$`
- Updates item selection state
- Exports to `DataService.excelData$`

**Theme Features:**
- EOS theme: Light blue background (#e6eefc), EOS branding images
- Hi Marine theme: Light red background (#ffe6e6), Hi Marine branding images
- Top logo and bottom border images switch based on company selection
- Images only display when `processedData.length > 0`

#### 5. CaptainsOrderComponent

**Location:** `src/app/components/captains-order/`

**Route:** `/invoicing/captains-order`

**Purpose:** Upload and process captain order Excel files

**Key Features:**
- Excel file upload interface
- Similar processing workflow to supplier files
- Specialized for captain orders

#### 6. InvoiceComponent

**Location:** `src/app/components/invoice/`

**Route:** `/invoicing/invoice`

**Purpose:** Generate and export professional invoices

**Key Features:**

- Company selection (US, UK, EOS) with theme switching
- Split Invoices / One Invoice toggle switch (default: Split Invoices)
- Dynamic company branding (top and bottom images) - only shown when invoice items exist
- Export filename input with separate `.xlsx` label
- Invoice preview and summary statistics
- Multiple export formats:
        - Excel (.xlsx) with formulas and formatting
        - PDF (.pdf) print-ready format
- Split file logic:
        - Creates separate files for 'Bonded' and 'Provisions' categories
        - Appends "A" to invoice number for Provisions when both exist
        - Skips empty categories
- Excel export features:
        - Billing Information (values only, no labels) - formerly "Vessel Details"
        - Tab column excluded
        - Items renumbered from 1 in each file
        - Company branding and bank details
        - Automatic calculations and formulas
- Fees & Discounts section (Delivery fee, Port fee, Agency fee, Transport/Customs/Launch fees, Launch fee, Discount %)
- Billing Information section (Name, Name 2, Address, Address 2, City, Country)
- Invoice Details section (No, Invoice Date, Vessel, Country, Port, Category, Invoice Due)

**Data Flow:**
- Reads from `DataService.excelData$`
- Generates invoices based on selected mode
- Downloads files via FileSaver

**Theme Features:**
- EOS theme: Light blue background (#e6eefc), EOS branding images, light blue editable fields
- Hi Marine theme: Light blue background (#f0f7ff), Hi Marine branding images
- Top logo and bottom border images switch based on company selection
- Images only display when `invoiceData.items.length > 0`

#### 7. HistoryComponent

**Location:** `src/app/components/history/`

**Route:** `/history`

**Purpose:** Display application logs and audit trail

**Key Features:**

- Comprehensive log viewing
- Advanced filtering:
        - Category filter (dropdown)
        - Level filter (dropdown)
        - Component filter (dropdown)
        - IP Address filter (dropdown, sorted by most recent)
        - Date range picker
        - Search text input
- Pagination with direct page jump
- CSV export functionality
- Real-time log updates
- Responsive table display

**Data Flow:**
- Reads from Firebase Firestore via `LoggingService`
- Filters applied client-side
- Pagination handles large datasets

#### 8. InvoicingComponent

**Location:** `src/app/components/invoicing/`

**Route:** `/invoicing` (legacy/alternative route)

**Purpose:** Alternative invoicing interface

#### 9. SuppliersComponent

**Location:** `src/app/components/suppliers/`

**Route:** `/suppliers-new` (legacy/alternative route)

**Purpose:** Alternative supplier interface

## Services Architecture

### DataService

**Location:** `src/app/services/data.service.ts`

**Responsibilities:**
- Centralized data management
- Excel file analysis and processing
- State management via BehaviorSubjects
- Data transformation between formats

**Observables Exposed:**
- `supplierFiles$` - Uploaded file metadata
- `processedData$` - Extracted data rows
- `priceDivider$` - Price adjustment factor
- `excelData$` - Final processed data for invoicing

**Key Methods:**

- `addSupplierFiles()` - Add and analyze new files
- `processSupplierFiles()` - Extract data from files
- `updateProcessedData()` - Modify processed data
- `setPriceDivider()` - Adjust price calculations
- `exportToExcel()` - Format data for Excel export

**Interfaces:**

- `SupplierFileInfo` - File metadata structure
- `ProcessedDataRow` - Extracted row data
- `ExcelItemData` - Final invoice item format
- `ExcelProcessedData` - Organized invoice data

### RfqStateService

**Location:** `src/app/services/rfq-state.service.ts`

**Responsibilities:**
- RFQ (Request for Quotation) state management
- Proposal data management
- Company and currency selection state
- RFQ generation and export

**Key Features:**
- Manages RFQ data structure (company details, bank details, billing information, invoice details)
- Handles proposal tables and items
- Company selection state (HI US, HI UK, EOS)
- Currency selection state
- Country and port management
- Export filename generation

### LoggingService

**Location:** `src/app/services/logging.service.ts`

**Responsibilities:**
- Application event logging
- User action tracking
- Error logging
- Firebase Firestore integration
- Offline queue management

**Log Categories:**

- `user_action` - Button clicks, navigation
- `file_upload` - File operations
- `data_processing` - Data extraction
- `export` - Invoice generation
- `error` - Error conditions
- `system` - System events

**Key Features:**

- Batch processing for performance
- Offline queue with sync on reconnect
- IP address detection
- Session tracking
- Automatic cleanup (7-day retention)

**Methods:**
- `logUserAction()` - General user actions
- `logButtonClick()` - Button interactions
- `logFileUpload()` - File upload events
- `logError()` - Error logging
- `getLogs()` - Retrieve logs from Firestore

## Data Flow Architecture

### Supplier File Processing Flow

```
1. User uploads files → SuppliersDocsComponent
   ↓
2. Files sent to DataService.addSupplierFiles()
   ↓
3. Each file analyzed:
   - Find top-left data cell
   - Detect description column
   - Detect price column (must be < 25 chars)
   - Detect unit column
   - Detect remarks column
   ↓
4. File metadata stored in supplierFiles$
   ↓
5. User navigates to Price List → PriceListComponent
   ↓
6. User clicks "Process Supplier Files"
   ↓
7. DataService.processSupplierFiles() extracts data:
   - Reads XLSX files
   - Extracts rows with descriptions and prices
   - Applies price divider
   - Maintains upload order
   ↓
8. Processed data stored in processedData$
   ↓
9. User selects items via checkboxes
   ↓
10. Data exported to Excel format → excelData$
```

### Invoice Generation Flow

```
1. User navigates to Invoice tab → InvoiceComponent
   ↓
2. Component reads from DataService.excelData$
   ↓
3. User selects export mode:
   - Split Invoices (default) → separate files per category
   - One Invoice → single combined file
   ↓
4. User clicks "Export Invoice"
   ↓
5. If Split Invoices:
   - Group items by category (BOND vs PROVISIONS/FRESH PROVISIONS)
   - Create separate workbook for each non-empty category
   - Append "A" to Provisions invoice number if both exist
   ↓
6. Generate Excel workbook with:
   - Company details and branding
   - Billing Information (values only) - formerly "Vessel Details"
   - Invoice details (number, date, category)
   - Items table (Pos, Description, Remark, Unit, Qty, Price, Total)
   - Totals and fees section
   - Terms and conditions
   - Bank details
   - Footer image
   ↓
7. Apply formatting:
   - Column widths
   - Cell styles
   - Formulas (Total = Qty * Price)
   - Merged cells
   - Print area
   ↓
8. Download file(s) via FileSaver
```

### Logging Flow

```
1. Action occurs in any component
   ↓
2. Component calls LoggingService method
   ↓
3. Service creates LogEntry object:
   - Timestamp
   - Level (info/warn/error/debug)
   - Category
   - Action name
   - Details object
   - Session ID
   - Component name
   - URL
   - User agent
   - IP address
   ↓
4. Entry added to local queue
   ↓
5. Batch processor sends to Firebase Firestore:
   - Batches of 10 entries
   - 5-second timeout
   - Retry on failure
   ↓
6. HistoryComponent reads from Firestore
   ↓
7. Client-side filtering and pagination
```

## File Structure

```
HiMarine-Invoicing/
├── src/
│   ├── app/
│   │   ├── components/
│   │   │   ├── captains-request/        # RFQ captain request processing
│   │   │   │   ├── captains-request.component.ts
│   │   │   │   ├── captains-request.component.html
│   │   │   │   └── captains-request.component.scss
│   │   │   ├── our-quote/               # RFQ proposal/quote generation
│   │   │   │   ├── our-quote.component.ts
│   │   │   │   ├── our-quote.component.html
│   │   │   │   └── our-quote.component.scss
│   │   │   ├── suppliers-docs/          # Main supplier file upload
│   │   │   │   ├── suppliers-docs.component.ts
│   │   │   │   ├── suppliers-docs.component.html
│   │   │   │   └── suppliers-docs.component.scss
│   │   │   ├── price-list/              # Data processing and selection
│   │   │   │   ├── price-list.component.ts
│   │   │   │   ├── price-list.component.html
│   │   │   │   └── price-list.component.scss
│   │   │   ├── captains-order/         # Captain order processing
│   │   │   ├── invoice/                 # Invoice generation
│   │   │   ├── history/                 # Log viewing
│   │   │   ├── information-modal/       # Information/help modal
│   │   │   ├── invoicing/               # Alternative invoicing (legacy)
│   │   │   ├── suppliers/                # Alternative suppliers (legacy)
│   │   │   └── process-data/            # (Empty/unused)
│   │   ├── services/
│   │   │   ├── data.service.ts          # Core data management
│   │   │   ├── logging.service.ts       # Logging and analytics
│   │   │   └── rfq-state.service.ts     # RFQ state management
│   │   ├── utils/
│   │   │   └── invoice-workbook-builder.ts  # Excel workbook generation utility
│   │   ├── app.component.ts             # Root component
│   │   ├── app.component.html           # Navigation and modal
│   │   ├── app.component.scss           # Global navigation styles
│   │   └── app.routes.ts                # Route configuration
│   ├── assets/                          # Static assets
│   │   └── images/                      # Images and logos
│   ├── index.html                       # HTML entry point
│   ├── main.ts                          # Application bootstrap
│   └── styles.scss                      # Global styles
├── public/
│   └── favicon.png                      # Ship emoji favicon
├── functions/                           # Firebase Functions (if any)
├── angular.json                         # Angular CLI config
├── package.json                         # Dependencies
├── tsconfig.json                        # TypeScript config
└── Documentation files...
```

## Key Features by Component

### SuppliersDocsComponent Features
- Multi-file drag-and-drop upload
- Categorized drop zones (Bonded/Provisions)
- Automatic Excel file analysis
- Column detection (description, price, unit, remarks)
- File metadata table display
- Real-time processing feedback

### PriceListComponent Features

- Company selection (EOS, Hi Marine) with theme switching
- Batch file processing
- Currency selection (EUR, AUD, NZD, USD, CAD)
- Country selection dropdown
- Export filename with separate `.xlsx` label (flexible width)
- Dynamic company branding (top/bottom images, conditional display)
- Item inclusion checkboxes
- Price divider adjustment
- Excel export of processed data
- Maintains file upload order
- Advanced filtering (file name, description, text search)
- Sortable columns
- Expandable rows showing original data

### InvoiceComponent Features

- Company selection (US, UK, EOS) with theme switching
- Split Invoices / One Invoice toggle
- Dynamic company branding (top/bottom images, conditional display)
- Export filename with separate `.xlsx` label
- Invoice preview with summary
- Excel export with full formatting
- PDF export capability
- Automatic item renumbering
- Billing Information export (values only) - formerly "Vessel Details"
- Tab column exclusion
- Conditional invoice numbering (append "A")
- Company branding and bank details
- Fees & Discounts configuration
- Billing Information section (Name, Name 2, Address, Address 2, City, Country)
- Invoice Details section (No, Invoice Date, Vessel, Country, Port, Category, Invoice Due)

### HistoryComponent Features

- Multi-criteria filtering:
        - Category, Level, Component, IP Address
        - Date range selection
        - Text search
- Pagination with direct page jump
- CSV export
- Real-time log updates
- IP address sorting by most recent activity

## State Management Pattern

### Observable-Based State
All state is managed through RxJS Observables in services:

**DataService State:**
- `supplierFilesSubject` → `supplierFiles$`
- `processedDataSubject` → `processedData$`
- `priceDividerSubject` → `priceDivider$`
- `excelDataSubject` → `excelData$`

**Component State:**
- Components subscribe to service observables
- Local component state for UI-specific data
- Automatic UI updates via async pipe or subscriptions

### Data Persistence
- **Session-based:** Data persists during browser session
- **Firebase Firestore:** Logs persist across sessions (7-day retention)
- **No database:** Invoice data and file contents are in-memory only
- **Reset on refresh:** Application state clears on page reload

## User Workflows

### RFQ Workflow

1. **Upload Captain Request**

            - Navigate to RFQ → Captain's Request
            - Upload captain request Excel files
            - Process files for RFQ generation

2. **Create Proposal/Quote**

            - Navigate to RFQ → Our Quote
            - Select company (HI US, HI UK, or EOS)
            - Select currency
            - Fill in company details, bank details, billing information
            - Configure invoice details
            - Review proposal tables
            - Click "Create RFQ(s)" to generate

### Standard Invoice Workflow

1. **Upload Supplier Files**

            - Navigate to Suppliers → Supplier Docs
            - Drag files into appropriate category zones
            - Review detected columns in table

2. **Process and Select Items**

            - Navigate to Suppliers → Price List
            - Select company (EOS or Hi Marine)
            - Select currency
            - Select country (optional)
            - Click "Process Supplier Files"
            - Review extracted data
            - Check items to include
            - Adjust prices if needed
            - Export price list if needed

3. **Generate Invoice**

            - Navigate to Invoicing → Invoice
            - Select company (US, UK, or EOS)
            - Choose export mode (Split Invoices or One Invoice)
            - Fill in billing information
            - Configure fees and discounts
            - Review preview and summary
            - Click "Export Invoice" (Excel) or "Generate PDF"
            - Download file(s)

### Captain Order Workflow

1. Navigate to Invoicing → Captain's Order
2. Upload captain order Excel file
3. Process similar to supplier files
4. Generate invoice from captain order data

### Log Review Workflow

1. Navigate to History tab
2. Apply filters as needed
3. Navigate pages or jump to specific page
4. Export logs to CSV if needed

## Styling Architecture

### Global Styles

- `src/styles.scss` - Base styles, resets, variables
- `src/app/app.component.scss` - Navigation and modal styles

### Component Styles

- Each component has its own `.scss` file
- Component-specific styles scoped to component
- Shared variables and mixins in global styles

### Design Patterns
- Modern, professional appearance
- Gradient summary cards
- Color-coded states (active, hover, disabled)
- Responsive layout
- Smooth animations and transitions
- Theme-based styling (EOS vs Hi Marine)
- Dynamic branding with conditional image display
- Flexible form layouts with proper spacing

## Deployment Architecture

### Build Process

- Angular CLI production build
- Output to `dist/hi-marine-invoicing/`
- Firebase Hosting deployment

### Environment

- **Development:** `npm start` → `http://localhost:4200`
- **Production:** Firebase Hosting → `https://himarine-invoicing.web.app`
- **Deployment Script:** `deploy.bat` (Windows)

### Firebase Integration
- **Hosting:** Static site hosting
- **Firestore:** Log storage and retrieval
- **Configuration:** `firebase.config.ts`

## Security and Performance

### Security Features

- Client-side only processing (no server exposure)
- Firebase security rules for Firestore
- No sensitive data stored locally
- Session-based data (clears on refresh)

### Performance Optimizations

- Lazy-loaded components
- Route-based code splitting
- Batch log processing
- Efficient Excel file parsing
- Optimized Observable subscriptions

## Theme System

### Company Theme Switching

The application supports dynamic theme switching based on company selection:

**EOS Theme:**
- Background: Light blue (#e6eefc)
- Top Logo: `EosSupplyLtd.png`
- Bottom Border: `EosSupplyLtdBottomBorder.png`
- Editable fields: Light blue highlight (#d9edf7)
- Radio button accent: Dark blue (#0B2E66)

**Hi Marine Theme:**
- Background: Light red/blue (varies by component)
- Top Logo: `HIMarineTopImage_sm.png`
- Bottom Border: `HIMarineBottomBorder.png`
- Radio button accent: Dark green (#2E7D32)

**Implementation:**
- Theme classes applied via `[ngClass]` directive
- Top and bottom images conditionally displayed based on data availability
- Images only show when processed data/proposal tables/invoice items exist
- Theme switching happens automatically when company selection changes

**Components with Theme Support:**
- PriceListComponent (EOS, Hi Marine)
- InvoiceComponent (EOS, Hi Marine - US/UK)
- OurQuoteComponent (EOS, Hi Marine - HI US/HI UK)

## Recent Updates

### UI/UX Improvements

1. **Export Filename Enhancement:**
   - `.xlsx` extension removed from input field
   - Added as separate label after input (matching invoice component)
   - Flexible input width to fit container properly
   - Applied to PriceListComponent

2. **Terminology Update:**
   - "Vessel Details" renamed to "Billing Information" throughout application
   - Updated in all components, services, and documentation
   - More accurate representation of the data being collected

3. **Conditional Image Display:**
   - Top logo and bottom border images only display when data exists
   - Prevents empty state from showing branding
   - Applied to all components with theme support

## Future Enhancement Areas

Based on codebase analysis, potential improvements:
- History tab: Invoice history tracking, search, statistics
- Data persistence: LocalStorage or database for invoice history
- User management: Multi-user support, roles
- Template management: Save invoice templates
- Reporting: Advanced analytics and reporting
- RFQ workflow: Enhanced proposal management and tracking

