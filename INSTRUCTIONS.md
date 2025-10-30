# HiMarine Invoicing - User Instructions

## Getting Started

### First Time Setup
1. Open a terminal/command prompt in the project folder
2. Run `npm install` to install dependencies
3. Run `npm start` or double-click `start-app.bat` to start the application
4. Open your browser to `http://localhost:4200`

## How to Use

### 1. Suppliers Tab - Upload Your Files

**Purpose:** Upload XLSX files from your suppliers.

**Steps:**
1. Navigate to the "Suppliers" tab (first tab)
2. Either:
   - Drag and drop Excel files into the drop zone
   - Click the drop zone to browse and select files
3. You can upload multiple files at once
4. The system automatically analyzes each file and detects:
   - File name (without extension)
   - Top left cell where the data table starts
   - Column containing descriptions
   - Column containing prices
5. View all uploaded files in the table below

**What to expect:**
- Files are processed immediately upon upload
- The detected columns are shown in the table
- You can upload more files at any time

---

### 2. Price List Tab - Extract and Review Data

**Purpose:** Process the uploaded files and prepare data for invoicing.

**Steps:**
1. Navigate to the "Price List" tab
2. Click the "Process Supplier Files" button
   - This button is only enabled when files have been uploaded
3. Wait for processing to complete
4. Review the processed data in the table
5. Check the "Include" checkbox for each item you want to include in the invoice

**What to expect:**
- Data maintains the original order from uploaded files
- Each row shows:
  - Include checkbox
  - File name
  - Item description
  - Price
- You can select/deselect items as needed

---

### 3. Invoice Tab - Generate Your Invoice

**Purpose:** Create an Excel invoice from selected items.

**Steps:**
1. Navigate to the "Invoice" tab
2. Review the summary card showing:
   - Total items processed
   - Items selected for invoice
   - Total amount
3. View the preview table of selected items
4. Click "Generate Invoice" button
5. The Excel file will download automatically

**What to expect:**
- Only items with "Include" checked will be in the invoice
- The invoice is formatted as an Excel file (.xlsx)
- File name includes the current date
- The button is disabled if no items are selected

---

### 4. History Tab

**Purpose:** Placeholder for future functionality.

Currently displays a fun historical illustration of George Washington meeting Henry VIII.

---

## Info Icon (ℹ️)

Click the blue Info icon in the top-right corner of any page to view these instructions in a modal dialog.

---

## Tips

- **File Format:** Only XLSX and XLS files are supported
- **Multiple Uploads:** You can upload files multiple times on the Suppliers tab
- **Data Order:** Data maintains the original order from uploaded files
- **Column Detection:** The system looks for common column headers like "Description", "Item", "Price", "Cost", etc.
- **Checkboxes:** Only items with "Include" checked will appear in the generated invoice

---

## Troubleshooting

**Problem:** Button is disabled
- **Solution:** Make sure you've completed the previous steps (upload files, process data, select items)

**Problem:** No data appears after processing
- **Solution:** Check that your Excel files have data tables with description and price columns

**Problem:** Wrong columns detected
- **Solution:** The system looks for common headers. Make sure your Excel files have clear column headers

---

## Technical Details

- Built with Angular 18
- Uses XLSX library for Excel file processing
- All processing happens in your browser (no data is sent to any server)
- Files are stored in memory only while using the application

---

## Need Help?

For technical support, contact your system administrator.

