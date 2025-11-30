import { Component, OnInit } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { LoggingService } from '../../services/logging.service';
import { SupplierAnalysisService, SupplierAnalysisFileInfo } from '../../services/supplier-analysis.service';
import * as XLSX from 'xlsx';

interface InputDropzoneSet {
    id: number;
    invoiceFiles: SupplierAnalysisFileInfo[];
    supplierQuotationFiles: SupplierAnalysisFileInfo[];
    expanded: boolean;
    hoveredDropzone: string | null;
}

@Component({
    selector: 'app-supplier-analysis-inputs',
    standalone: true,
    imports: [CommonModule, FormsModule],
    templateUrl: './supplier-analysis-inputs.component.html',
    styleUrls: ['./supplier-analysis-inputs.component.scss']
})
export class SupplierAnalysisInputsComponent implements OnInit {
    dropzoneSets: InputDropzoneSet[] = [];
    isDragging = false;
    isProcessing = false;
    instructionsExpanded = false;
    isDraggingFile = false;
    draggedFile: SupplierAnalysisFileInfo | null = null;
    dragOverFile: SupplierAnalysisFileInfo | null = null;

    constructor(
        private loggingService: LoggingService,
        private supplierAnalysisService: SupplierAnalysisService
    ) { }

    ngOnInit(): void {
        const fileSets = this.supplierAnalysisService.getFileSets();
        
        if (fileSets.length > 0) {
            this.dropzoneSets = fileSets.map(set => ({
                id: set.id,
                invoiceFiles: set.files.filter(f => f.category === 'Invoice'),
                supplierQuotationFiles: set.files.filter(f => f.category === 'Supplier Quotations'),
                expanded: false, 
                hoveredDropzone: null
            }));
            
            // Expand sets that are not complete
            this.dropzoneSets.forEach(set => {
                 if (!this.isSetComplete(set)) {
                     set.expanded = true;
                 }
            });

            // If the last set is complete, add a new one
            const lastSet = this.dropzoneSets[this.dropzoneSets.length - 1];
            if (this.isSetComplete(lastSet)) {
                this.addNewSet();
            }
        } else {
            this.addNewSet();
        }
    }

    addNewSet(): void {
        const newId = this.dropzoneSets.length > 0 ? Math.max(...this.dropzoneSets.map(s => s.id)) + 1 : 1;
        this.dropzoneSets.push({
            id: newId,
            invoiceFiles: [],
            supplierQuotationFiles: [],
            expanded: true,
            hoveredDropzone: null
        });
    }

    isSetComplete(set: InputDropzoneSet): boolean {
        return set.invoiceFiles.length > 0 && set.supplierQuotationFiles.length > 0;
    }

    onDragOver(event: DragEvent, category: string, setIndex: number): void {
        event.preventDefault();
        event.stopPropagation();
        this.isDragging = true;
        this.dropzoneSets[setIndex].hoveredDropzone = category;
    }

    onDragLeave(event: DragEvent, setIndex: number): void {
        event.preventDefault();
        event.stopPropagation();
        this.isDragging = false;
        this.dropzoneSets[setIndex].hoveredDropzone = null;
    }

    async onDrop(event: DragEvent, category: 'Invoice' | 'Supplier Quotations', setIndex: number): Promise<void> {
        event.preventDefault();
        event.stopPropagation();
        this.isDragging = false;
        this.dropzoneSets[setIndex].hoveredDropzone = null;

        const files = event.dataTransfer?.files;
        if (files) {
            this.loggingService.logUserAction('files_dropped', {
                category: category,
                filesCount: files.length,
                fileNames: Array.from(files).map(f => f.name),
                setIndex: setIndex
            }, 'SupplierAnalysisInputsComponent');

            await this.processFiles(files, category, setIndex);
        }
    }

    onFileSelect(event: Event, category: 'Invoice' | 'Supplier Quotations', setIndex: number): void {
        const input = event.target as HTMLInputElement;
        if (input.files) {
            this.loggingService.logUserAction('files_selected', {
                category: category,
                filesCount: input.files.length,
                fileNames: Array.from(input.files).map(f => f.name),
                setIndex: setIndex
            }, 'SupplierAnalysisInputsComponent');

            this.processFiles(input.files, category, setIndex);
        }
    }

    async processFiles(fileList: FileList, category: 'Invoice' | 'Supplier Quotations', setIndex: number): Promise<void> {
        this.isProcessing = true;
        const files = Array.from(fileList).filter(f =>
            f.name.toLowerCase().endsWith('.xlsx') || 
            f.name.toLowerCase().endsWith('.xls') ||
            f.name.toLowerCase().endsWith('.xlsm')
        );

        if (files.length > 0) {
            files.forEach(file => {
                this.loggingService.logFileUpload(
                    file.name,
                    file.size,
                    file.type,
                    category,
                    'SupplierAnalysisInputsComponent'
                );
            });

            try {
                const currentSet = this.dropzoneSets[setIndex];

                // Invoice: only allow 1 file, replace existing
                if (category === 'Invoice') {
                    currentSet.invoiceFiles = [];
                    // Process only the first file
                    if (files.length > 0) {
                        const fileInfo = await this.analyzeFile(files[0], category);
                        currentSet.invoiceFiles.push(fileInfo);
                    }
                } else {
                    // Supplier Quotations: allow multiple files
                    for (const file of files) {
                        const fileInfo = await this.analyzeFile(file, category);
                        currentSet.supplierQuotationFiles.push(fileInfo);
                    }
                }

                this.loggingService.logDataProcessing('files_processed_successfully', {
                    fileCount: files.length,
                    category: category,
                    fileNames: files.map(f => f.name),
                    setIndex: setIndex
                }, 'SupplierAnalysisInputsComponent');

                // Update service with files for this set
                this.updateServiceFileSet(currentSet);
                
                // Expand set if there's a mismatch
                if (this.hasAnyMismatch(currentSet)) {
                    currentSet.expanded = true;
                }

                // Check if set is complete and we need to add a new one
                if (this.isSetComplete(currentSet)) {
                    currentSet.expanded = false; // Collapse current set
                    
                    // Only add new set if this was the last one
                    if (setIndex === this.dropzoneSets.length - 1) {
                        this.addNewSet();
                    }
                }

            } catch (error) {
                this.loggingService.logError(error as Error, 'file_processing', 'SupplierAnalysisInputsComponent', {
                    fileCount: files.length,
                    category: category,
                    setIndex: setIndex
                });
            }
        }
        this.isProcessing = false;
    }

    private async analyzeFile(file: File, category: 'Invoice' | 'Supplier Quotations'): Promise<SupplierAnalysisFileInfo> {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();

            reader.onload = (e: any) => {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const firstSheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[firstSheetName];

                    const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1:Z100');
                    
                    // Auto-detect the top left cell of the datatable (similar to captains-request)
                    const topLeftCell = this.findFirstDataCell(worksheet, range) || 'A1';
                    const topLeftCellRef = XLSX.utils.decode_cell(topLeftCell);

                    // Count rows starting from the row after topLeftCell
                    let rowCount = 0;
                    let descriptionColumn = -1;

                    // Normalize function to match headers
                    const normalize = (str: string): string => {
                        return str.toLowerCase()
                            .replace(/[^\w\s]/g, '') // Remove punctuation
                            .replace(/s$/, '') // Remove trailing 's'
                            .trim();
                    };

                    // Find description column in header row
                    for (let col = topLeftCellRef.c; col <= range.e.c; col++) {
                        const headerAddress = XLSX.utils.encode_cell({ r: topLeftCellRef.r, c: col });
                        const headerCell = worksheet[headerAddress];
                        if (headerCell && headerCell.v) {
                            const normalizedValue = normalize(String(headerCell.v));
                            if (normalizedValue === 'description') {
                                descriptionColumn = col;
                                break;
                            }
                        }
                    }

                    // Count data rows
                    if (descriptionColumn !== -1) {
                        for (let dataRow = topLeftCellRef.r + 1; dataRow <= range.e.r; dataRow++) {
                            const descAddress = XLSX.utils.encode_cell({ r: dataRow, c: descriptionColumn });
                            const descCell = worksheet[descAddress];
                            const hasDescription = descCell && descCell.v !== null && descCell.v !== undefined && String(descCell.v).trim() !== '';
                            
                            if (hasDescription) {
                                rowCount++;
                            } else {
                                // Stop counting if we hit consecutive empty rows
                                let consecutiveEmptyRows = 0;
                                for (let checkRow = dataRow; checkRow <= Math.min(dataRow + 2, range.e.r); checkRow++) {
                                    const checkAddress = XLSX.utils.encode_cell({ r: checkRow, c: descriptionColumn });
                                    const checkCell = worksheet[checkAddress];
                                    if (!checkCell || !checkCell.v || String(checkCell.v).trim() === '') {
                                        consecutiveEmptyRows++;
                                    } else {
                                        break;
                                    }
                                }
                                if (consecutiveEmptyRows >= 3) {
                                    break;
                                }
                            }
                        }
                    } else {
                        // Fallback: count rows with any data
                        for (let dataRow = topLeftCellRef.r + 1; dataRow <= range.e.r; dataRow++) {
                            let hasData = false;
                            for (let col = range.s.c; col <= range.e.c; col++) {
                                const cellAddress = XLSX.utils.encode_cell({ r: dataRow, c: col });
                                const cell = worksheet[cellAddress];
                                if (cell && cell.v !== null && cell.v !== undefined && String(cell.v).trim() !== '') {
                                    hasData = true;
                                    break;
                                }
                            }
                            if (hasData) {
                                rowCount++;
                            } else {
                                break;
                            }
                        }
                    }

                    const fileName = file.name;

                    resolve({
                        fileName,
                        file,
                        rowCount,
                        topLeftCell,
                        category,
                        discount: 0
                    });
                } catch (error) {
                    reject(error);
                }
            };

            reader.onerror = () => {
                reject(new Error('File reading error'));
            };

            reader.readAsArrayBuffer(file);
        });
    }

    private findFirstDataCell(worksheet: XLSX.WorkSheet, range: XLSX.Range): string | undefined {
        // Required column headers: 'Pos', 'Description', 'Remark', 'Unit', 'Qty', 'Price', 'Total'
        // Ignore punctuation and trailing 's'
        const requiredHeaders = ['pos', 'description', 'remark', 'unit', 'qty', 'price', 'total'];
        const maxRowsToScan = Math.min(range.e.r, 50);
        const maxColsToScan = Math.min(range.e.c, 25);

        // Normalize a string by removing punctuation and trailing 's'
        const normalize = (str: string): string => {
            return str.toLowerCase()
                .replace(/[^\w\s]/g, '') // Remove punctuation
                .replace(/s$/, '') // Remove trailing 's'
                .trim();
        };

        for (let row = range.s.r; row <= maxRowsToScan; row++) {
            // Check if this row contains the required headers
            const headersInRow: string[] = [];
            for (let col = range.s.c; col <= maxColsToScan; col++) {
                const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
                const cell = worksheet[cellAddress];
                if (cell && cell.v !== undefined && cell.v !== null) {
                    const normalizedValue = normalize(String(cell.v));
                    if (normalizedValue && requiredHeaders.includes(normalizedValue)) {
                        headersInRow.push(normalizedValue);
                    }
                }
            }

            // If we found at least one required header in this row, this is likely the header row
            // Return the first cell of this row (leftmost column)
            if (headersInRow.length > 0) {
                return XLSX.utils.encode_cell({ r: row, c: range.s.c });
            }
        }

        return undefined;
    }

    triggerFileInput(category: 'Invoice' | 'Supplier Quotations', setIndex: number): void {
        this.loggingService.logButtonClick('file_input_triggered', 'SupplierAnalysisInputsComponent', {
            category: category,
            setIndex: setIndex
        });

        const fileInput = document.getElementById(`fileInput-${category}-${setIndex}`) as HTMLInputElement;
        fileInput?.click();
    }

    getTopLeftCellOptions(): string[] {
        const options: string[] = [];
        for (let col = 0; col < 8; col++) {
            for (let row = 1; row <= 20; row++) {
                const cellRef = XLSX.utils.encode_cell({ r: row - 1, c: col });
                options.push(cellRef);
            }
        }
        return options;
    }

    onTopLeftCellChange(event: Event, fileInfo: SupplierAnalysisFileInfo, setIndex: number): void {
        const input = event.target as HTMLInputElement;
        let value = input.value.toUpperCase();

        const regex = /^[A-Z][0-9]{1,2}$/;
        if (value && !regex.test(value)) {
            value = value.replace(/[^A-Z0-9]/g, '');
            if (value && !/^[A-Z]/.test(value)) {
                value = '';
            }
            const match = value.match(/^([A-Z])([0-9]{0,2})/);
            if (match) {
                value = match[0];
            } else if (value.length > 0 && /^[A-Z]/.test(value)) {
                value = value.charAt(0);
            }
        }

        input.value = value;
        fileInfo.topLeftCell = value;

        if (value && value.trim() !== '' && value !== 'NOT FOUND') {
            input.removeAttribute('list');
            void this.reanalyzeFileWithTopLeft(fileInfo, value, setIndex);
        } else {
            fileInfo.rowCount = 0;
        }
    }

    validateTopLeftCell(fileInfo: SupplierAnalysisFileInfo): void {
        const regex = /^[A-Z][0-9]{1,2}$/;
        if (fileInfo.topLeftCell && fileInfo.topLeftCell !== 'NOT FOUND' && !regex.test(fileInfo.topLeftCell)) {
            fileInfo.topLeftCell = 'A1';
            fileInfo.rowCount = 0;
        }
    }

    private async reanalyzeFileWithTopLeft(fileInfo: SupplierAnalysisFileInfo, topLeftCell: string, setIndex: number): Promise<void> {
        this.isProcessing = true;
        try {
            return new Promise((resolve) => {
                const reader = new FileReader();

                reader.onload = (e: any) => {
                    try {
                        const data = new Uint8Array(e.target.result);
                        const workbook = XLSX.read(data, { type: 'array' });
                        const firstSheetName = workbook.SheetNames[0];
                        const worksheet = workbook.Sheets[firstSheetName];

                        const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1:Z100');
                        let topLeftCellRef: XLSX.CellAddress;
                        
                        try {
                            topLeftCellRef = XLSX.utils.decode_cell(topLeftCell);
                        } catch {
                            topLeftCellRef = XLSX.utils.decode_cell('A1');
                        }

                        // Count rows starting from the row after topLeftCell
                        let rowCount = 0;
                        let descriptionColumn = -1;

                        // Normalize function to match headers
                        const normalize = (str: string): string => {
                            return str.toLowerCase()
                                .replace(/[^\w\s]/g, '') // Remove punctuation
                                .replace(/s$/, '') // Remove trailing 's'
                                .trim();
                        };

                        // Find description column in header row
                        for (let col = topLeftCellRef.c; col <= range.e.c; col++) {
                            const headerAddress = XLSX.utils.encode_cell({ r: topLeftCellRef.r, c: col });
                            const headerCell = worksheet[headerAddress];
                            if (headerCell && headerCell.v) {
                                const normalizedValue = normalize(String(headerCell.v));
                                if (normalizedValue === 'description') {
                                    descriptionColumn = col;
                                    break;
                                }
                            }
                        }

                        // Count data rows
                        if (descriptionColumn !== -1) {
                            for (let dataRow = topLeftCellRef.r + 1; dataRow <= range.e.r; dataRow++) {
                                const descAddress = XLSX.utils.encode_cell({ r: dataRow, c: descriptionColumn });
                                const descCell = worksheet[descAddress];
                                const hasDescription = descCell && descCell.v !== null && descCell.v !== undefined && String(descCell.v).trim() !== '';
                                
                                if (hasDescription) {
                                    rowCount++;
                                } else {
                                    // Stop counting if we hit consecutive empty rows
                                    let consecutiveEmptyRows = 0;
                                    for (let checkRow = dataRow; checkRow <= Math.min(dataRow + 2, range.e.r); checkRow++) {
                                        const checkAddress = XLSX.utils.encode_cell({ r: checkRow, c: descriptionColumn });
                                        const checkCell = worksheet[checkAddress];
                                        if (!checkCell || !checkCell.v || String(checkCell.v).trim() === '') {
                                            consecutiveEmptyRows++;
                                        } else {
                                            break;
                                        }
                                    }
                                    if (consecutiveEmptyRows >= 3) {
                                        break;
                                    }
                                }
                            }
                        } else {
                            // Fallback: count rows with any data
                            for (let dataRow = topLeftCellRef.r + 1; dataRow <= range.e.r; dataRow++) {
                                let hasData = false;
                                for (let col = range.s.c; col <= range.e.c; col++) {
                                    const cellAddress = XLSX.utils.encode_cell({ r: dataRow, c: col });
                                    const cell = worksheet[cellAddress];
                                    if (cell && cell.v !== null && cell.v !== undefined && String(cell.v).trim() !== '') {
                                        hasData = true;
                                        break;
                                    }
                                }
                                if (hasData) {
                                    rowCount++;
                                } else {
                                    break;
                                }
                            }
                        }

                        fileInfo.topLeftCell = topLeftCell;
                        fileInfo.rowCount = rowCount;

                        // Update service
                        this.updateServiceFileSet(this.dropzoneSets[setIndex]);

                        this.loggingService.logUserAction('top_left_cell_changed', {
                            fileName: fileInfo.fileName,
                            newTopLeftCell: topLeftCell,
                            rowCount: rowCount
                        }, 'SupplierAnalysisInputsComponent');

                        resolve();
                    } catch (error) {
                        this.loggingService.logError(error as Error, 'reanalyze_file_error', 'SupplierAnalysisInputsComponent', {
                            fileName: fileInfo.fileName,
                            topLeftCell: topLeftCell
                        });
                        resolve();
                    }
                };

                reader.onerror = () => {
                    this.loggingService.logError(new Error('File reading error'), 'file_reader_error', 'SupplierAnalysisInputsComponent');
                    resolve();
                };

                reader.readAsArrayBuffer(fileInfo.file);
            });
        } finally {
            this.isProcessing = false;
        }
    }

    removeFile(fileName: string, category: 'Invoice' | 'Supplier Quotations', setIndex: number): void {
        const set = this.dropzoneSets[setIndex];
        if (category === 'Invoice') {
            set.invoiceFiles = set.invoiceFiles.filter(f => f.fileName !== fileName);
        } else {
            set.supplierQuotationFiles = set.supplierQuotationFiles.filter(f => f.fileName !== fileName);
        }
        
        this.updateServiceFileSet(set);
        
        // Expand set if there's a mismatch after removal
        if (this.hasAnyMismatch(set)) {
            set.expanded = true;
        }
    }

    removeAllFiles(setIndex: number): void {
        this.loggingService.logButtonClick('remove_all_files', 'SupplierAnalysisInputsComponent', {
            setIndex: setIndex
        });
        
        const set = this.dropzoneSets[setIndex];
        set.invoiceFiles = [];
        set.supplierQuotationFiles = [];
        this.updateServiceFileSet(set);
    }

    hasMatchingRowCount(file: SupplierAnalysisFileInfo, set: InputDropzoneSet): boolean {
        if (file.category !== 'Supplier Quotations') {
            return true;
        }
        if (set.invoiceFiles.length === 0) {
            return true;
        }
        const invoiceRowCount = set.invoiceFiles[0].rowCount;
        return file.rowCount === invoiceRowCount;
    }
    
    hasAnyMismatch(set: InputDropzoneSet): boolean {
        return this.getAllFiles(set).some(file => !this.hasMatchingRowCount(file, set));
    }

    private updateServiceFileSet(set: InputDropzoneSet): void {
        const allFiles = this.getAllFiles(set);
        this.supplierAnalysisService.updateFileSet(set.id, allFiles);
    }

    getAllFiles(set: InputDropzoneSet): SupplierAnalysisFileInfo[] {
        // Always return Invoice files first, then Supplier Quotation files
        const allFiles = [...set.invoiceFiles, ...set.supplierQuotationFiles];
        return allFiles.sort((a, b) => {
            if (a.category === 'Invoice' && b.category === 'Supplier Quotations') return -1;
            if (a.category === 'Supplier Quotations' && b.category === 'Invoice') return 1;
            return 0;
        });
    }

    async openFile(fileInfo: SupplierAnalysisFileInfo): Promise<void> {
        this.loggingService.logUserAction('file_opened', {
            fileName: fileInfo.fileName,
            fileSize: fileInfo.file.size,
            category: fileInfo.category
        }, 'SupplierAnalysisInputsComponent');

        const url = URL.createObjectURL(fileInfo.file);
        window.open(url, '_blank');

        setTimeout(() => {
            URL.revokeObjectURL(url);
        }, 1000);
    }

    hasInvoiceFiles(set: InputDropzoneSet): boolean {
        return set.invoiceFiles.length > 0;
    }

    hasSupplierQuotationFiles(set: InputDropzoneSet): boolean {
        return set.supplierQuotationFiles.length > 0;
    }

    getInvoiceLastWord(set: InputDropzoneSet): string {
        if (set.invoiceFiles.length === 0) {
            return '';
        }
        const fileName = set.invoiceFiles[0].fileName;
        const nameWithoutExt = fileName.replace(/\.[^/.]+$/, '');
        const words = nameWithoutExt.split(/[\s\-_]+/);
        
        if (words.length === 0) {
            return fileName;
        }
        
        if (words.length === 1) {
            return words[0];
        }
        
        return `${words[0]} ${words[1]}`;
    }

    getSupplierQuotationCount(set: InputDropzoneSet): number {
        return set.supplierQuotationFiles.length;
    }

    toggleInstructions(): void {
        this.instructionsExpanded = !this.instructionsExpanded;
    }

    toggleSet(set: InputDropzoneSet): void {
        set.expanded = !set.expanded;
    }

    getDiscountPercentage(discount: number | undefined): string {
        if (discount === undefined || discount === null) return '0';
        const percentage = (1 - discount) * 100;
        return percentage.toFixed(0);
    }

    onFileDragStart(event: DragEvent, file: SupplierAnalysisFileInfo, setIndex: number): void {
        if (file.category !== 'Supplier Quotations') {
            event.preventDefault();
            return;
        }
        
        this.isDraggingFile = true;
        this.draggedFile = file;
        
        if (event.dataTransfer) {
            event.dataTransfer.effectAllowed = 'move';
            event.dataTransfer.setData('text/plain', file.fileName);
        }

        this.loggingService.logUserAction('file_drag_start', {
            fileName: file.fileName,
            setIndex: setIndex
        }, 'SupplierAnalysisInputsComponent');
    }

    onFileDragEnd(event: DragEvent): void {
        this.isDraggingFile = false;
        this.draggedFile = null;
        this.dragOverFile = null;
    }

    onFileDragOver(event: DragEvent, file: SupplierAnalysisFileInfo, setIndex: number): void {
        if (!this.isDraggingFile || !this.draggedFile || file.category !== 'Supplier Quotations') {
            return;
        }

        event.preventDefault();
        
        if (event.dataTransfer) {
            event.dataTransfer.dropEffect = 'move';
        }

        this.dragOverFile = file;
    }

    onFileDrop(event: DragEvent, targetFile: SupplierAnalysisFileInfo, setIndex: number): void {
        event.preventDefault();
        event.stopPropagation();

        if (!this.draggedFile || !targetFile || targetFile.category !== 'Supplier Quotations') {
            this.onFileDragEnd(event);
            return;
        }

        if (this.draggedFile === targetFile) {
            this.onFileDragEnd(event);
            return;
        }

        const set = this.dropzoneSets[setIndex];
        const draggedIndex = set.supplierQuotationFiles.findIndex(f => f === this.draggedFile);
        const targetIndex = set.supplierQuotationFiles.findIndex(f => f === targetFile);

        if (draggedIndex === -1 || targetIndex === -1) {
            this.onFileDragEnd(event);
            return;
        }

        // Reorder the array
        const [removed] = set.supplierQuotationFiles.splice(draggedIndex, 1);
        set.supplierQuotationFiles.splice(targetIndex, 0, removed);

        // Update service with new order
        this.updateServiceFileSet(set);

        this.loggingService.logUserAction('file_reordered', {
            fileName: this.draggedFile.fileName,
            fromIndex: draggedIndex,
            toIndex: targetIndex,
            setIndex: setIndex
        }, 'SupplierAnalysisInputsComponent');

        this.onFileDragEnd(event);
    }
}
