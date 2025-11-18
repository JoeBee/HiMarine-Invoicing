import { Component, OnInit, OnDestroy } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { LoggingService, LogEntry } from '../../services/logging.service';

@Component({
    selector: 'app-history',
    standalone: true,
    imports: [CommonModule, FormsModule],
    templateUrl: './history.component.html',
    styleUrls: ['./history.component.scss']
})
export class HistoryComponent implements OnInit, OnDestroy {
    logs: LogEntry[] = [];
    filteredLogs: LogEntry[] = [];
    isLoading = false;
    error: string | null = null;

    // Filter properties
    selectedCategory = '';
    selectedLevel = '';
    selectedComponent = '';
    selectedIpAddress = '';
    selectedTimezone = '';
    selectedLanguage = '';
    selectedAdmin = '';
    searchText = '';
    dateRange = {
        start: '',
        end: ''
    };

    // Available filter options
    categories: string[] = [];
    levels: string[] = ['info', 'warn', 'error', 'debug'];
    components: string[] = [];
    ipAddresses: string[] = [];
    timezones: string[] = [];
    languages: string[] = [];

    // Pagination
    currentPage = 1;
    itemsPerPage = 50;
    totalPages = 1;

    // User IP address
    userIpAddress: string = 'Loading...';
    private ipUpdateInterval: any;

    constructor(private loggingService: LoggingService) { }

    ngOnInit(): void {
        this.loadLogs();
        
        // Get user IP address
        this.updateUserIpAddress();
        // Update IP address periodically in case it's still loading
        this.ipUpdateInterval = setInterval(() => {
            this.updateUserIpAddress();
        }, 2000);
    }

    private updateUserIpAddress(): void {
        this.userIpAddress = this.loggingService.getIpAddress();
    }

    private async loadLogs(): Promise<void> {
        this.isLoading = true;
        this.error = null;

        try {
            this.logs = await this.loggingService.getLogs(1000); // Load up to 1000 logs
            this.updateFilterOptions();
            this.applyFilters();
            this.updatePagination();
        } catch (error) {
            this.error = 'Failed to load logs. Please try again.';
            this.loggingService.logError(
                error as Error,
                'log_loading_failure',
                'HistoryComponent',
                {
                    attempted_limit: 1000,
                    current_logs_count: this.logs.length,
                    filters: {
                        category: this.selectedCategory,
                        level: this.selectedLevel,
                        component: this.selectedComponent,
                        searchText: this.searchText,
                        dateRange: this.dateRange
                    }
                }
            );
        } finally {
            this.isLoading = false;
        }
    }

    private updateFilterOptions(): void {
        // Extract unique categories
        this.categories = [...new Set(this.logs.map(log => log.category))].sort();

        // Extract unique components
        this.components = [...new Set(this.logs.map(log => log.component))].sort();

        // Extract unique IP addresses, ordered by most recent first
        const ipMap = new Map<string, Date>();
        this.logs.forEach(log => {
            if (log.ipAddress && log.ipAddress !== 'Unknown') {
                const existingDate = ipMap.get(log.ipAddress);
                if (!existingDate || log.timestamp > existingDate) {
                    ipMap.set(log.ipAddress, log.timestamp);
                }
            }
        });

        // Sort by most recent timestamp first, then by IP address
        this.ipAddresses = Array.from(ipMap.entries())
            .sort((a, b) => {
                // First sort by timestamp (most recent first)
                if (b[1].getTime() !== a[1].getTime()) {
                    return b[1].getTime() - a[1].getTime();
                }
                // Then by IP address if timestamps are equal
                return a[0].localeCompare(b[0]);
            })
            .map(entry => entry[0]);

        // Extract unique timezones, ordered by most recent first
        const timezoneMap = new Map<string, Date>();
        this.logs.forEach(log => {
            if (log.timezone && log.timezone !== 'Unknown') {
                const existingDate = timezoneMap.get(log.timezone);
                if (!existingDate || log.timestamp > existingDate) {
                    timezoneMap.set(log.timezone, log.timestamp);
                }
            }
        });

        this.timezones = Array.from(timezoneMap.entries())
            .sort((a, b) => {
                if (b[1].getTime() !== a[1].getTime()) {
                    return b[1].getTime() - a[1].getTime();
                }
                return a[0].localeCompare(b[0]);
            })
            .map(entry => entry[0]);

        // Extract unique languages, ordered by most recent first
        const languageMap = new Map<string, Date>();
        this.logs.forEach(log => {
            if (log.language && log.language !== 'Unknown') {
                const existingDate = languageMap.get(log.language);
                if (!existingDate || log.timestamp > existingDate) {
                    languageMap.set(log.language, log.timestamp);
                }
            }
        });

        this.languages = Array.from(languageMap.entries())
            .sort((a, b) => {
                if (b[1].getTime() !== a[1].getTime()) {
                    return b[1].getTime() - a[1].getTime();
                }
                return a[0].localeCompare(b[0]);
            })
            .map(entry => entry[0]);
    }

    onFilterChange(): void {
        this.applyFilters();
        this.currentPage = 1; // Reset to first page when filters change
    }

    applyFilters(): void {
        let filtered = [...this.logs];

        // Filter by category
        if (this.selectedCategory) {
            filtered = filtered.filter(log => log.category === this.selectedCategory);
        }

        // Filter by level
        if (this.selectedLevel) {
            filtered = filtered.filter(log => log.level === this.selectedLevel);
        }

        // Filter by component
        if (this.selectedComponent) {
            filtered = filtered.filter(log => log.component === this.selectedComponent);
        }

        // Filter by IP address
        if (this.selectedIpAddress) {
            filtered = filtered.filter(log => log.ipAddress === this.selectedIpAddress);
        }

        // Filter by timezone
        if (this.selectedTimezone) {
            filtered = filtered.filter(log => log.timezone === this.selectedTimezone);
        }

        // Filter by language
        if (this.selectedLanguage) {
            filtered = filtered.filter(log => log.language === this.selectedLanguage);
        }

        // Filter by admin
        if (this.selectedAdmin) {
            const isAdminValue = this.selectedAdmin === 'true';
            filtered = filtered.filter(log => log.isAdmin === isAdminValue);
        }

        // Filter by search text
        if (this.searchText.trim()) {
            const searchLower = this.searchText.toLowerCase();
            filtered = filtered.filter(log =>
                log.action.toLowerCase().includes(searchLower) ||
                log.details?.toString().toLowerCase().includes(searchLower) ||
                log.component.toLowerCase().includes(searchLower)
            );
        }

        // Filter by date range
        if (this.dateRange.start) {
            const startDate = new Date(this.dateRange.start);
            filtered = filtered.filter(log => log.timestamp >= startDate);
        }

        if (this.dateRange.end) {
            const endDate = new Date(this.dateRange.end);
            endDate.setHours(23, 59, 59, 999); // Include entire end date
            filtered = filtered.filter(log => log.timestamp <= endDate);
        }

        this.filteredLogs = filtered;
        this.updatePagination();
    }

    clearFilters(): void {
        this.selectedCategory = '';
        this.selectedLevel = '';
        this.selectedComponent = '';
        this.selectedIpAddress = '';
        this.selectedTimezone = '';
        this.selectedLanguage = '';
        this.selectedAdmin = '';
        this.searchText = '';
        this.dateRange = { start: '', end: '' };
        this.applyFilters();
    }

    private updatePagination(): void {
        this.totalPages = Math.ceil(this.filteredLogs.length / this.itemsPerPage);
        if (this.currentPage > this.totalPages) {
            this.currentPage = 1;
        }
    }

    get paginatedLogs(): LogEntry[] {
        const startIndex = (this.currentPage - 1) * this.itemsPerPage;
        const endIndex = startIndex + this.itemsPerPage;
        return this.filteredLogs.slice(startIndex, endIndex);
    }

    onPageChange(page: number): void {
        this.currentPage = page;
    }

    onPageJump(inputPage: string): void {
        const pageNumber = parseInt(inputPage, 10);
        if (!isNaN(pageNumber) && pageNumber >= 1 && pageNumber <= this.totalPages) {
            this.currentPage = pageNumber;
        }
    }

    getLevelClass(level: string): string {
        switch (level) {
            case 'error': return 'log-error';
            case 'warn': return 'log-warn';
            case 'info': return 'log-info';
            case 'debug': return 'log-debug';
            default: return 'log-info';
        }
    }

    getCategoryIcon(category: string): string {
        switch (category) {
            case 'user_action': return '👤';
            case 'file_upload': return '📁';
            case 'data_processing': return '⚙️';
            case 'export': return '📤';
            case 'error': return '❌';
            case 'system': return '🔧';
            default: return '📝';
        }
    }

    formatTimestamp(timestamp: Date): string {
        return new Date(timestamp).toLocaleString();
    }

    formatDetails(details: any): string {
        if (typeof details === 'string') {
            return details;
        }
        if (typeof details === 'object' && details !== null) {
            // Format as key/value pairs instead of JSON
            return Object.entries(details)
                .map(([key, value]) => `${key}: ${value}`)
                .join('\n');
        }
        return String(details);
    }

    formatAction(action: string, category: string, details: any): string {
        // Make action more specific based on category and details
        switch (category) {
            case 'user_action':
                if (action === 'button_click') {
                    if (details && details.buttonName) {
                        return details.buttonName.replace(/_/g, ' ');
                    }
                    // If button_click action but no buttonName, return empty string to avoid "button click" text
                    return '';
                }
                if (action.includes('button') && action !== 'button_click') {
                    return action.replace('button_click_', '').replace(/_/g, ' ');
                }
                if (action.includes('form')) {
                    return `Form submitted: ${action.replace('form_submit_', '').replace(/_/g, ' ')}`;
                }
                return action.replace(/_/g, ' ');

            case 'file_upload':
                if (details && details.fileName) {
                    return `File uploaded: ${details.fileName}`;
                }
                return `File upload: ${action.replace(/_/g, ' ')}`;

            case 'data_processing':
                if (details && details.operation) {
                    return `Data processed: ${details.operation}`;
                }
                return `Data processing: ${action.replace(/_/g, ' ')}`;

            case 'export':
                if (details && details.format) {
                    return `Export generated: ${details.format.toUpperCase()}`;
                }
                return `Export: ${action.replace(/_/g, ' ')}`;

            case 'error':
                if (details && details.error) {
                    return `Error occurred: ${details.error}`;
                }
                return `Error: ${action.replace(/_/g, ' ')}`;

            case 'system':
                if (details && details.event) {
                    return `System event: ${details.event}`;
                }
                return `System: ${action.replace(/_/g, ' ')}`;

            default:
                return action.replace(/_/g, ' ');
        }
    }

    refreshLogs(): void {
        this.loadLogs();
    }

    exportLogs(): void {
        // Create CSV content
        const headers = ['Timestamp', 'Category', 'Action', 'IP Address', 'Timezone', 'Language', 'Admin', 'Details'];
        const csvContent = [
            headers.join(','),
            ...this.filteredLogs.map(log => [
                this.formatTimestamp(log.timestamp),
                log.category,
                `"${log.component} - ${this.formatAction(log.action, log.category, log.details).replace(/"/g, '""')}"`,
                log.ipAddress || 'Unknown',
                log.timezone || 'Unknown',
                log.language || 'Unknown',
                log.isAdmin ? 'true' : 'false',
                `"${this.formatDetails(log.details).replace(/"/g, '""')}"`
            ].join(','))
        ].join('\n');

        // Create and download file
        const blob = new Blob([csvContent], { type: 'text/csv' });
        const url = window.URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = `application_logs_${new Date().toISOString().split('T')[0]}.csv`;
        link.click();
        window.URL.revokeObjectURL(url);
    }

    formatIpAddress(ipAddress?: string): string {
        if (!ipAddress || ipAddress === 'Unknown') {
            return '🌐 Unknown';
        }
        return `🌐 ${ipAddress}`;
    }

    formatTimezone(timezone?: string): string {
        if (!timezone || timezone === 'Unknown') {
            return '🕐 Unknown';
        }
        // Format timezone to be more readable (e.g., "America/New_York" -> "New York")
        const parts = timezone.split('/');
        if (parts.length > 1) {
            return `🕐 ${parts[parts.length - 1].replace(/_/g, ' ')}`;
        }
        return `🕐 ${timezone}`;
    }

    formatLanguage(language?: string): string {
        if (!language || language === 'Unknown') {
            return '🌍 Unknown';
        }
        // Format language code (e.g., "en-US" -> "🇺🇸 en-US" or just show the language)
        return `🌍 ${language}`;
    }

    ngOnDestroy(): void {
        // Clear IP update interval
        if (this.ipUpdateInterval) {
            clearInterval(this.ipUpdateInterval);
        }
    }
}

