import { Injectable } from '@angular/core';
import { HttpClient } from '@angular/common/http';
import { collection, addDoc, query, orderBy, limit, getDocs, where, Timestamp } from 'firebase/firestore';
import { db } from '../firebase.config';
import { catchError, timeout } from 'rxjs/operators';
import { of, Observable } from 'rxjs';

export interface LogEntry {
    id?: string;
    timestamp: Date;
    level: 'info' | 'warn' | 'error' | 'debug';
    category: 'user_action' | 'file_upload' | 'data_processing' | 'export' | 'error' | 'system';
    action: string;
    details: any;
    userId?: string;
    sessionId: string;
    component: string;
    url: string;
    userAgent: string;
    ipAddress?: string;
    timezone?: string;
    language?: string;
}

@Injectable({
    providedIn: 'root'
})
export class LoggingService {
    private sessionId: string;
    private logQueue: LogEntry[] = [];
    private isOnline: boolean = navigator.onLine;
    private batchSize: number = 10;
    private batchTimeout: number = 5000; // 5 seconds
    private retentionDays: number = 7; // Logs are retained for 1 week
    private batchTimer: any;
    private cachedIpAddress: string | null = null;
    private ipFetchInProgress: boolean = false;

    constructor(private http: HttpClient) {
        this.sessionId = this.generateSessionId();
        this.setupOnlineListener();
        this.startBatchProcessor();
        this.fetchIpAddress(); // Pre-fetch IP address on service initialization
    }

    private generateSessionId(): string {
        return 'session_' + Date.now() + '_' + Math.random().toString(36).substr(2, 9);
    }

    private setupOnlineListener(): void {
        window.addEventListener('online', () => {
            this.isOnline = true;
            this.processLogQueue();
        });

        window.addEventListener('offline', () => {
            this.isOnline = false;
        });
    }

    private startBatchProcessor(): void {
        this.batchTimer = setInterval(() => {
            if (this.logQueue.length > 0 && this.isOnline) {
                this.processLogQueue();
            }
        }, this.batchTimeout);
    }

    private async processLogQueue(): Promise<void> {
        if (this.logQueue.length === 0 || !this.isOnline) {
            return;
        }

        const logsToProcess = this.logQueue.splice(0, this.batchSize);

        try {
            const logCollection = collection(db, 'application_logs');
            const batch = logsToProcess.map(log => addDoc(logCollection, {
                ...log,
                timestamp: Timestamp.fromDate(log.timestamp)
            }));

            await Promise.all(batch);
            console.log(`Successfully logged ${logsToProcess.length} entries to Firestore`);
        } catch (error) {
            console.error('Error logging to Firestore:', error);
            // Re-add failed logs to the queue for retry
            this.logQueue.unshift(...logsToProcess);
        }
    }

    private createLogEntry(
        level: LogEntry['level'],
        category: LogEntry['category'],
        action: string,
        details: any,
        component: string
    ): LogEntry {
        // Capture timezone
        let timezone: string | undefined;
        try {
            timezone = Intl.DateTimeFormat().resolvedOptions().timeZone;
        } catch (e) {
            timezone = 'Unknown';
        }

        // Capture language
        const language = navigator.language || 'Unknown';

        return {
            timestamp: new Date(),
            level,
            category,
            action,
            details,
            sessionId: this.sessionId,
            component,
            url: window.location.href,
            userAgent: navigator.userAgent,
            ipAddress: this.cachedIpAddress || 'Unknown',
            timezone: timezone,
            language: language
        };
    }

    private fetchIpAddress(): void {
        if (this.ipFetchInProgress || this.cachedIpAddress) {
            return;
        }

        this.ipFetchInProgress = true;

        // Use ipify API to get public IP address
        this.http.get<{ ip: string }>('https://api.ipify.org?format=json')
            .pipe(
                timeout(5000), // 5 second timeout
                catchError(() => {
                    // Fallback to another service if ipify fails
                    return this.http.get<{ ip: string }>('https://httpbin.org/ip')
                        .pipe(
                            timeout(5000),
                            catchError(() => {
                                // Final fallback - return observable with 'Unknown'
                                return of({ ip: 'Unknown' });
                            })
                        );
                })
            )
            .subscribe({
                next: (response) => {
                    this.cachedIpAddress = response.ip;
                    this.ipFetchInProgress = false;
                },
                error: () => {
                    this.cachedIpAddress = 'Unknown';
                    this.ipFetchInProgress = false;
                }
            });
    }

    private refreshIpAddressIfNeeded(): void {
        // Refresh IP address every hour or if it's unknown
        if (!this.cachedIpAddress || this.cachedIpAddress === 'Unknown') {
            this.fetchIpAddress();
        }
    }

    // Public logging methods
    logUserAction(action: string, details: any, component: string): void {
        this.refreshIpAddressIfNeeded();
        const logEntry = this.createLogEntry('info', 'user_action', action, details, component);
        this.addToQueue(logEntry);
    }

    logFileUpload(fileName: string, fileSize: number, fileType: string, category: string, component: string): void {
        this.refreshIpAddressIfNeeded();
        const logEntry = this.createLogEntry('info', 'file_upload', 'file_uploaded', {
            fileName,
            fileSize,
            fileType,
            category
        }, component);
        this.addToQueue(logEntry);
    }

    logDataProcessing(action: string, details: any, component: string): void {
        this.refreshIpAddressIfNeeded();
        const logEntry = this.createLogEntry('info', 'data_processing', action, details, component);
        this.addToQueue(logEntry);
    }

    logExport(action: string, details: any, component: string): void {
        this.refreshIpAddressIfNeeded();
        const logEntry = this.createLogEntry('info', 'export', action, details, component);
        this.addToQueue(logEntry);
    }


    logError(error: Error | string, context: string, component: string, additionalDetails?: any): void {
        this.refreshIpAddressIfNeeded();
        let errorMessage: string;
        let errorStack: string | undefined;
        let errorName: string | undefined;

        if (error instanceof Error) {
            errorMessage = error.message;
            errorStack = error.stack;
            errorName = error.name;
        } else {
            errorMessage = error;
            errorStack = undefined;
            errorName = 'StringError';
        }

        const logEntry = this.createLogEntry('error', 'error', 'error_occurred', {
            errorMessage,
            errorStack,
            errorName,
            context,
            timestamp: new Date().toISOString(),
            userAgent: navigator.userAgent,
            url: window.location.href,
            screenResolution: `${screen.width}x${screen.height}`,
            windowSize: `${window.innerWidth}x${window.innerHeight}`,
            language: navigator.language,
            platform: navigator.platform,
            cookieEnabled: navigator.cookieEnabled,
            onLine: navigator.onLine,
            additionalDetails
        }, component);

        // Also log to console for immediate debugging
        console.error(`[${component}] ${context}:`, {
            error: error instanceof Error ? error : new Error(error),
            additionalDetails,
            timestamp: new Date().toISOString()
        });

        this.addToQueue(logEntry);
    }

    logSystemEvent(action: string, details: any, component: string): void {
        this.refreshIpAddressIfNeeded();
        const logEntry = this.createLogEntry('info', 'system', action, details, component);
        this.addToQueue(logEntry);
    }

    logButtonClick(buttonName: string, component: string, additionalDetails?: any): void {
        this.refreshIpAddressIfNeeded();
        const logEntry = this.createLogEntry('info', 'user_action', 'button_click', {
            buttonName,
            ...additionalDetails
        }, component);
        this.addToQueue(logEntry);
    }

    logFormSubmission(formName: string, formData: any, component: string): void {
        this.refreshIpAddressIfNeeded();
        const logEntry = this.createLogEntry('info', 'user_action', 'form_submission', {
            formName,
            formData: this.sanitizeFormData(formData)
        }, component);
        this.addToQueue(logEntry);
    }

    logFilterChange(filterType: string, filterValue: any, component: string): void {
        this.refreshIpAddressIfNeeded();
        const logEntry = this.createLogEntry('info', 'user_action', 'filter_change', {
            filterType,
            filterValue
        }, component);
        this.addToQueue(logEntry);
    }

    logSortChange(column: string, direction: string, component: string): void {
        this.refreshIpAddressIfNeeded();
        const logEntry = this.createLogEntry('info', 'user_action', 'sort_change', {
            column,
            direction
        }, component);
        this.addToQueue(logEntry);
    }

    logDataSelection(selectionType: string, selectedCount: number, totalCount: number, component: string): void {
        this.refreshIpAddressIfNeeded();
        const logEntry = this.createLogEntry('info', 'user_action', 'data_selection', {
            selectionType,
            selectedCount,
            totalCount
        }, component);
        this.addToQueue(logEntry);
    }

    private addToQueue(logEntry: LogEntry): void {
        this.logQueue.push(logEntry);

        // Process immediately if batch is full
        if (this.logQueue.length >= this.batchSize) {
            this.processLogQueue();
        }
    }

    private sanitizeFormData(formData: any): any {
        // Remove sensitive information from form data
        const sanitized = { ...formData };
        const sensitiveFields = ['password', 'token', 'secret', 'key'];

        sensitiveFields.forEach(field => {
            if (sanitized[field]) {
                sanitized[field] = '[REDACTED]';
            }
        });

        return sanitized;
    }

    // Method to get logs for display in History component
    async getLogs(limitCount: number = 100): Promise<LogEntry[]> {
        try {
            const logCollection = collection(db, 'application_logs');
            const q = query(
                logCollection,
                orderBy('timestamp', 'desc'),
                limit(limitCount)
            );

            const querySnapshot = await getDocs(q);
            const logs: LogEntry[] = [];

            querySnapshot.forEach((doc) => {
                const data = doc.data();
                logs.push({
                    id: doc.id,
                    timestamp: data['timestamp'].toDate(),
                    level: data['level'],
                    category: data['category'],
                    action: data['action'],
                    details: data['details'],
                    userId: data['userId'],
                    sessionId: data['sessionId'],
                    component: data['component'],
                    url: data['url'],
                    userAgent: data['userAgent'],
                    ipAddress: data['ipAddress'],
                    timezone: data['timezone'],
                    language: data['language']
                });
            });

            return logs;
        } catch (error) {
            console.error('Error fetching logs:', error);
            // Create a temporary log entry for this error without adding to queue to avoid recursion
            console.error('[LoggingService] Failed to retrieve logs from Firestore:', {
                error: error instanceof Error ? error : new Error(String(error)),
                limitCount,
                timestamp: new Date().toISOString(),
                url: window.location.href
            });
            return [];
        }
    }

    // Method to get logs by category
    async getLogsByCategory(category: string, limitCount: number = 50): Promise<LogEntry[]> {
        try {
            const logCollection = collection(db, 'application_logs');
            const q = query(
                logCollection,
                where('category', '==', category),
                orderBy('timestamp', 'desc'),
                limit(limitCount)
            );

            const querySnapshot = await getDocs(q);
            const logs: LogEntry[] = [];

            querySnapshot.forEach((doc) => {
                const data = doc.data();
                logs.push({
                    id: doc.id,
                    timestamp: data['timestamp'].toDate(),
                    level: data['level'],
                    category: data['category'],
                    action: data['action'],
                    details: data['details'],
                    userId: data['userId'],
                    sessionId: data['sessionId'],
                    component: data['component'],
                    url: data['url'],
                    userAgent: data['userAgent'],
                    ipAddress: data['ipAddress'],
                    timezone: data['timezone'],
                    language: data['language']
                });
            });

            return logs;
        } catch (error) {
            console.error('Error fetching logs by category:', error);
            // Create a temporary log entry for this error without adding to queue to avoid recursion
            console.error('[LoggingService] Failed to retrieve logs by category from Firestore:', {
                error: error instanceof Error ? error : new Error(String(error)),
                category,
                limitCount,
                timestamp: new Date().toISOString(),
                url: window.location.href
            });
            return [];
        }
    }

    // Method to get logs by date range
    async getLogsByDateRange(startDate: Date, endDate: Date, limitCount: number = 100): Promise<LogEntry[]> {
        try {
            const logCollection = collection(db, 'application_logs');
            const q = query(
                logCollection,
                where('timestamp', '>=', Timestamp.fromDate(startDate)),
                where('timestamp', '<=', Timestamp.fromDate(endDate)),
                orderBy('timestamp', 'desc'),
                limit(limitCount)
            );

            const querySnapshot = await getDocs(q);
            const logs: LogEntry[] = [];

            querySnapshot.forEach((doc) => {
                const data = doc.data();
                logs.push({
                    id: doc.id,
                    timestamp: data['timestamp'].toDate(),
                    level: data['level'],
                    category: data['category'],
                    action: data['action'],
                    details: data['details'],
                    userId: data['userId'],
                    sessionId: data['sessionId'],
                    component: data['component'],
                    url: data['url'],
                    userAgent: data['userAgent'],
                    ipAddress: data['ipAddress'],
                    timezone: data['timezone'],
                    language: data['language']
                });
            });

            return logs;
        } catch (error) {
            console.error('Error fetching logs by date range:', error);
            // Create a temporary log entry for this error without adding to queue to avoid recursion
            console.error('[LoggingService] Failed to retrieve logs by date range from Firestore:', {
                error: error instanceof Error ? error : new Error(String(error)),
                startDate: startDate.toISOString(),
                endDate: endDate.toISOString(),
                limitCount,
                timestamp: new Date().toISOString(),
                url: window.location.href
            });
            return [];
        }
    }

    // Cleanup method
    ngOnDestroy(): void {
        if (this.batchTimer) {
            clearInterval(this.batchTimer);
        }

        // Process remaining logs
        if (this.logQueue.length > 0) {
            this.processLogQueue();
        }
    }
}
