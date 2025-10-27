import { Injectable } from '@angular/core';
import { collection, addDoc, query, orderBy, limit, getDocs, where, Timestamp } from 'firebase/firestore';
import { db } from '../firebase.config';

export interface LogEntry {
    id?: string;
    timestamp: Date;
    level: 'info' | 'warn' | 'error' | 'debug';
    category: 'user_action' | 'file_upload' | 'data_processing' | 'export' | 'navigation' | 'error' | 'system';
    action: string;
    details: any;
    userId?: string;
    sessionId: string;
    component: string;
    url: string;
    userAgent: string;
    ipAddress?: string;
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

    constructor() {
        this.sessionId = this.generateSessionId();
        this.setupOnlineListener();
        this.startBatchProcessor();
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
        return {
            timestamp: new Date(),
            level,
            category,
            action,
            details,
            sessionId: this.sessionId,
            component,
            url: window.location.href,
            userAgent: navigator.userAgent
        };
    }

    // Public logging methods
    logUserAction(action: string, details: any, component: string): void {
        const logEntry = this.createLogEntry('info', 'user_action', action, details, component);
        this.addToQueue(logEntry);
    }

    logFileUpload(fileName: string, fileSize: number, fileType: string, category: string, component: string): void {
        const logEntry = this.createLogEntry('info', 'file_upload', 'file_uploaded', {
            fileName,
            fileSize,
            fileType,
            category
        }, component);
        this.addToQueue(logEntry);
    }

    logDataProcessing(action: string, details: any, component: string): void {
        const logEntry = this.createLogEntry('info', 'data_processing', action, details, component);
        this.addToQueue(logEntry);
    }

    logExport(action: string, details: any, component: string): void {
        const logEntry = this.createLogEntry('info', 'export', action, details, component);
        this.addToQueue(logEntry);
    }

    logNavigation(fromUrl: string, toUrl: string, component: string): void {
        const logEntry = this.createLogEntry('info', 'navigation', 'navigation', {
            fromUrl,
            toUrl
        }, component);
        this.addToQueue(logEntry);
    }

    logError(error: Error, context: string, component: string, additionalDetails?: any): void {
        const logEntry = this.createLogEntry('error', 'error', 'error_occurred', {
            errorMessage: error.message,
            errorStack: error.stack,
            context,
            additionalDetails
        }, component);
        this.addToQueue(logEntry);
    }

    logSystemEvent(action: string, details: any, component: string): void {
        const logEntry = this.createLogEntry('info', 'system', action, details, component);
        this.addToQueue(logEntry);
    }

    logButtonClick(buttonName: string, component: string, additionalDetails?: any): void {
        const logEntry = this.createLogEntry('info', 'user_action', 'button_click', {
            buttonName,
            ...additionalDetails
        }, component);
        this.addToQueue(logEntry);
    }

    logFormSubmission(formName: string, formData: any, component: string): void {
        const logEntry = this.createLogEntry('info', 'user_action', 'form_submission', {
            formName,
            formData: this.sanitizeFormData(formData)
        }, component);
        this.addToQueue(logEntry);
    }

    logFilterChange(filterType: string, filterValue: any, component: string): void {
        const logEntry = this.createLogEntry('info', 'user_action', 'filter_change', {
            filterType,
            filterValue
        }, component);
        this.addToQueue(logEntry);
    }

    logSortChange(column: string, direction: string, component: string): void {
        const logEntry = this.createLogEntry('info', 'user_action', 'sort_change', {
            column,
            direction
        }, component);
        this.addToQueue(logEntry);
    }

    logDataSelection(selectionType: string, selectedCount: number, totalCount: number, component: string): void {
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
                    ipAddress: data['ipAddress']
                });
            });

            return logs;
        } catch (error) {
            console.error('Error fetching logs:', error);
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
                    ipAddress: data['ipAddress']
                });
            });

            return logs;
        } catch (error) {
            console.error('Error fetching logs by category:', error);
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
                    ipAddress: data['ipAddress']
                });
            });

            return logs;
        } catch (error) {
            console.error('Error fetching logs by date range:', error);
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
