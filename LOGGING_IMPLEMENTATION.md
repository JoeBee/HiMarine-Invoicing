# Logging Implementation for Hi Marine Invoicing

This document describes the comprehensive logging system implemented in the Hi Marine Invoicing application.

## Overview

The application now includes a complete logging system that captures all user interactions, file uploads, data processing events, exports, navigation, and errors. All logs are stored in Firebase Firestore and displayed in the History tab.

## Features

### 1. Comprehensive Logging Coverage
- **User Actions**: Button clicks, form submissions, data selections
- **File Operations**: File uploads, downloads, processing
- **Data Processing**: File analysis, data transformation, calculations
- **Exports**: Excel and PDF generation
- **Errors**: Application errors with context and stack traces
- **System Events**: Component initialization, data updates

### 2. Log Categories
- `user_action`: User interactions (clicks, form submissions)
- `file_upload`: File upload and processing events
- `data_processing`: Data transformation and analysis
- `export`: File generation and downloads
- `error`: Application errors and exceptions (with comprehensive error details)
- `system`: System-level events and initialization

### 3. Log Levels
- `info`: General information events
- `warn`: Warning messages
- `error`: Error conditions
- `debug`: Debug information

## Implementation Details

### 1. Logging Service (`src/app/services/logging.service.ts`)

The `LoggingService` provides methods for different types of logging:

```typescript
// User actions
logUserAction(action: string, details: any, component: string)

// File uploads
logFileUpload(fileName: string, fileSize: number, fileType: string, category: string, component: string)

// Data processing
logDataProcessing(action: string, details: any, component: string)

// Exports
logExport(action: string, details: any, component: string)


// Errors (Enhanced with comprehensive information)
logError(error: Error | string, context: string, component: string, additionalDetails?: any)

// Button clicks
logButtonClick(buttonName: string, component: string, additionalDetails?: any)

// Form submissions
logFormSubmission(formName: string, formData: any, component: string)

// Filter changes
logFilterChange(filterType: string, filterValue: any, component: string)

// Sort changes
logSortChange(column: string, direction: string, component: string)

// Data selection
logDataSelection(selectionType: string, selectedCount: number, totalCount: number, component: string)
```

### 2. Log Entry Structure

Each log entry contains:

```typescript
interface LogEntry {
  id?: string;                    // Firestore document ID
  timestamp: Date;                // When the event occurred
  level: 'info' | 'warn' | 'error' | 'debug';
  category: string;               // Event category
  action: string;                 // Specific action taken
  details: any;                   // Additional context data
  userId?: string;                // User identifier (if available)
  sessionId: string;              // Session identifier
  component: string;              // Angular component name
  url: string;                    // Current URL
  userAgent: string;              // Browser information
  ipAddress?: string;             // IP address (if available)
}
```

### 3. Firebase Configuration

#### Environment Files
- `src/environments/environment.ts` - Development configuration
- `src/environments/environment.prod.ts` - Production configuration

#### Firebase Setup
- Project ID: `himarine-invoicing`
- Project Number: `31012344989`
- Firestore database for log storage

### 4. History Component

The History tab (`src/app/components/history/history.component.ts`) provides:

#### Features
- **Real-time Log Display**: Shows logs as they're generated
- **Advanced Filtering**: Filter by category, level, component, date range
- **Search**: Text search across log content
- **Pagination**: Handle large numbers of logs efficiently
- **Auto-refresh**: Automatically updates every 30 seconds
- **Export**: Download logs as CSV files

#### Filter Options
- Category: user_action, file_upload, data_processing, export, error, system
- Level: info, warn, error, debug
- Component: All Angular components
- Date Range: Start and end date selection
- Search: Text search across all log fields

### 5. Data Retention

#### Automatic Cleanup
- Cloud Function runs daily at 2 AM
- Automatically deletes logs older than 1 week (7 days)
- Maintains database performance and storage costs

#### Cloud Function (`functions/index.js`)
```javascript
exports.cleanupOldLogs = functions.pubsub.schedule('0 2 * * *').onRun(async (context) => {
  // Deletes logs older than 7 days
});
```

### 6. Security Rules

Firestore security rules (`firestore.rules`):
```javascript
rules_version = '2';
service cloud.firestore {
  match /databases/{database}/documents {
    match /application_logs/{document} {
      allow read, write: if true; // Permissive for development
    }
  }
}
```

**Note**: In production, implement proper authentication and restrict access.

## Usage

### 1. Viewing Logs
1. Navigate to the History tab
2. Use filters to narrow down logs
3. Search for specific events
4. Export logs as CSV if needed

### 2. Log Analysis
- **User Behavior**: Track how users interact with the application
- **Error Monitoring**: Identify and fix application errors
- **Performance**: Monitor data processing and export operations
- **Usage Patterns**: Understand feature utilization

### 3. Troubleshooting
- Filter by error level to find issues
- Search for specific error messages
- Check component-specific logs
- Analyze user actions leading to errors

## Deployment

### 1. Firebase Setup
```bash
# Install Firebase CLI
npm install -g firebase-tools

# Login to Firebase
firebase login

# Initialize Firebase project
firebase init

# Deploy Firestore rules
firebase deploy --only firestore:rules

# Deploy Cloud Functions
firebase deploy --only functions
```

### 2. Environment Configuration
Update the Firebase configuration in environment files with your actual API keys and project details.

### 3. Production Considerations
- Implement proper authentication
- Restrict Firestore security rules
- Set up monitoring and alerting
- Configure log retention policies
- Implement user privacy controls

## Monitoring and Maintenance

### 1. Log Volume
- Monitor log volume to ensure reasonable storage usage
- Adjust retention period if needed
- Consider log aggregation for high-volume applications

### 2. Performance
- Monitor Firestore read/write operations
- Optimize queries for large datasets
- Consider pagination for better performance

### 3. Security
- Regularly review security rules
- Monitor for suspicious activity
- Implement proper access controls

## Enhanced Error Logging

The logging system now provides comprehensive error information including:

### Error Information Captured
- **Error Message**: The actual error message
- **Error Stack**: Full stack trace for debugging
- **Error Name**: Type of error (e.g., TypeError, ReferenceError)
- **Context**: Specific context where the error occurred
- **Environment Details**:
  - Timestamp (ISO format)
  - User Agent (browser information)
  - Current URL
  - Screen resolution
  - Window size
  - Language settings
  - Platform information
  - Cookie and online status
- **Additional Details**: Custom context provided by the component

### Error Logging Methods
- Accepts both `Error` objects and string messages
- Automatically logs to console for immediate debugging
- Stores comprehensive information in Firestore
- Includes component-specific context

### Error Logging Best Practices
1. Always provide meaningful context strings
2. Include relevant additional details (file names, operation types, etc.)
3. Log errors at the point where they can be handled meaningfully
4. Use consistent error categories across components

## Future Enhancements

### 1. Advanced Analytics
- User journey tracking
- Feature usage analytics
- Performance metrics
- Error trend analysis

### 2. Real-time Monitoring
- Live log streaming
- Real-time error alerts
- Performance dashboards

### 3. Integration
- External logging services
- Monitoring tools integration
- Alert systems

## Troubleshooting

### Common Issues

1. **Logs not appearing**
   - Check Firebase configuration
   - Verify Firestore rules
   - Check browser console for errors

2. **Performance issues**
   - Reduce log volume
   - Implement pagination
   - Optimize queries

3. **Storage concerns**
   - Adjust retention period
   - Implement log compression
   - Consider archiving old logs

### Support

For issues or questions about the logging implementation, check:
1. Browser console for errors
2. Firebase console for Firestore issues
3. Cloud Functions logs for retention issues
4. Network tab for API call problems

## Conclusion

The logging system provides comprehensive visibility into application usage, errors, and user behavior. It's designed to be scalable, maintainable, and useful for both development and production environments. Regular monitoring and maintenance will ensure optimal performance and security.
