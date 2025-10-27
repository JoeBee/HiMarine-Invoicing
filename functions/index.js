const functions = require('firebase-functions');
const admin = require('firebase-admin');

admin.initializeApp();

// Cloud Function to clean up old logs (older than 1 week)
exports.cleanupOldLogs = functions.pubsub.schedule('0 2 * * *').onRun(async (context) => {
    const db = admin.firestore();
    const oneWeekAgo = new Date();
    oneWeekAgo.setDate(oneWeekAgo.getDate() - 7); // 7 days ago

    console.log('Starting cleanup of logs older than:', oneWeekAgo);

    try {
        const logsRef = db.collection('application_logs');
        const oldLogsQuery = logsRef.where('timestamp', '<', admin.firestore.Timestamp.fromDate(oneWeekAgo));

        const snapshot = await oldLogsQuery.get();

        if (snapshot.empty) {
            console.log('No old logs found to delete');
            return null;
        }

        const batch = db.batch();
        let deleteCount = 0;

        snapshot.docs.forEach((doc) => {
            batch.delete(doc.ref);
            deleteCount++;
        });

        await batch.commit();

        console.log(`Successfully deleted ${deleteCount} old log entries`);
        return { deletedCount: deleteCount };
    } catch (error) {
        console.error('Error during log cleanup:', error);
        throw error;
    }
});

// Cloud Function to get log statistics
exports.getLogStats = functions.https.onCall(async (data, context) => {
    const db = admin.firestore();

    try {
        const logsRef = db.collection('application_logs');

        // Get total count
        const totalSnapshot = await logsRef.count().get();
        const totalLogs = totalSnapshot.data().count;

        // Get logs by category
        const categories = ['user_action', 'file_upload', 'data_processing', 'export', 'navigation', 'error', 'system'];
        const categoryStats = {};

        for (const category of categories) {
            const categorySnapshot = await logsRef.where('category', '==', category).count().get();
            categoryStats[category] = categorySnapshot.data().count;
        }

        // Get logs by level
        const levels = ['info', 'warn', 'error', 'debug'];
        const levelStats = {};

        for (const level of levels) {
            const levelSnapshot = await logsRef.where('level', '==', level).count().get();
            levelStats[level] = levelSnapshot.data().count;
        }

        // Get recent activity (last 24 hours)
        const oneDayAgo = new Date();
        oneDayAgo.setDate(oneDayAgo.getDate() - 1);

        const recentSnapshot = await logsRef
            .where('timestamp', '>=', admin.firestore.Timestamp.fromDate(oneDayAgo))
            .count()
            .get();

        const recentLogs = recentSnapshot.data().count;

        return {
            totalLogs,
            categoryStats,
            levelStats,
            recentLogs,
            lastUpdated: new Date().toISOString()
        };
    } catch (error) {
        console.error('Error getting log stats:', error);
        throw new functions.https.HttpsError('internal', 'Failed to get log statistics');
    }
});
