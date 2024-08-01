import cron from 'node-cron';
import { executeJob } from './index.js';

console.log('CRON Job is getting configured');

cron.schedule('*/2 * * * *', () => {
    console.log('Executing function at ' + new Date().toLocaleString());
    executeJob();
}, {
    scheduled: true,
    timezone: "Asia/Kolkata" // India time zone
});

console.log('Cron job scheduled');
