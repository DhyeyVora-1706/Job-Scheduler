import dotenv from 'dotenv';
dotenv.config();

// Import necessary AWS SDK v3 modules
import { S3Client, GetObjectCommand } from '@aws-sdk/client-s3';
import { fromEnv } from '@aws-sdk/credential-provider-env';
import xlsx from 'xlsx';

// Initialize the S3 client with the environment credentials
const s3Client = new S3Client({
  region: process.env.AWS_REGION,
  credentials: fromEnv(), // Uses AWS credentials from environment variables
});

// Parameters for fetching the object from S3
const params = {
  Bucket: 'job-scheduler-bucket', // Replace with your bucket name
  Key: 'JobScheduler.xlsx', // Replace with your file name
};

// Function to convert a readable stream to a buffer
const streamToBuffer = async (stream) => {
  return new Promise((resolve, reject) => {
    const chunks = [];
    stream.on('data', (chunk) => chunks.push(chunk));
    stream.on('end', () => resolve(Buffer.concat(chunks)));
    stream.on('error', reject);
  });
};

// Function to fetch and process the Excel file from S3
export const getObject = async () => {
  try {
    const command = new GetObjectCommand(params);
    const response = await s3Client.send(command);

    // Convert the response body (stream) to a buffer
    const data = await streamToBuffer(response.Body);

    // Process the Excel file
    const workbook = xlsx.read(data, { type: 'buffer' });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const jsonData = xlsx.utils.sheet_to_json(sheet);

    // Output the JSON data
    return jsonData;
  } catch (err) {
    console.error('Error fetching or processing the Excel file:', err);
  }
};
