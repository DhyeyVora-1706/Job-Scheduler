import dotenv from 'dotenv';
dotenv.config();
import axios from 'axios';
import qs from 'qs';
import XLSX from 'xlsx';
import { generateAccessToken } from './workflow.js';

let accessToken = process.env.GRAPH_API_ACCESS_TOKEN;
// Function to get the access token
const getAccessToken = async () => {
    const tenantId = process.env.TENANTID; // Replace with your tenant ID
    const clientId = process.env.CLIENTID; // Replace with your client ID
    const clientSecret = process.env.CLIENTSECRET; // Replace with your client secret
    const scope = 'https://graph.microsoft.com/.default';
    const url = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;

    const requestBody = {
        client_id: clientId,
        client_secret: clientSecret,
        scope: scope,
        grant_type: 'client_credentials'
    };

    try {
        const response = await axios.post(url, qs.stringify(requestBody), {
            headers: {
                'Content-Type': 'application/x-www-form-urlencoded'
            }
        });
        return response.data.access_token;
    } catch (error) {
        console.error('Error fetching access token', error);
    }
};

// Function to get Drive ID
const getDriveId = async () => {
    const url = 'https://graph.microsoft.com/v1.0/me/drive';

    try {
        const response = await axios.get(url, {
            headers: {
                Authorization: `Bearer ${accessToken}`
            }
        });
        const driveId = response.data.id;
        console.log(`Drive ID: ${driveId}`);
        return driveId;
    } catch (error) {
        console.error('Error fetching drive ID', error);
    }
};

// Function to get Item ID
const getItemId = async (fileName) => {
    const driveId = await getDriveId();
    const url = `https://graph.microsoft.com/v1.0/me/drive/root/search(q='${fileName}')`;

    try {
        const response = await axios.get(url, {
            headers: {
                Authorization: `Bearer ${accessToken}`
            }
        });
        const item = response.data.value[0];
        const itemId = item.id;
        console.log(`Item ID: ${itemId}`);
        return itemId;
    } catch (error) {
        console.error('Error fetching item ID', error);
    }
};

// Function to get file content
const getFileContent = async (itemId) => {
    const url = `https://graph.microsoft.com/v1.0/me/drive/items/${itemId}/content`;

    try {
        const response = await axios.get(url, {
            headers: {
                Authorization: `Bearer ${accessToken}`
            },
            responseType: 'arraybuffer' // For binary data
        });
        
        const workbook = XLSX.read(response.data, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        const dataObjects = XLSX.utils.sheet_to_json(sheet);

        console.log('File content parsed successfully');
        return dataObjects;
    } catch (error) {
        console.error('Error fetching file content');
        throw error;
    }
};


// Main function to execute the workflow
export const getDataFromSharePoint = async () => {
    const fileName = 'Employee.xlsx'; // Replace with your file name
    try {
        const itemId = await getItemId(fileName);
        const content = await getFileContent(itemId);
        return content;
    } catch (error) {
        console.error('Error in workflow', error);
    }
};

