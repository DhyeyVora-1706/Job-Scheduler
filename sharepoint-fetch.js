import dotenv from 'dotenv';
dotenv.config();
import axios from 'axios';
import qs from 'qs';
import XLSX from 'xlsx';
import { generateAccessToken } from './workflow.js';



// Function to get Drive ID
const getDriveId = async (accessToken) => {
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
        throw new Error('Error fetching drive ID', error);
    }
};

// Function to get Item ID
const getItemId = async (fileName,accessToken) => {
    const driveId = await getDriveId(accessToken);
    const url = `https://graph.microsoft.com/v1.0/me/drive/root/search(q='${fileName}')`;

    try {
        const response = await axios.get(url, {
            headers: {
                Authorization: `Bearer ${accessToken}`
            }
        });
        
        // const item = response.data.value[0];
        

        const files = response.data.value;
        const exactMatchFile = files.find(file => file.name === fileName);
        const itemId = exactMatchFile.id;


        console.log(`Item ID: ${itemId}`);
        return itemId;
    } catch (error) {
        throw new Error('Error fetching item ID',error);
    }
};

// Function to get file content
const getFileContent = async (itemId,accessToken) => {
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
        throw error;
    }
};


// Main function to execute the workflow
export const getDataFromSharePoint = async (token) => {
    const fileName = 'Employees.xlsx'; // Replace with your file name
    try {
        const itemId = await getItemId(fileName,token);
        const content = await getFileContent(itemId,token);
        return content;
    } catch (error) {
        
        throw new Error(error);
    }
};

