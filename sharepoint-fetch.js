import dotenv from 'dotenv';
dotenv.config();
import axios from 'axios';
import XLSX from 'xlsx';




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
    const folderName = 'Scheduler';
    const fileName = 'Employees.xlsx'; // Replace with your file name
    try {
        // const itemId = await getItemId(fileName,token);
        // const content = await getFileContent(itemId,token);
        // return content;
        const folderId = await getFolderIdByName(folderName, token);
        const fileId = await getFileIdByName(folderId, fileName, token);
        const content = await getFileContent(fileId, token);
        return content;
    } catch (error) {
        
        throw new Error(error);
    }
};

const getFolderIdByName = async (folderName, accessToken) => {
    const url = `https://graph.microsoft.com/v1.0/me/drive/root/children`;

    try {
        const response = await axios.get(url, {
            headers: {
                Authorization: `Bearer ${accessToken}`
            }
        });

        const folder = response.data.value.find(item => item.name === folderName && item.folder);
        if (!folder) {
            throw new Error(`Folder ${folderName} not found`);
        }
        console.log(`Folder ID: ${folder.id}`);
        return folder.id;
    } catch (error) {
        console.error('Error fetching folder ID:', error);
        throw error;
    }
};


const getFileIdByName = async (folderId, fileName, accessToken) => {
    const url = `https://graph.microsoft.com/v1.0/me/drive/items/${folderId}/children`;

    try {
        const response = await axios.get(url, {
            headers: {
                Authorization: `Bearer ${accessToken}`
            }
        });

        const file = response.data.value.find(item => item.name === fileName && item.file);
        if (!file) {
            throw new Error(`File ${fileName} not found in folder ${folderId}`);
        }
        
        console.log(`File ID: ${file.id}`);
        return file.id;
    } catch (error) {
        console.error('Error fetching file ID:', error);
        throw error;
    }
};