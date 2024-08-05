import express from 'express';
import { sendMessageToTeams } from './workflow.js';
import xlsx from 'xlsx';
// import './scheduler.js';
import {getDataFromSharePoint} from './sharepoint-fetch.js'


const app = express();
let token;

app.get('/', async (req, res) => {
    const authorizationHeader = req.headers.authorization;

    if (!authorizationHeader) {
        return res.status(401).send('Authorization header is missing');
    }

    // Extract the bearer token
    token = authorizationHeader.split(' ')[1];
   
    if (!token) {
        return res.status(401).send('Bearer token is missing');
    }    
    // process.env.GRAPH_API_ACCESS_TOKEN = token;
    await executeJob(token);
    return res.status(200).send("OK");
});

export async function executeJob(token)
{
    const data = await fetchExcelDataFromSharePoint(token);
    const filteredData = filterData(data);
    await sendMessageToTeams(filteredData,token);
}

// async function fetchExcelDataFromS3(){
//     try{      
//        const data = await getObject();
//        return data;
//     }catch(err)
//     {
//         console.log(err);
//     }
// }

async function fetchExcelDataFromSharePoint(token){
    try{      
       const data = await getDataFromSharePoint(token);
       return data;
    }catch(err)
    {
        console.log(err);
    }
}


const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

function parseDate(dateString) {
    const [day, month, year] = dateString.split('-');
    const monthIndex = monthNames.findIndex(m => m.toLowerCase() === month.toLowerCase());
    return new Date(year, monthIndex, day);
}

function filterData(data) {
    try {
        const today = new Date();
        const sevenDaysFromNow = new Date();
        sevenDaysFromNow.setDate(today.getDate() - 7);

        const filteredData = data.filter((item) => {
            const itemDate = parseDate(item.Hire_Date);
            return itemDate <= sevenDaysFromNow;
        });

        return filteredData;
    } catch (err) {
        console.log(err);
    }
}

app.listen(4000,async () =>{
    console.log("Server is running on port 4000");
    // await executeJob();
})

// function parseDate(dateString)
// {
//     const [day,month,year] = dateString.split("/");
//     return new Date(year,month-1,day);
// }

// function filterData(data)
// {
//     try{
//         const today = new Date();
//         const sevenDaysFromNow = new Date();
//         sevenDaysFromNow.setDate(today.getDate() - 7)

//         const filteredData = data.filter((item) => {
//             const itemDate = parseDate(item.Date)
//             return itemDate <= sevenDaysFromNow
//         });
//         return filteredData;
//     }
//     catch(err){
//         console.log(err);
//     }
// }

// async function jsonToExcel(jsonData)
// {
//     const worksheet = xlsx.utils.json_to_sheet(jsonData);
//     const workbook = xlsx.utils.book_new();
//     xlsx.utils.book_append_sheet(workbook, worksheet, 'Sheet1')

//     const fileName = 'filteredData.xlsx';
//     const filepath = "./"+fileName;
//     xlsx.writeFile(workbook, filepath);
// }


