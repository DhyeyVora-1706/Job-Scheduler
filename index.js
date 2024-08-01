import express from 'express';
import {getObject} from './aws-data-fetch.js';
import { sendMessageToTeams } from './workflow.js';
import xlsx from 'xlsx';
// import './scheduler.js';
import {getDataFromSharePoint} from './sharepoint-fetch.js'


const app = express();

app.get('/', async (req, res) => {
    await executeJob();
    return res.status(200).send("OK");
});

export async function executeJob()
{
    const data = await fetchExcelDataFromSharePoint();
    const filteredData = filterData(data);
    await sendMessageToTeams(filteredData);
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

async function fetchExcelDataFromSharePoint(){
    try{      
       const data = await getDataFromSharePoint();
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


