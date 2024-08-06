import dotenv from 'dotenv';
dotenv.config();
import express from 'express';
import { sendMessageToTeams } from './workflow.js';
import {getDataFromSharePoint} from './sharepoint-fetch.js'


const app = express();
let token;
let direction="before";
let days=7;

app.get('/', async (req, res) => {
    try{

        const authorizationHeader = req.headers.authorization;

        if (!authorizationHeader) {
            return res.status(401).send('Authorization header is missing');
        }

        // Extract the bearer token
        token = authorizationHeader.split(' ')[1];
    
        if (!token) {
            return res.status(401).send('Bearer token is missing');
        }
        
        if(req.query.direction !== '' && req.query.direction !== undefined)
        {
            direction=req.query.direction;
        }

        if(req.query.days !== '' && req.query.days !== undefined)
        {
            if(parseInt(req.query.days) <= 0)
            {
                return res.status(400).send("Error in days defined in URL , it should be before of after only");
            }
            days=parseInt(req.query.days);
        }
        
        await executeJob(token);
        return res.status(200).send("Data Sent teams channel successfully");
    }catch(err)
    {
        return res.status(500).send("Internal Server Error");
    }
});

export async function executeJob(token)
{
    const data = await fetchExcelDataFromSharePoint(token);
    const filteredData = filterData(data,days,direction);
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
        throw new Error(err);
    }
}


const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];


function createISTDate(year, monthIndex, day) {
    // Create a date object in UTC
    const utcDate = new Date(Date.UTC(year, monthIndex, day));
    // Convert UTC date to IST by adding 5 hours and 30 minutes
    return new Date(utcDate.getTime() + (5.5 * 60 * 60 * 1000));
}

// Function to parse a date string and return a date object in IST
function parseDate(dateString) {
    const [day, month, year] = dateString.split('-');
    const monthIndex = monthNames.findIndex(m => m.toLowerCase() === month.toLowerCase());
    // Create and return the IST date
    return createISTDate(year, monthIndex, day);
}

// Function to filter data based on date range
function filterData(data, days, direction) {
    try {
        // Get today's date in IST
        const today = new Date();
        const todayIST = createISTDate(today.getUTCFullYear(), today.getUTCMonth(), today.getUTCDate());
        

        // Create reference date based on direction
        const referenceDate = new Date(todayIST);
        if (direction === 'before') {
            referenceDate.setDate(todayIST.getDate() - days);
        } else if (direction === 'after') {
            referenceDate.setDate(todayIST.getDate() + days);
        } else {
            throw new Error('Invalid direction. Use "before" or "after".');
        }
        
        referenceDate.setHours(0, 0, 0, 0);

        
        const filteredData = data.filter((item) => {
            if (item['Available Start Date']) {
                const itemDate = parseDate(item['Available Start Date']);
                if (direction === 'before') {
                    return itemDate <= referenceDate;
                } else if (direction === 'after') {
                    return itemDate >= referenceDate;
                }
            }
            return false;
        });

        return filteredData;
    } catch (err) {
        throw new Error(err);
    }
}



app.listen(80,async () =>{
    console.log("Server is running on port 80");
})



