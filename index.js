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
            days=req.query.days;
        }
        
        await executeJob(token);
        return res.status(200).send("Data Sent teams channel successfully");
    }catch(err)
    {
        console.log(err);
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
        console.log(err);
    }
}


const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

function parseDate(dateString) {
    const [day, month, year] = dateString.split('-');
    const monthIndex = monthNames.findIndex(m => m.toLowerCase() === month.toLowerCase());
    return new Date(year, monthIndex, day);
}

function filterData(data,days,direction) {
    try {
        const today = new Date();
        const referenceDate = new Date();

        console.log(direction);
        console.log(days);
        

        if (direction === 'before') {
            referenceDate.setDate(today.getDate() - days);
        } else if (direction === 'after') {
            referenceDate.setDate(today.getDate() + days);
        } else {
            throw new Error('Invalid direction. Use "before" or "after".');
        }

        console.log(referenceDate);

        const filteredData = data.filter((item) => {
            const itemDate = parseDate(item['Hire Date']);
            if (direction === 'before') {
                return itemDate <= referenceDate;
            } else if (direction === 'after') {
                return itemDate >= referenceDate;
            }
        });

        return filteredData;
    } catch (err) {
        console.log(err);
    }
}

app.listen(4000,async () =>{
    console.log("Server is running on port 4000");
})



