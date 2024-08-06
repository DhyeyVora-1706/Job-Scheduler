import dotenv from 'dotenv';
dotenv.config();
import 'node-fetch';
import { ClientSecretCredential } from '@azure/identity';
import 'node-fetch';
import querystring from 'querystring';

const tenantId = process.env.TENANTID;
const clientId = process.env.CLIENTID;
const clientSecret = process.env.CLIENTSECRET;
const teamId = process.env.TEAMID;
const channelId = process.env.CHANNELID;


const getISTDateTime = () => {
    // Create a date object
    const date = new Date();
    
    // Define options for formatting
    const options = {
        timeZone: 'Asia/Kolkata', // IST time zone
        year: 'numeric',
        month: '2-digit',
        day: '2-digit',
        hour: '2-digit',
        minute: '2-digit',
        second: '2-digit',
        hour12: false // Use 24-hour format
    };

    // Format the date object according to IST
    const formatter = new Intl.DateTimeFormat('en-IN', options);
    return formatter.format(date);
};

function generateTableHTML(data)
{
    let ISTDateTime = getISTDateTime();
    let tableHtml = `<h4>Data Captured as of ${ISTDateTime}</h4> <br><br><hr/>`;
    tableHtml += '<table border="1" style="border-collapse: collapse;">';
    tableHtml+=` <thead>
            <tr>
                <th style="width: 10%;">Employee ID</th>
                <th style="width: 13%;">Employee Name</th>
                <th style="width: 10%;">Availability Date</th>
                <th style="width: 12%;">Work Location</th>
                <th style="width: 10%;">Designation</th>
                <th style="width: 15%;">Advanced Level Skills</th>
                <th style="width: 10%;">Reporting Manager</th>
                <th style="width: 10%;">Billing Status</th>
                <th style="width: 15%;">Mail Id</th>
            </tr>
        </thead> '<tbody>`
    data.forEach(element => {
        tableHtml += `
            <tr>
                <td>
                    ${element['Employee ID']}
                </td>
                <td>
                    ${element['Employee Name']}
                </td>
                <td>
                    ${element['Available Start Date']}
                </td>
                <td>
                    ${element['Work Location']}
                </td>
                 <td>
                    ${element['Designation']}
                </td>
                <td>
                    ${element['Advanced level proficiency Skills']}
                </td>
                <td>
                    ${element['Reporting Manager']}
                </td>
                <td>
                    ${element['Billing Status']}
                </td>
                <td>
                    ${element['Mail ID']}
                </td>               
            </tr>`;
    });
    tableHtml += "</tbody></table>";

    return tableHtml;
}

export async function generateAccessToken() {
    try {
        const response = await fetch(`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/x-www-form-urlencoded'
            },
            body: querystring.stringify({
                grant_type: 'client_credentials',
                client_id: clientId,
                client_secret: clientSecret,
                scope: 'https://graph.microsoft.com/.default'
            })
        });

        if (!response.ok) {
            throw new Error(`Failed to obtain access token: ${response.statusText} (${response.status})`);
        }

        const data = await response.json();
        console.log('Access token generated successfully:', data.access_token); // Logging the token for debugging
        return data.access_token;
    } catch (err) {
        console.error('Error generating access token:', err);
    }
}



export async function sendMessageToTeams(dataSent,accessToken) {
    try{
    let HTMLoutput = generateTableHTML(dataSent);
    const response = await fetch(`https://graph.microsoft.com/v1.0/teams/${teamId}/channels/${channelId}/messages`, {
        method: 'POST',
        headers: {
            'Authorization': `Bearer ${accessToken}`,
            'Content-Type': 'application/json'
        },
        body: JSON.stringify({
            body: {
                contentType: 'html',
                content: HTMLoutput
            }
        })
    });

    if (response.status === 401) {
        throw new Error('API unauthorised');
    }

    const data = await response.json();
    }
    catch(err)
    {
        throw new Error(err);
    }
}


// async function generateAccessToken()
// {
//     try{
//        const response = await fetch(`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`, {
//             method: 'POST',
//             headers: {
//                 'Content-Type': 'application/x-www-form-urlencoded'
//             },
//             body: new URLSearchParams({
//                 grant_type: 'client_credentials',
//                 client_id: clientId,
//                 client_secret: clientSecret,
//                 scope: 'https://graph.microsoft.com/.default'
//             })
//         });

//         if (!response.ok) {
//             throw new Error(`Failed to obtain access token: ${response.statusText}`);
//         }

//         const data = await response.json();
//         return data.access_token;
//     }
//     catch(err)
//     {
//         console.log(err);
//     }
// }