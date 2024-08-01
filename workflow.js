import 'node-fetch';
import { ClientSecretCredential } from '@azure/identity';
import 'node-fetch';
import querystring from 'querystring';

const tenantId = '6ba04439-8b0e-43ee-ad26-c2ac9ef9e765';
const clientId = 'ba1c4221-5aec-4b7b-8d82-ba7a8611377f';
const clientSecret = '6LM8Q~I.eabZMrXRY7v5y4kpl6gX0qaCbtGlCbZL';
const teamId = '1434253a-d09e-432d-b3b3-b902f07927f3';
const channelId = '19:rzs-cne6AtTfc0JmbaX5f2TZJR6lbOqu2Sv-k_dur6s1@thread.tacv2';

function generateTableHTML(data)
{
    let tableHtml = `<h4>Data Captured as of ${new Date().toLocaleString()}</h4> <br><br><hr/>`;
    tableHtml += '<table border="1" style="border-collapse: collapse;">';
    tableHtml+=` <thead>
            <tr>
                <th style="width: 10%;">Employee ID</th>
                <th style="width: 13%;">Employee Name</th>
                <th style="width: 10%;">Hire Date</th>
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
                    ${element['Hire_Date']}
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



export async function sendMessageToTeams(dataSent) {
    try{
    let HTMLoutput = generateTableHTML(dataSent);
    const accessToken = process.env.GRAPH_API_ACCESS_TOKEN;
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
        console.log(err);
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