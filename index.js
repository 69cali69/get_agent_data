import exceljs from 'exceljs';
import fs from 'fs';

const delay = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

const getAgentData = async () => {
    let agents;
    const agentData = [];

    try {
        agents = JSON.parse(fs.readFileSync('./member_lists/boston.json', 'utf8'));
    } catch (error) {
        console.error('Error reading agents.json');
        return;
    }

    for (const agent of agents) {
        const index = agents.indexOf(agent);

         //if (index > 9) continue; // Uncomment to limit the number of agents fetched

        try {
            const response = await fetch(`https://directories.apps.realtor/directories/getMemberDetail`, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify({
                    "memberLastName": agent.LastName,
                    "officeStreetCountry": "US",
                    "personId": agent.PersonId,
                })
            })

            const status = response.status;

            if (status === 200) {
                const data = await response.json();

                agentData.push(data)

                console.log(`Fetched agent: ${agent.FirstName} ${agent.LastName}`);
            } else {
                console.log(`Error fetching agent: ${agent.FirstName} ${agent.LastName}`);
            }
        } catch (error) {
            console.error('Error parsing agent:', error);
            break;
        }

        await delay(5000);
    }

    return agentData;
}

async function main() {
    const workbook = new exceljs.Workbook();
    await workbook.xlsx.readFile('./agents.xlsx');

    const sheet = workbook.addWorksheet('Boston');

    await getAgentData().then((data) => {
        const columns = [
            { header: 'First Name', key: 'firstName', width: 20 },
            { header: 'Last Name', key: 'lastName', width: 20 },
            { header: 'Email', key: 'email', width: 40 },
            { header: 'Brokerage', key: 'brokerage', width: 40 },
            { header: 'City', key: 'city', width: 20 },
            { header: 'State', key: 'state', width: 20 },
        ];
        
        sheet.columns = columns;
        
        data.forEach((agent) => {
           
            sheet.addRow({
                firstName: agent.FirstName,
                lastName: agent.LastName,
                email: agent.BusinessEmailAddress,
                brokerage: agent.Office.OfficeBusinessName || 'N/A',
                city: agent.Office.MailingCity || 'N/A',
                state: agent.Office.MailingState || 'N/A',
            });
        })
    }).then(async () => {
        await workbook.xlsx.writeFile('./agents.xlsx').then(() => {
            console.log('File saved');
        });
    })
}

main();