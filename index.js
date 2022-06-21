const xlsx = require('xlsx')
const nodemailer = require('nodemailer')
const smtpPool = require('nodemailer-smtp-pool')
const dotenv = require('dotenv')
const path = require('path');
const fs = require('fs');
const readline = require('node:readline');
const { stdin: input, stdout: output } = require('node:process');

dotenv.config()

function convertDateExcel (excelDate) {
    return (excelDate - 25569) * 86400 * 1000;
}

const generateMessage = (row) => {
    return {
        to: row.mail_id,
        from: process.env.AUTH_EMAIL,
        subject: row.subject,
        text: 'Dear Sir/Mam',
        // html: '<p style="font-family:verdana;">Dear Sir/Mam</p><p style="font-family:verdana;">'+row.body+'</p>',
        html: `<!DOCTYPE html>

        <html lang="en" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:v="urn:schemas-microsoft-com:vml">
        <head>
        <title></title>
        <meta content="text/html; charset=utf-8" http-equiv="Content-Type"/>
        <meta content="width=device-width, initial-scale=1.0" name="viewport"/>
        <style>
                * {
                    box-sizing: border-box;
                }
        
                body {
                    margin: 0;
                    padding: 0;
                }
        
                a[x-apple-data-detectors] {
                    color: inherit !important;
                    text-decoration: inherit !important;
                }
        
                #MessageViewBody a {
                    color: inherit;
                    text-decoration: none;
                }
        
                p {
                    line-height: inherit
                }
        
                .desktop_hide,
                .desktop_hide table {
                    mso-hide: all;
                    display: none;
                    max-height: 0px;
                    overflow: hidden;
                }
        
                @media (max-width:520px) {
                    .desktop_hide table.icons-inner {
                        display: inline-block !important;
                    }
        
                    .icons-inner {
                        text-align: center;
                    }
        
                    .icons-inner td {
                        margin: 0 auto;
                    }
        
                    .row-content {
                        width: 100% !important;
                    }
        
                    .mobile_hide {
                        display: none;
                    }
        
                    .stack .column {
                        width: 100%;
                        display: block;
                    }
        
                    .mobile_hide {
                        min-height: 0;
                        max-height: 0;
                        max-width: 0;
                        overflow: hidden;
                        font-size: 0px;
                    }
        
                    .desktop_hide,
                    .desktop_hide table {
                        display: table !important;
                        max-height: none !important;
                    }
                }
            </style>
        </head>
        <body style="background-color: #FFFFFF; margin: 0; padding: 0; -webkit-text-size-adjust: none; text-size-adjust: none;">
        <table border="0" cellpadding="0" cellspacing="0" class="nl-container" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; background-color: #FFFFFF;" width="100%">
        <tbody>
        <tr>
        <td>
        <table align="center" border="0" cellpadding="0" cellspacing="0" class="row row-1" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt;" width="100%">
        <tbody>
        <tr>
        <td>
        <table align="center" border="0" cellpadding="0" cellspacing="0" class="row-content stack" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; color: #000000; width: 500px;" width="500">
        <tbody>
        <tr>
        <td class="column column-1" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; font-weight: 400; text-align: left; vertical-align: top; padding-left: 5px; padding-right: 5px; padding-top: 5px; padding-bottom: 5px; border-top: 0px; border-right: 0px; border-bottom: 0px; border-left: 0px;" width="100%">
        <table border="0" cellpadding="10" cellspacing="0" class="paragraph_block" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; word-break: break-word;" width="100%">
        <tr>
        <td>
        <div style="color:#000000;direction:ltr;font-family:Verdana, Geneva, sans-serif;font-size:14px;font-weight:700;letter-spacing:0px;line-height:120%;text-align:left;mso-line-height-alt:16.8px;">
        <p style="margin: 0;">Dear Sir/Mam</p>
        </div>
        </td>
        </tr>
        </table>
        <table border="0" cellpadding="10" cellspacing="0" class="paragraph_block" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; word-break: break-word;" width="100%">
        <tr>
        <td>
        <div style="color:#000000;direction:ltr;font-family:Verdana, Geneva, sans-serif;font-size:14px;font-weight:400;letter-spacing:0px;line-height:120%;text-align:justify;mso-line-height-alt:16.8px;">
        <p style="margin: 0;">${row.body}</p>
        </div>
        </td>
        </tr>
        </table>
        <table border="0" cellpadding="10" cellspacing="0" class="paragraph_block" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; word-break: break-word;" width="100%">
        <tr>
        <td>
        <div style="color:#000000;direction:ltr;font-family:Verdana, Geneva, sans-serif;font-size:14px;font-weight:400;letter-spacing:0px;line-height:120%;text-align:left;mso-line-height-alt:16.8px;">
        <p style="margin: 0; margin-bottom: 0px;">With Regards,</p>
        <p style="margin: 0; margin-bottom: 0px;">Lorem ipsum</p>
        <p style="margin: 0; margin-bottom: 0px;">Lorem</p>
        <p style="margin: 0;">sct</p>
        </div>
        </td>
        </tr>
        </table>
        </td>
        </tr>
        </tbody>
        </table>
        </td>
        </tr>
        </tbody>
        </table>
        </td>
        </tr>
        </table>
        </td>
        </tr>
        </tbody>
        </table>
        </td>
        </tr>
        </tbody>
        </table>
        </td>
        </tr>
        </tbody>
        </table>
        </body>
        </html>`
    };
}

const readFileAndSendMail = (dirPath, fileName) => {

    const file = xlsx.readFile(`${dirPath}/${fileName}`);
    let fileData = []
    const sheets = file.SheetNames
    
    for(let i = 0; i < sheets.length; i++){
        const temp = xlsx.utils.sheet_to_json(file.Sheets[file.SheetNames[i]])
        temp.forEach((res) => {
            fileData.push(res)
        })
    }
    
    const transporter = nodemailer.createTransport(smtpPool({
        name: 'smtp.office365.com',
        host: 'smtp.office365.com',
        port: 587,
        auth: {
            user: process.env.AUTH_EMAIL,
            pass: process.env.AUTH_PASS
        },
        maxConnections: 5,
        maxMessages: 10,
        rateLimit: 5
    }));
    transporter.verify((error, success) => {
        if(error) console.log(error)
    })
    var messages = []
    fileData.forEach((row) => {
        if(convertDateExcel(row.deadline) < Date.now()){
            if(row.status === 'not completed'){
                console.log(`${row.name} has passed the deadline`)
                messages.push(generateMessage(row))
            }
        }
    })
    transporter.on('idle', async function(){
        while(transporter.isIdle() && messages.length){
            const info = await transporter.sendMail(messages.shift());
            if(!info){
                console.log("Unable to send mail")
            }
            console.log(`Mail sent to ${info.to[0]}`);
        }
    });    
}

const directoryPath = path.join(__dirname, 'spread_sheets');

fs.readdir(directoryPath, async (err, files) => {
    if (err) {
        return console.log('Unable to scan directory: ' + err);
    }
    console.log("Available files:");
    let counter = 0;
    files.forEach(function (file) {
        counter++;
        console.log(`${counter}. `,file); 
    });
    const rl = readline.createInterface({ input, output });
    rl.question('Enter the file number to check deadline: ', (fileNum) => {
        console.log("You have selected the file --> ",files[fileNum-1]);
        rl.question('Do you need to check deadline for this file?(y/n)\n', (flag) => {
            if(flag.toLowerCase() === "y" || flag.toLowerCase() === "yes"){
                readFileAndSendMail(directoryPath, files[fileNum-1]);
            }
            else{
                console.log("You have selected the wrong file, please try again later...");
            }
            rl.close();
        });
    });
});
