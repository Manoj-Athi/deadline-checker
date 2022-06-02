const xlsx = require('xlsx')
const nodemailer = require('nodemailer')
const dotenv = require('dotenv')
const path = require('path');
const fs = require('fs');
const readline = require('node:readline');
const { stdin: input, stdout: output } = require('node:process');

dotenv.config()

function convertDateExcel (excelDate) {
    // unix time stamp
    return (excelDate - 25569) * 86400 * 1000;
}

const sendMail = async (row) => {
    const transporter = nodemailer.createTransport({
        name: 'gmail.com',
        host: 'smtp.gmail.com',
        port: 587,
        auth: {
            user: process.env.AUTH_EMAIL,
            pass: process.env.AUTH_PASS
        }
    });
    transporter.verify((error, success) => {
        if(error) console.log(error)
    })
    var mailOptions = {
        to: row.mail_id,
        from: process.env.AUTH_EMAIL,
        subject: row.subject,
        text: `Hello ${row.name}! Time's up.\r\n\r\n`,
        html: '<h3>Hello '+row.name+'!</h3><p>Your time has been ended to complete the task!!!</p>',
    };

    let info = await transporter.sendMail(mailOptions)
    if(!info){
        console.log("Unable to send mail")
    }
    console.log(`Mail sent to ${row.name} with id: ${row.id}`)
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
    
    fileData.forEach(async (row) => {
        if(convertDateExcel(row.deadline) < Date.now()){
            if(row.status === 'not completed'){
                console.log(`${row.name} has passed the deadline`)
                await sendMail(row)
            }
        }
    })
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
