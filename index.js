const express = require('express');
const fileUpload = require('express-fileupload');
const nodemailer = require('nodemailer');
const xlsx = require('xlsx');
require('dotenv').config();

const app = express();
const PORT = process.env.PORT || 3000;
// default options
app.use(fileUpload());

app.post('/upload', async function (req, res) {

    try {

        /* File Upload*/
        let excelFile;
        let uploadPath;
    
        if (!req.files || Object.keys(req.files).length === 0) {
            return res.status(400).send('No files were uploaded.');
        }
    
        excelFile = req.files.excelFile;
        uploadPath = __dirname + '/upload/' + excelFile.name;
    
        await excelFile.mv(uploadPath);
        /* File Upload Ends*/

        /* Reading Excel File */
        const workbook = xlsx.readFile(`./upload/${excelFile.name}`);
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const emails = [];
        for(let row in worksheet){
            emails.push(worksheet[row].v);
        }
        emails.shift(); 
        emails.pop();
        /* Reading Excel File Ends*/

        /* Sending Mails*/
        const client = nodemailer.createTransport({
            service: "Gmail",
            auth: {
                user: process.env.EMAIL,
                pass: process.env.PASSWORD
            }
        });

        const options = {
            from: process.env.EMAIL,
            subject: "From Nodemailer",
            text: "Hi Ankit, I am Ankit!"
        }

        for(let i=0; i<emails.length; i++){
            options.to = emails[i];
            await client.sendMail(options);
            console.log("Success: " + options.to);
        }
        /* Sending Mails Ends*/

        return res.status(200).json({
            message: "Email sent successfully",
        });
    } catch (error) {
        return res.status(500).json({
            error: error.message,
            message: "Internal Server Error"
        });
    }
});

app.listen(PORT, function () {
    console.log('Express server listening on port ', PORT);
});