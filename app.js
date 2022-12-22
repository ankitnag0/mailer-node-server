const express = require('express');
const fileUpload = require('express-fileupload');
const nodemailer = require('nodemailer');
const xlsx = require('xlsx');
const cors = require('cors');
require('dotenv').config();

const {createServer} = require('http');
const {Server} = require('socket.io');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors());

app.use(fileUpload());

app.post('/upload', async function (req, res) {

    try {

        /* File Upload*/
        let emailList;
        let emailListUplaodPath;

        let attachment;
        let attachmentUploadPath;

        let {subject, mailContent} = req.body;
        if (!req.files || Object.keys(req.files).length === 0) {
            console.log('No files were uploaded')
            return res.status(400).send('No files were uploaded.');
        }
    
        emailList = req.files.emailList;
        emailListUplaodPath = __dirname + '/upload/' + emailList.name;

        attachment = req.files.attachment;
        attachmentUploadPath = __dirname + '/upload/' + attachment.name;
    
        await emailList.mv(emailListUplaodPath);
        await attachment.mv(attachmentUploadPath);
        /* File Upload Ends*/

        /* Reading Excel File */
        const workbook = xlsx.readFile(`./upload/${emailList.name}`);
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
            subject: subject,
            text: mailContent,
            attachments: [
                {
                    filename: attachment.name,
                    path: './upload/'+attachment.name
                }
            ]
        }

        for(let i=0; i<emails.length; i++){
            options.to = emails[i];
            await client.sendMail(options);
            console.log("Success: " + options.to);
            io.emit('mailSuccess', {email: options.to});
        }
        /* Sending Mails Ends*/

        return res.status(200).json({
            message: "Email sent successfully",
        });
    } catch (error) {
        console.log(error);
        return res.status(500).json({
            error: error.message,
            message: "Internal Server Error"
        });
    }
});

const server = createServer(app);

const io = new Server(server, {
    cors: {
        origin: "http://127.0.0.1:5173"
    }
});
//localhost is not equal to 127.0.0.1

io.on('connection', socket => {
    console.log(socket.id);
})

server.listen(PORT, () => {
    console.log(`Server is listening on port ${PORT}`);
});

module.exports = app;