const express = require("express");
const bodyParser = require("body-parser");
const multer = require('multer');
const exceljs = require('exceljs');
const mongoose = require("mongoose");

const app = express();
app.use(bodyParser.json());
app.use(express.static('public'));
app.use(bodyParser.urlencoded({
    extended: true
}));

mongoose.connect('mongodb://localhost:27017/MoneyList');
const db = mongoose.connection;
db.on('error', () => console.log("Error in connecting to the Database"));
db.once('open', () => console.log("Connected to Database"));

app.post("/add", (req, res) => {
    const category_select = req.body.category_select;
    const amount_input = req.body.amount_input;
    const info = req.body.info;
    const date_input = req.body.date_input;

    const data = {
        "Category": category_select,
        "Amount": amount_input,
        "Info": info,
        "Date": date_input
    };
    db.collection('users').insertOne(data, (err, collection) => {
        if (err) {
            throw err;
        }
        console.log("Record Inserted Successfully");
    });
});

const storage = multer.memoryStorage();
const upload = multer({ storage: storage });

app.post('/upload', upload.single('file'), async (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({ message: 'No file uploaded' });
        }

        const workbook = new exceljs.Workbook();
        const buffer = req.file.buffer;
        await workbook.xlsx.load(buffer);

        const worksheet = workbook.worksheets[0];
        const data = [];

        worksheet.eachRow((row, rowNumber) => {
            if (rowNumber !== 1) { // Skip header row
                data.push({
                    category: row.getCell(1).value,
                    amount: row.getCell(2).value,
                    info: row.getCell(3).value,
                    date: row.getCell(4).value
                });
            }
        });

        // Insert data into MongoDB
        // Example:
        // await db.collection('users').insertMany(data);

        res.status(200).json(data);
    } catch (error) {
        console.error('Error uploading file:', error);
        res.status(500).json({ message: 'Internal server error' });
    }
});

app.get("/", (req, res) => {
    res.set({
        "Allow-access-Allow-Origin": '*'
    });
    return res.redirect('index.html');
}).listen(5000);

console.log("Listening on port 5000");