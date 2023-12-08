const express = require('express');
const mongoose = require('mongoose');
const multer = require('multer');
const xlsx = require('xlsx');

const app = express();
const port = 4800;


require('dotenv').config();

mongoose.connect(process.env.MONGODB_URI, {
useNewUrlParser: true,
useUnifiedTopology: true,
});


const excelDataSchema = new mongoose.Schema({
jsonData: Object,
});

const ExcelData = mongoose.model('ExcelData', excelDataSchema);

app.set('view engine', 'ejs');

const storage = multer.memoryStorage();
const upload = multer({ storage: storage }).single('excelFile');

app.get('/', (req, res) => {
res.render('index');
});

app.post('/upload', upload, async (req, res) => {
if (!req.file) {
return res.status(400).send('No file uploaded.');
}

const workbook = xlsx.read(req.file.buffer, { type: 'buffer' });
const sheetName = workbook.SheetNames[0];
const sheet = workbook.Sheets[sheetName];
const jsonData = xlsx.utils.sheet_to_json(sheet);

try {
const newExcelData = new ExcelData({ jsonData });
await newExcelData.save();
res.json({ message: 'Excel data saved successfully!' });
} catch (error) {
res.status(500).send('Error saving data to MongoDB.');
}
});

app.listen(port, () => {
console.log(`Server is running on port ${port}`);
});
