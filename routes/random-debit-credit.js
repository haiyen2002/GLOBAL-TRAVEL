const express = require('express');
const router = express.Router();
const fileUpload = require('express-fileupload');
const multer = require('multer');
const path = require('path');
const ExcelJS = require("exceljs");

// Xử lý yêu cầu GET đến /random-sk-duyet
router.get('/', (req, res) => {
    // Thực hiện các xử lý khi nhận yêu cầu GET tới /random-sk-duyet
    res.render('random-edit-credit', { title: 'Random Edit Credit' });
});

// Cấu hình Multer để lưu trữ tệp tin được tải lên vào thư mục 'uploads'
const storage = multer.diskStorage({
    destination: (req, file, cb) => {
        cb(null, 'uploads/');
    },
    filename: (req, file, cb) => {
        cb(null, file.originalname);
    }
});

const upload = multer({ storage: storage });

// *********XỬ LÝ FILE
router.post('/upload', upload.single('uploadedFile'), async (req, res) => {
    if (!req.file) {
        return res.status(400).send('Không có tệp tin nào được tải lên.');
    }

    // function random
    function getRandomValue() {
        let randomValue = Math.floor(Math.random() * (1500000 - 100000 + 1)) + 100000; // Tạo giá trị ngẫu nhiên từ 100,000 đến 1,500,000
        randomValue = Math.round(randomValue / 1000) * 1000; // Làm tròn đến phần nghìn
        return randomValue;
      }

    // Xử lý tệp tin nếu cần
    const uploadedFileName = `./uploads/${req.file.filename}`;
    let start = req.body.start;
    let end = req.body.end;
    let debitName = req.body.debit;
    let amountDebit = req.body.amountDebit;
    let creditName = req.body.credit;

    const xlsxFilePath = uploadedFileName; // Tên file input
    let workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(xlsxFilePath);
    let sheet = workbook.getWorksheet("Sheet1");
    
    const rowsToFillC = [];
    for (let row = start; row <= end; row++) {
      rowsToFillC.push(row);
    }
    const randomRowsC = getRandom(rowsToFillC, amountDebit);
    for (let i = 0; i < amountDebit; i++) {
      let randomValueC = getRandomValue();
      sheet.getCell(`${creditName}${randomRowsC[i]}`).value = randomValueC;
    }

    const rowsToFillD = [];
    for (let row = start; row <= end; row++) {
      if (!sheet.getCell(`${creditName}${row}`).value) {
        rowsToFillD.push(row);
      }
    }
    for (let i = 0; i < rowsToFillD.length; i++) {
      let randomValueD = getRandomValue();
      sheet.getCell(`${debitName}${rowsToFillD[i]}`).value = randomValueD;
    }

    // Lưu file output
    const outputFilePath = `outputs/${req.file.filename}_output.xlsx`;
    await workbook.xlsx.writeFile(outputFilePath);

    res.json({
      message: "success",
      outputFileName: outputFilePath, 
})
});

function getRandom(arr, n) {
    const result = new Array(n);
    let len = arr.length;
    const taken = new Array(len);
    if (n > len) {
      throw new RangeError("getRandom: more elements taken than available");
    }
    while (n--) {
      const x = Math.floor(Math.random() * len);
      result[n] = arr[x in taken ? taken[x] : x];
      taken[x] = --len in taken ? taken[len] : len;
    }
    return result;
  }

module.exports = router;
