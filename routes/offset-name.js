const express = require('express');
const router = express.Router();
const fileUpload = require('express-fileupload');
const multer = require('multer');
const path = require('path');
const ExcelJS = require("exceljs");
const xlsx = require("xlsx");

router.get('/', (req, res) => {
    res.render('offset-name', { title: 'OFFSET NAME' });
});

// Cấu hình Multer để lưu trữ tệp tin được tải lên vào thư mục 'uploads'
const storage = multer.diskStorage({
    destination: (req, file, cb) => {
        cb(null, 'uploads/offset-name');
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

    
    // Xử lý tệp tin nếu cần
    const uploadedFileName = `./uploads/offset-name/${req.file.filename}`;
    let start = req.body.start;
    let end = req.body.end;
    let mainContent = req.body.mainContent;
    let splitContent = req.body.splitContent;
    let username = req.body.username;

    console.log([start,end,mainContent,splitContent,username])

    const xlsxFilePath = uploadedFileName; // Tên file input
    let workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(xlsxFilePath);
    // sheet gốc
    let worksheet = workbook.getWorksheet("Sheet1");
    
    // read w2
    const workbook2 = xlsx.readFile(xlsxFilePath);
    // Get the names of all sheets in the workbook.
    const sheetNames = workbook2.SheetNames;

    // sheet lấy data thứ tự bắt đầu từ 0
    const dataSheetName = sheetNames[2];

    // Get the first worksheet in the workbook.
    const worksheet2 = workbook2.Sheets[dataSheetName];

    
    // lấy mảng data từ sheet chưa data
    const data2 = xlsx.utils.sheet_to_json(worksheet2);
    const col1 = data2.map((item) => item.name); // cot name

    // lặp dòng từ bắt đàu đến kết thúc
    for (let i = start; i <= end; i++) {
     
        let cellC = worksheet.getCell(`${mainContent}${i}`).text;
      worksheet.getCell(`${mainContent}${i}`).value = cellC;
      let cellG = worksheet.getCell(`${splitContent}${i}`);

      let result = "";

     if(cellC.includes('Thanh toan - Ma khach hang')){
       result = 'THU HO,CHI HO VNTOPUP VNPAY - A/C:' + generateRandomNumber()
     }else if(cellC.includes(`${username} chuyen tien (`)){
       result = `${getStringBetweenOrToEnd(cellC,"(","00000000")} - A/C:${generateRandomNumber()}`; 
     }else if(cellC.includes(`${username} chuyen tien`)){
       result = `${col1[parseInt(Math.random() * col1.length)].trim()} -  A/C:${generateRandomNumber()}`;
     }else if(cellC.includes("Chuyen tien den tu NAPAS Noi dung:")){
       result = `${getStringBetweenOrToEnd(cellC,":"," chuyen khoan")} - A/C:${generateRandomNumber()}`; 
     }else if(cellC.includes("CT nhanh 247 den: QR -" && cellC.includes(" chuyen tien"))){
       result = `${getStringBetweenOrToEnd(cellC,"QR -"," chuyen tien")} - A/C:${generateRandomNumber()}`; 
     }else if(cellC.includes("CT nhanh 247 den: QR -")){
       result = "MBBANK IBFT - A/C:0345985058"; 
     }else if(cellC.includes("348H91N4820E14LY/")){
       result = `${getStringBetweenOrToEnd(cellC,"/"," chuyen tien")} - A/C:${generateRandomNumber()}`; 
     }else if(cellC.includes("Chuyen tien di qua NAPAS Noi dung:")){
       result = `${getStringBetweenOrToEnd(cellC,":"," chuyen tien")} - A/C:${generateRandomNumber()}`; 
     }else {
       result = `${getStringBetweenOrToEnd(cellC,""," chuyen tien")} - A/C:${generateRandomNumber()}`; 
     }

     cellG.value = result;
   }
    
        // Lưu file output
    const outputFilePath = `outputs/${req.file.filename}_offset-name.xlsx`;
    await workbook.xlsx.writeFile(outputFilePath);

    res.json({
      message: "success",
      outputFileName: outputFilePath, 
})
});

function generateRandomNumber() {
    let num = Math.floor(Math.random() * 9) + 1; // Ensure the first digit isn't 0
    let digits = 13; // Remaining digits
    for (let i = 0; i < digits; i++) {
        num = num * 10 + Math.floor(Math.random() * 10);
    }
    return num;
}

function getStringBetweenOrToEnd(input, startString, endString) {
    var startIndex = input.indexOf(startString);
    
    return input.substring(startIndex + 1,input.length - endString.length);
  }

module.exports = router;
