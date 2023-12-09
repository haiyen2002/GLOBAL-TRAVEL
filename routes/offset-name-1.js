const express = require("express");
const router = express.Router();
const xlsx = require("xlsx");
const ExcelJS = require("exceljs");
const xlsxPopulate = require('xlsx-populate');

router.get("/", async (req, res) => {
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

  try {
    const xlsxFilePath = "offset-name/input/sk-phuong-1.xlsx";
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
    for (let i = 20; i <= 567; i++) {
     
        let cellC = worksheet.getCell(`C${i}`).text;
      worksheet.getCell(`C${i}`).value = cellC;
      let cellG = worksheet.getCell(`G${i}`);

      let result = "";


      if(cellC.includes('Thanh toan - Ma khach hang')){
        result = 'THU HO,CHI HO VNTOPUP VNPAY - A/C:' + generateRandomNumber()
      }else if(cellC.includes("LE THI PHUONG chuyen tien (")){
        result = `${getStringBetweenOrToEnd(cellC,"(","00000000")} - A/C:${generateRandomNumber()}`; 
      }else if(cellC.includes("LE THI PHUONG chuyen tien")){
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

    await workbook.xlsx.writeFile("offset-name/output/sk_phuong_output.xlsx");

    res.json({
      message: "success",
    });
  } catch (error) {
    res.json({
      message: error.message,
    });
  }
});

router.get("/edit-row-height", async (req, res) => {
  const xlsxFilePath = "offset-name/input/sk-phuong-1.xlsx";
  try {
     xlsxPopulate.fromFileAsync(xlsxFilePath)
    .then( async (workbook )=> {
        const sheet = workbook.sheet("Sheet1");
        sheet.row(1).height(50); // Set the height of the first row to 50
        await workbook.toFileAsync(xlsxFilePath.replace(".xlsx", "_output.xlsx"));
        return res.json({
          message: "success",
        });
    }).catch((error) => {
      res.json({
        message: error.message,
      });
    })

    
  } catch (error) {
    res.json({
      message: error.message,
    });
  }
})

module.exports = router
