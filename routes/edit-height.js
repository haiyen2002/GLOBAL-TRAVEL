const express = require("express");
const router = express.Router();
const ExcelJS = require("exceljs");

router.get("/", async (req, res) => {
  try {
    const xlsxFilePath = "edit-height/input/stop-DUYET-output.xlsx";
    let workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(xlsxFilePath);

    let sheet = workbook.getWorksheet("Sheet1");

  // Tăng chiều cao cho các ô từ B13 đến B230
  for (let row = 12; row <= 662; row++) {
    const rowHeight = sheet.getRow(row).height;
      sheet.getRow(row).height = rowHeight + 10; 
    }
    // Tăng chiều cao mỗi ô lên 2px
    // const cell = sheet.getCell(`C${row}`);
    // if (cell.value.includes("SML-ECOM HANOI VNM")) {
    //   sheet.getRow(row).height = 22; // Nếu ô chứa chuỗi "SML-ECOM HANOI VNM", set chiều cao của hàng là 26
    // }
      

    const outputFilePath = "edit-height/output/sk-duyet-height-output.xlsx";
    await workbook.xlsx.writeFile(outputFilePath);

    res.json({
      message: "success",
      outputFileName: outputFilePath,
    });
  } catch (error) {
    res.json({
      message: error.message,
    });
  }
});

module.exports = router;
