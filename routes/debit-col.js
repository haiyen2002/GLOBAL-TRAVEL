const express = require("express");
const router = express.Router();
const ExcelJS = require("exceljs");
const fs = require('fs');

router.get("/", async (req, res) => {
  try {
    const xlsxFilePath = "debit-col/input/debit_col.xlsx";
    let workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(xlsxFilePath);

    let sheet = workbook.getWorksheet("Sheet1");

    const cellsToFillF = [];
    const cellsToFillH = [];

    for (let row = 20; row <= 567; row++) {
      const cellAddressF = `F${row}`;
      const cellAddressH = `H${row}`;

      const cellF = sheet.getCell(cellAddressF);
      const cellH = sheet.getCell(cellAddressH);

      if (!cellF.value && !cellH.value) {
        cellsToFillF.push(cellAddressF);
        cellsToFillH.push(cellAddressH);
      } else if (cellF.value && !cellH.value) {
        cellsToFillF.push(cellAddressF);
      }
    }

    // Điền giá trị ngẫu nhiên vào 200 ô từ danh sách cellsToFillF
    for (let i = 0; i < 200; i++) {
      const randomIndex = Math.floor(Math.random() * cellsToFillF.length);
      const randomCellF = cellsToFillF.splice(randomIndex, 1)[0];

      const randomNumberF = Math.floor(Math.random() * (10000000 - 200000 + 1)) + 200000;
      sheet.getCell(randomCellF).value = randomNumberF;
    }

    // Điền giá trị ngẫu nhiên vào cột H khi có điều kiện
    cellsToFillH.forEach(cellAddressH => {
      const cellAddressF = `F${cellAddressH.substring(1)}`;
      const cellF = sheet.getCell(cellAddressF);

      if (!cellF.value) {
        const randomNumberH = Math.floor(Math.random() * (10000000 - 100000 + 1)) + 100000;
        sheet.getCell(cellAddressH).value = randomNumberH;
      }
    });

    const buffer = await workbook.xlsx.writeBuffer();
    fs.writeFileSync("debit-col/output/debit_col_output.xlsx", buffer);

    res.json({
      message: "success",
    });
  } catch (error) {
    res.json({
      message: error.message,
    });
  }
});

module.exports = router;
