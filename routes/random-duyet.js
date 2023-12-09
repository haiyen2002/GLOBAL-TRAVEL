const express = require("express");
const router = express.Router();
const ExcelJS = require("exceljs");


router.get("/", async (req, res) => {
  try {
    const xlsxFilePath = "sk-duyet-random/input/SK-hoa.xlsx"; // Tên file input
    let workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(xlsxFilePath);

    let sheet = workbook.getWorksheet("Sheet1");
    function getRandomValue() {
      let randomValue = Math.floor(Math.random() * (1500000 - 100000 + 1)) + 100000; // Tạo giá trị ngẫu nhiên từ 100,000 đến 1,500,000
      randomValue = Math.round(randomValue / 1000) * 1000; // Làm tròn đến phần nghìn
      return randomValue;
    }

    const rowsToFillC = [];
    for (let row = 12; row <= 662; row++) {
      rowsToFillC.push(row);
    }
    const randomRowsC = getRandom(rowsToFillC, 270);
    for (let i = 0; i < 270; i++) {
      let randomValueC = getRandomValue();
      sheet.getCell(`D${randomRowsC[i]}`).value = randomValueC;
    }

    const rowsToFillD = [];
    for (let row = 12; row <= 662; row++) {
      if (!sheet.getCell(`D${row}`).value) {
        rowsToFillD.push(row);
      }
    }
    for (let i = 0; i < rowsToFillD.length; i++) {
      let randomValueD = getRandomValue();
      sheet.getCell(`C${rowsToFillD[i]}`).value = randomValueD;
    }

    // Lưu file output
    const outputFilePath = "sk-duyet-random/output/sk-hoa-random-output.xlsx"; // Tên file output
    await workbook.xlsx.writeFile(outputFilePath);

    res.json({
      message: "success",
      outputFileName: outputFilePath, // Trả về tên file output
    });
  } catch (error) {
    res.json({
      message: error.message,
    });
  }
});

// Hàm lấy một mảng ngẫu nhiên từ một mảng ban đầu
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
