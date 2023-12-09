const express = require("express");
const router = express.Router();
const ExcelJS = require("exceljs");

router.get("/", async (req, res) => {
  try {
    const xlsxFilePath = "stop-random/input/bidv.xlsx"; // Tên file input
    let workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(xlsxFilePath);

    let sheet = workbook.getWorksheet("Sheet1");

    for (let i = 23; i <= 42; i++) {
        let cellC = sheet.getCell(`K${i}`).text;
      sheet.getCell(`K${i}`).value = cellC;
    }

    for (let i = 44; i <= 73; i++) {
      let cellC = sheet.getCell(`K${i}`).text;
    sheet.getCell(`K${i}`).value = cellC;
  }
  for (let i = 75; i <= 102; i++) {
    let cellC = sheet.getCell(`K${i}`).text;
  sheet.getCell(`K${i}`).value = cellC;
}

for (let i = 104; i <= 131; i++) {
  let cellC = sheet.getCell(`K${i}`).text;
sheet.getCell(`K${i}`).value = cellC;
}

for (let i = 133; i <= 161; i++) {
  let cellC = sheet.getCell(`K${i}`).text;
sheet.getCell(`K${i}`).value = cellC;
}

for (let i = 163; i <= 194; i++) {
  let cellC = sheet.getCell(`K${i}`).text;
sheet.getCell(`K${i}`).value = cellC;
}

for (let i = 196; i <= 227; i++) {
  let cellC = sheet.getCell(`K${i}`).text;
sheet.getCell(`K${i}`).value = cellC;
}

for (let i = 229; i <= 257; i++) {
  let cellC = sheet.getCell(`K${i}`).text;
sheet.getCell(`K${i}`).value = cellC;
}

for (let i = 259; i <= 287; i++) {
  let cellC = sheet.getCell(`K${i}`).text;
sheet.getCell(`K${i}`).value = cellC;
}

for (let i = 289; i <= 317; i++) {
  let cellC = sheet.getCell(`K${i}`).text;
sheet.getCell(`K${i}`).value = cellC;
}

for (let i = 319; i <= 347; i++) {
  let cellC = sheet.getCell(`K${i}`).text;
sheet.getCell(`K${i}`).value = cellC;
}


    // Lưu file output
    const outputFilePath = "stop-random/output/stop-bidv-output.xlsx"; // Tên file output
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
