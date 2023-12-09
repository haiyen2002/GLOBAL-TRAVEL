require('dotenv').config()
const express = require('express');
const app = express();
const PORT = process.env.PORT || 3003;
const cors = require("cors");
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const xlsx = require('xlsx');
const ExcelJS = require('exceljs');

// routers
const offsetNameRouter = require('./routes/offset-name-1')
const debitColRouter = require('./routes/debit-col')
const editHeightRouter = require('./routes/edit-height')
const randomSkDuyet = require('./routes/random-duyet')
const stopRandom = require('./routes/stop-random')
const randomDebitCredit = require('./routes/random-debit-credit');
const offsetName = require('./routes/offset-name');

const bodyParser = require("body-parser");
app.use(bodyParser.urlencoded({extended:true})); 
app.set("view engine","ejs");
app.set("views","./views");
app.use(express.static('public'));

const fs = require("fs");
const path = require("path");


// Cors
app.use(cors());

// Middleware
app.use(express.json());

// Routes

app.use('/offset-name-1', offsetNameRouter)
app.use('/debit-col', debitColRouter)
app.use('/edit-height', editHeightRouter)
app.use('/random-sk-duyet', randomSkDuyet)
app.use('/stop-random', stopRandom)
app.use('/random-debit-credit',randomDebitCredit)
app.use('/offset-name',offsetName)



app.get('/excel', async (req, res) => {
    function getStringBetweenOrToEnd(input, startString, endString) {
        var startIndex = input.indexOf(startString);
        
        if (startIndex === -1 || !startString) {
          var endIndex = input.indexOf(endString);
          if (endIndex === -1) {
            return "End string not found in the input";
          }
          return input.substring(0, endIndex);
        }
      
        startIndex += startString.length;
      
        var endIndex = input.indexOf(endString, startIndex);
        if (endIndex === -1) {
          return "End string not found in the input after start string";
        }
      
        return Math.random() * 10 > 6 ? "VND-TGTT-" + input.substring(startIndex, endIndex) + "VIETNAM" : input.substring(startIndex, endIndex);
      }

    try {
        const xlsxFilePath = 'SK-UY-TEST.xlsx'
        let workbook = new ExcelJS.Workbook();
await workbook.xlsx.readFile(xlsxFilePath);
let worksheet = workbook.getWorksheet('Sheet1');


// read w2
const workbook2 = xlsx.readFile(xlsxFilePath);
    // Get the names of all sheets in the workbook.
    const sheetNames = workbook2.SheetNames;

    // Assume we want the first sheet. You can choose a different sheet if needed.
    const firstSheetName = sheetNames[3];

    // Get the first worksheet in the workbook.
    const worksheet2 = workbook2.Sheets[firstSheetName];

    // Convert the worksheet data into an array of objects.
    const data2 = xlsx.utils.sheet_to_json(worksheet2);
    const col1 = data2.map(item => item.name)
    const col2 = data2.map(item => item.column2).filter(item => item)
    const col3 = data2.map(item => item.column3).filter(item => item)

for(let i = 11; i <= 661; i++){
    let cellE = worksheet.getCell(`E${i}`).text;
    worksheet.getCell(`E${i}`).value = cellE;
    let cellG = worksheet.getCell(`G${i}`);

    let result = ""
    
    if (cellE.includes("chuyen tien") && cellE.includes(";")){
        result = getStringBetweenOrToEnd(cellE, ";", "chuyen tien")
    }else if(cellE.includes("chuyen tien")){
        result = getStringBetweenOrToEnd(cellE, undefined, "chuyen tien")
    }else if(cellE.includes("chuyen khoan")){
        result = getStringBetweenOrToEnd(cellE, "-", "chuyen khoan")
    }else{
        result = ""
    }

    cellG.value = result.includes("NGUYEN QUOC UY") ? result.replace("NGUYEN QUOC UY", col1[parseInt(Math.random() * col1.length)]).trim() : result.trim()


}

await workbook.xlsx.writeFile('test-edited.xlsx');

        res.json({
            message: 'success'
        })
    } catch (error) {
        res.json({

            message: error.message
        })
    }
})

app.get('/home', (req, res)=>{
    res.render("home.ejs");
})

app.get('/', async (req, res) => {
   try {
    const dataFile = req.query.file
    const subName = req.query.sub_name

    if(!dataFile){
        return res.status(400).json({
            message: "Truyền tên file data"
        })
    }
    // Get the path to the XLSX file.
    const xlsxFilePath = `./data/${dataFile}.xlsx`;

    // Read the XLSX file.
    const workbook = xlsx.readFile(xlsxFilePath);

    // Get the names of all sheets in the workbook.
    const sheetNames = workbook.SheetNames;

    // Assume we want the first sheet. You can choose a different sheet if needed.
    const firstSheetName = sheetNames[0];

    // Get the first worksheet in the workbook.
    const worksheet = workbook.Sheets[firstSheetName];

    // Convert the worksheet data into an array of objects.
    const data = xlsx.utils.sheet_to_json(worksheet);


    // Load the docx file as binary content
    
    // Replace the string in the text.
    data.forEach(async (row, index) => {
        const content = fs.readFileSync(
            path.resolve(__dirname, `./template/${row.template}.docx`),
            "binary"
        );

        const zip = new PizZip(content);

        const doc = new Docxtemplater(zip, {
            paragraphLoop: true,
            linebreaks: true,
        });

        // Render the document (Replace {first_name} by John, {last_name} by Doe, ...)
        doc.render(row);

        const buf = doc.getZip().generate({
            type: "nodebuffer",
            // compression: DEFLATE adds a compression step.
            // For a 50MB output document, expect 500ms additional CPU time
            compression: "DEFLATE",
        });

        // buf is a nodejs Buffer, you can either write it to a
        // file or res.send it with express for example.
        fs.writeFileSync(path.resolve(__dirname, `./output/${row.template}-${row["tên lđ"]}-${subName && row[subName] ? row[subName] : index}.docx`), buf);
    })

    res.status(200).json({
        data: "Thành công"
    })
   } catch (error) {
    res.status(500).json({
        error: error.message
    })
   }
});

// Start the server
app.listen(PORT, () => {
    console.log(`Server is running on port ${PORT}`)
})
// })