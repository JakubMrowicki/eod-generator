const Excel = require('exceljs');
const express = require("express");
const path = require('path');
const fs = require('fs');
const app = express();
app.use(express.json());
app.use(express.static(path.join(__dirname, 'build')));

async function buildReport(data) {
    const workbook = new Excel.Workbook();
    await workbook.xlsx.readFile(`${__dirname}/eodTemplate.xlsx`);
    const worksheet = workbook.worksheets[0];

    // Data
    const { centre, booths } = data;

    worksheet.getCell('B2').value = centre.location;
    var [year, month, day] = centre.date.split("-");
    date = `${day}/${month}/${year}`;
    dateFile = `${day}_${month}_${year}`;
    worksheet.getCell('B3').value = date;
    worksheet.getCell('C5').value = centre.details;
    worksheet.getCell('B16').value = centre.lip;
    worksheet.getCell('B17').value = centre.apdirect;
    worksheet.getCell('B18').value = centre.resubs;
    worksheet.getCell('B24').value = centre.bags.green;
    worksheet.getCell('B25').value = centre.bags.blue;
    worksheet.getCell('B26').value = parseInt(centre.bags.blue) + parseInt(centre.bags.green);
    worksheet.getCell('C24').value = centre.dxlabel;
    worksheet.getCell('C26').value = centre.dxseal;
    worksheet.getCell('B30').value = centre.vo;
    worksheet.getCell('B31').value = centre.vo;
    worksheet.getCell('B32').value = centre.vo;
    worksheet.getCell('B33').value = centre.vo;
    worksheet.getCell('B34').value = centre.vo;
    worksheet.getCell('B35').value = centre.vo;
    worksheet.getCell('B36').value = centre.vo;

    //Booths
    var credit = 0;
    var debit = 0;
    var payzone = 0;
    var mobile = 0;
    var cardmachine = 0;
    var exp1 = 0;
    var exp2 = 0;
    for(var booth in booths) {
        credit += parseInt(booths[booth].credit)
        debit += parseInt(booths[booth].debit)
        payzone += parseInt(booths[booth].payzone)
        mobile += parseInt(booths[booth].mobile)
        cardmachine += parseInt(booths[booth].cardmachine)
        exp1 += parseInt(booths[booth].exp1)
        exp2 += parseInt(booths[booth].exp2)
    }
    worksheet.getCell('B5').value = credit;
    worksheet.getCell('B6').value = debit;
    worksheet.getCell('B7').value = payzone;
    worksheet.getCell('B8').value = mobile;
    worksheet.getCell('B9').value = credit + debit + payzone + mobile;
    worksheet.getCell('B20').value = cardmachine;
    worksheet.getCell('B12').value = exp1;
    worksheet.getCell('B13').value = exp2;

    // Build File
    var filename = `EOD_${centre.location}_${dateFile}.xlsx`;
    await workbook.xlsx.writeFile(`${__dirname}/archive/${filename}`);
    return filename;
}

async function buildSheet(data) {
    const workbook = new Excel.Workbook();
    await workbook.xlsx.readFile(`${__dirname}/eodTemplate2.xlsx`);
    const worksheet = workbook.worksheets[0];

    // Data
    const { centre, booths } = data;
    // console.log(booths);

    worksheet.getCell('A1').value = `End of Day Reconciliation - ${centre.location}`;
    worksheet.getCell('G1').value = `BARCODE: ${centre.dxlabel}`;
    worksheet.getCell('G2').value = `SEAL NO: ${centre.dxseal}`;
    var [year, month, day] = centre.date.split("-");
    date = `${day}/${month}/${year}`;
    dateFile = `${day}_${month}_${year}`;
    worksheet.getCell('M1').value = `DATE: ${date}`;

    //Booths
    var credit = 0;
    var debit = 0;
    var payzone = 0;
    var mobile = 0;
    var cardmachine = 0;
    var exp1 = 0;
    var exp2 = 0;
    var diff = 0;
    for(var booth in booths) {
        credit += parseInt(booths[booth].credit)
        debit += parseInt(booths[booth].debit)
        payzone += parseInt(booths[booth].payzone)
        mobile += parseInt(booths[booth].mobile)
        cardmachine += parseInt(booths[booth].cardmachine)
        exp1 += parseInt(booths[booth].exp1)
        exp2 += parseInt(booths[booth].exp2)
        diff += parseInt(booths[booth].diff)
        const boothNum = parseInt(booth) + 7;
        worksheet.getCell(`B${boothNum}`).value = booths[booth].name;
        worksheet.getCell(`E${boothNum}`).value = parseInt(booths[booth].credit);
        worksheet.getCell(`F${boothNum}`).value = parseInt(booths[booth].debit);
        worksheet.getCell(`G${boothNum}`).value = parseInt(booths[booth].payzone);
        worksheet.getCell(`H${boothNum}`).value = parseInt(booths[booth].mobile);
        worksheet.getCell(`I${boothNum}`).value = parseInt(booths[booth].cardmachine);
        worksheet.getCell(`J${boothNum}`).value = parseInt(booths[booth].mobile) + parseInt(booths[booth].payzone) + parseInt(booths[booth].debit) + parseInt(booths[booth].credit);
        worksheet.getCell(`M${boothNum}`).value = parseInt(booths[booth].exp1);
        worksheet.getCell(`N${boothNum}`).value = parseInt(booths[booth].exp2);
        worksheet.getCell(`O${boothNum}`).value = parseInt(booths[booth].diff);
    }
    worksheet.getCell('E17').value = credit;
    worksheet.getCell('F17').value = debit;
    worksheet.getCell('G17').value = payzone;
    worksheet.getCell('H17').value = mobile;
    worksheet.getCell('J17').value = credit + debit + payzone + mobile;
    worksheet.getCell('I17').value = cardmachine;
    worksheet.getCell('M17').value = exp1;
    worksheet.getCell('N17').value = exp2;
    worksheet.getCell('O17').value = diff;
    worksheet.getCell('B19').value = centre.details;

    // Build File
    var filename = `EOD_FILING_SHEET_${centre.location}_${dateFile}.xlsx`;
    await workbook.xlsx.writeFile(`${__dirname}/archive/${filename}`);
    return filename;
}

app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'build', 'index.html'));
});

app.get('/download/:fileName', function(req, res){
    const file = `${__dirname}/archive/${req.params.fileName}`;
    res.download(file, function(err) {
        if (err) {
          console.log(err);
        }
        fs.unlink(file, function(){
        });
      });
});

app.post("/submit", (req, res) => {
    buildReport(req.body).then((file) => {
        res.send({ file: `${file}` });
    });
});

app.post("/submit2", (req, res) => {
    buildSheet(req.body).then((file) => {
        res.send({ file: `${file}` });
    });
});

const PORT = process.env.PORT || 3001;
app.listen(PORT, () => {
    console.log(`Server listening on ${PORT}`);
});