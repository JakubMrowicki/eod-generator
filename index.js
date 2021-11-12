const XLSX = require('xlsx');
const express = require("express");
const app = express();
const path = require('path');
app.use(express.json());
app.use(express.static(path.join(__dirname, 'build')));

function buildReport(data) {
    var r_opts = { bookType:'xlsx', cellStyles:true, sheetStubs: true };
    var workbook = XLSX.readFile(`${__dirname}/eodTemplate.xlsx`, r_opts);
    var newWorkbook = workbook;
    var sheet = newWorkbook.Sheets['eod'];

    //Data
    const { centre, booths } = data;

    sheet['B2'].v = centre.location;
    var [year, month, day] = centre.date.split("-");
    var date = `${day}/${month}/${year}`;
    sheet['B3'].v = date;
    sheet['C5'].v = centre.details;
    sheet['B16'].v = centre.lip;
    sheet['B17'].v = centre.apdirect;
    sheet['B18'].v = centre.resubs;
    sheet['B24'].v = centre.bags.green;
    sheet['B25'].v = centre.bags.blue;
    sheet['B26'].v = parseInt(centre.bags.blue) + parseInt(centre.bags.green);
    sheet['C24'].v = centre.dxlabel;
    sheet['C26'].v = centre.dxseal;
    sheet['B30'].v = centre.vo;
    sheet['B31'].v = centre.vo;
    sheet['B32'].v = centre.vo;
    sheet['B33'].v = centre.vo;
    sheet['B34'].v = centre.vo;
    sheet['B35'].v = centre.vo;
    sheet['B36'].v = centre.vo;

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
    sheet['B5'].v = credit;
    sheet['B6'].v = debit;
    sheet['B7'].v = payzone;
    sheet['B8'].v = mobile;
    sheet['B9'].v = credit + debit + payzone + mobile;
    sheet['B20'].v = cardmachine;
    sheet['B12'].v = exp1;
    sheet['B13'].v = exp2;

    // Build File Name
    var number = Math.floor(Math.random() * 1000);
    var filename = `EOD_${centre.location}_${centre.date}_${number}.xlsx`;
    XLSX.writeFile(newWorkbook, `${__dirname}/archive/${filename}`);
    return filename;
}

app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'build', 'index.html'));
});

app.get('/download/:fileName', function(req, res){
    const file = `${__dirname}/archive/${req.params.fileName}`;
    res.download(file);
  });

app.post("/submit", (req, res) => {
    const file = buildReport(req.body);
    res.send({ file: `${file}` });
;});

const PORT = process.env.PORT || 3001;
app.listen(PORT, () => {
    console.log(`Server listening on ${PORT}`);
});