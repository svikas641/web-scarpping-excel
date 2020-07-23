const Nightmare = require('nightmare')
const cheerio = require('cheerio');
const nightmare = Nightmare({ show: true });
const Excel = require("exceljs");
const readln = require('readline');

var cl = readln.createInterface(process.stdin, process.stdout);
var question = function (q) {
  return new Promise((res, rej) => {
    cl.question(q, answer => {
      res(answer);
    })
  });
};


async function go(answer) {

  try {
    const url = 'https://ngodarpan.gov.in/index.php/home/statewise_ngo/7255/7/1?per_page=100';
    const response = await nightmare
      .goto(url)
      .wait('body')
      .click(`.table tbody tr:nth-child(${answer}) a`)
      .wait(5000)
      .evaluate(() => document.querySelector('div#ngo_info_modal.modal.fade.in').innerHTML)
      .end();

    const result =  getData(response);
    console.log("Got Data");

    const newWorkbook = new Excel.Workbook();
    await newWorkbook.xlsx.readFile("./export2.xlsx");

    const newworksheet = newWorkbook.getWorksheet("Sheet1");
    newworksheet.columns = [
      { header: "S.NO", key: "sno", width: 10 },
      { header: "UNIQUE ID", key: "uid", width: 25 },
      { header: "TYPE OF NGO", key: "tngo", width: 40 },
      { header: "REGISTRATION NO.", key: "rgno", width: 40 },
      { header: "DATE OF REGISTRATION", key: "dorg", width: 40 },
      { header: "AVAILABILITY OF FCRA", key: "aof", width: 40 },
      { header: "ADDRESS", key: "add", width: 50 },
      { header: "CONTACT NO.", key: "cno", width: 30 },
      { header: "WEBSITE", key: "web", width: 40 },
      { header: "EMAIL", key: "email", width: 50 },
    ];
    await newworksheet.addRow({
      sno: answer,
      uid: result.Unique_Id,
      tngo: result.Type_of_Ngo,
      rgno: result.Registration_No,
      dorg: result.Date_of_Registration,
      aof: result.FCRA,
      add: result.Address,
      cno: result.Contact_No,
      web: result.Website,
      email: result.Email,
    });

    await newWorkbook.xlsx.writeFile("export2.xlsx");

    console.log("File is written");

  } catch (error) {
    console.log(error)
  }

}

var getData =  html => {
  var data = {};
  const $ = cheerio.load(html);

  data = {
    Unique_Id: $("#UniqueID").text(),
    Type_of_Ngo: $("#ngo_type").text(),
    Registration_No: $("#ngo_regno").text(),
    Date_of_Registration: $("#ngo_reg_date").text(),
    FCRA: $("#FCRA_reg_no").text(),
    Address: $("#address").text(),
    Contact_No: $("#mobile_n").text(),
    Website: $("#ngo_web_url").text(),
    Email: $("#email_n").text(),
  };

  return data;
}

(async function main() {
  var answer;
  answer = await question('Enter Number: ');
  await go(answer);
  cl.close();
})();

