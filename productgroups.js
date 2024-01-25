var _ = require('lodash');

const excelToJson = require('convert-excel-to-json');

const excel = require('exceljs');  
let workbook = new excel.Workbook();
  const result = excelToJson({
    sourceFile: 'productGroup.xls'
  });

  let data = result["DistributorServiceAreas-TradeDe"];
  console.log("print data");
  console.log("result")
  console.log("print result")
  console.log("data")
  

  // {
  //   A: '_id',
  //     B: 'name',
  //   C: 'status',
  //   D: 'producerName',
  //   E: 'producerId',
  //   F: 'serviceAreaCodes',
  //   G: 'state',
  //   H: 'lga',
  //   I: 'address1',
  //   J: 'classifications'
  // },

   const bank = []

  data = _.filter(data, (d) => {
    return d['A'] !== '_id'
  });

  for (const d of data) {
    let codes = d['F'];
    if (codes && typeof codes === 'string') {
      codes = codes.split(';')
      codes = _.map(codes, function (code) {
        return code.substring(0, 6);
      });
      codes = codes.join(';')
    }
    bank.push({
      _id: d['A'],
      name: d['B'],
      status: d['C'],
      producerName: d['D'],
      producerId: d['E'],
      serviceAreaCodes: codes,
      state: d['G'],
      lga: d['H'],
      address1: d['I'],
      classifications: d['J'],
    })
  }

  let balanceSheet = workbook.addWorksheet('productGroup');

  balanceSheet.columns = [
    { header: '_id', key: '_id', width: 30 },
    { header: 'name', key: 'name', width: 90 },
    { header: 'status', key: 'status', width: 90 },
    { header: 'producerName', key: 'producerName', width: 90 },
    { header: 'producerId', key: 'producerId', width: 90 },
    { header: 'serviceAreaCodes', key: 'serviceAreaCodes', width: 90 },
    { header: 'state', key: 'state', width: 90 },
    { header: 'lga', key: 'lga', width: 90 },
    { header: 'address1', key: 'address1', width: 90 },
    { header: 'classifications', key: 'classifications', width: 90 },
  ];

  balanceSheet.addRows(bank);

  workbook.xlsx.writeFile("productGroupF.xlsx")
    .then(function() {
      console.log("file saved!");
    });