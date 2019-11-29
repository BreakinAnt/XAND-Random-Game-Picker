var Excel = require('exceljs');

var wb = new Excel.Workbook();
var scan = process.stdin;

scan.setEncoding('utf-8');

console.log("Please input .xlsx filename: ");

scan.on('data', data => {
    const excelFile = data.replace('\r\n', '');
    const path = `./excel/${excelFile}.xlsx`;
   
    wb.xlsx.readFile(path).then(wb => {
        const rowSize = wb.getWorksheet("Planilha1").rowCount;
        const randomCellNumber = Math.floor((Math.random() * rowSize) + 1);
	    
        const gameName = wb.getWorksheet("Planilha1").getRow(randomCellNumber).getCell(2).value;
	const gameYear = wb.getWorksheet("Planilha1").getRow(randomCellNumber).getCell(3).value;
	const gamePublisher = wb.getWorksheet("Planilha1").getRow(randomCellNumber).getCell(4).value;
        const gameGenre = wb.getWorksheet("Planilha1").getRow(randomCellNumber).getCell(5).value;
	    
        console.log(`--\nName: ${gameName}\n`);
	console.log(`Year: ${gameYear}\n`);
	console.log(`Publisher: ${gamePublisher}\n`);
	console.log(`Genre: ${gameGenre}\n--`);
    }).catch(() => console.log(`No ${excelFile}.xlsx found.`));
});
