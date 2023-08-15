const path = require('path');
const xlsx = require('xlsx');
const dotenv = require('dotenv').config();

const pathExcel = path.join(__dirname, '/files/inventario.xlsx');

const convertToJson = async (headers, data) => {
	const rows = [];

	data.forEach(async (row) => {
		const rowData = {};
		row.forEach(async (element, index) => {
			rowData[headers[index]] = element;
		});
		rows.push(rowData);
	});

	return rows;
};

const workBook = xlsx.readFile(pathExcel);
let workSheets = workBook.SheetNames;
console.log(workSheets);
const workSheetData = workBook.Sheets[workSheets[0]];
const fileData = xlsx.utils.sheet_to_json(workSheetData, {
	header: 1,
	blankRows: false,
});

const headers = fileData[3];
fileData.splice(0, 4);

const handleDataReport = (rows, codEstancia, codEquipo) => {
	const dataFilter = rows.filter(
		(row) => row['Cod. Estancia ']?.startsWith('CO') && row['Cod. Equipo']?.startsWith('EM')
	);

	return dataFilter;
};

(async () => {
	const rows = await convertToJson(headers, fileData);
	console.log(rows);

	const cooEm = handleDataReport(rows, 'CO', 'EM');
	console.log(cooEm);
	console.log(cooEm.length);
})();
