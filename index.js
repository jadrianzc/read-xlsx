const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');
const XlsxPopulate = require('xlsx-populate');
const dotenv = require('dotenv').config();
const { areasIess } = require('./dataAreas');

// Paths excels
const pathExcel = path.join(__dirname, '/files/inventario.xlsx');
const pathExcelPlantilla = path.join(__dirname, '/files/plantilla.xlsx');

// Lee excel
const workBook = xlsx.readFile(pathExcel);
let workSheets = workBook.SheetNames;
console.log(`Iniciando la generación de reportes de la hoja ${workSheets[0]}`);

// Data excel
const workSheetData = workBook.Sheets[workSheets[0]];
const fileData = xlsx.utils.sheet_to_json(workSheetData, {
	header: 1,
});

const headers = [
	'ctd_estancias',
	'cod_estancias',
	'description',
	'ctd_equipos',
	'cod_equipos',
	'marca',
	'modelo',
	'coste_unitario',
	'num_serie',
];

fileData.splice(0, 4);

// Methods
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

const handleDataReport = (rows, codEstancia, codEquipo) => {
	const dataFilter = rows.filter(
		(row) =>
			row.cod_estancias?.startsWith(codEstancia) && row.cod_equipos?.startsWith(codEquipo)
	);

	return dataFilter;
};

const generateReportXlsx = (dataFile, title) => {
	fs.mkdirSync(`archivosGenerados/${title}/`, { recursive: true });

	// Load an existing workbook
	for (const data of dataFile) {
		XlsxPopulate.fromFileAsync(pathExcelPlantilla)
			.then((workbook) => {
				const sheet1 = workbook.sheet(0);
				const cantidadHojas = data?.ctd_equipos;

				for (let index = 1; index <= cantidadHojas - 1; index++) {
					workbook.cloneSheet(sheet1, `HOJA ${index + 1}`);
				}

				const cellsValid = [
					{ cell: 2, value: data?.description },
					{ cell: 3, value: data?.marca },
					{ cell: 4, value: data?.modelo },
					{ cell: 5, value: data?.num_serie || 'S/N' },
					{ cell: 7, value: '' },
					{ cell: 9, value: title },
				];

				for (let index = 0; index < cantidadHojas; index++) {
					for (const cells of cellsValid) {
						workbook.sheet(index).row(6).cell(cells.cell).value(cells.value);
					}
				}

				workbook.activeSheet(0);
				const nameFile = `archivosGenerados/${title}/${data?.description
					.replace(/\//g, ' ')
					.toUpperCase()}.xlsx`;

				workbook.toFileAsync(nameFile);
			})
			.catch((error) => console.log(error));
	}

	console.log(`${dataFile.length} resgistros creados exitosamente.`);
};

(async () => {
	const rows = await convertToJson(headers, fileData);

	for (const area of areasIess) {
		const { title, codEstancia, codEquipo } = area;

		const cooEm = handleDataReport(rows, codEstancia, codEquipo);
		generateReportXlsx(cooEm, title);
	}
})();
