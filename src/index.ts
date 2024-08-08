import ExcelJS, { Cell, CellValue, Row, Workbook, Worksheet } from "exceljs";
import HyperFormula from "hyperformula";

async function generateExcels(filename: string) {
  const xlsxWorkbook = await readXlsxWorkbookFromFile(filename);
  const sheetsAsJavascriptArrays =
    convertXlsxWorkbookToJavascriptArrays(xlsxWorkbook);
  const hf = HyperFormula.buildFromSheets(sheetsAsJavascriptArrays, {
    licenseKey: "gpl-v3",
  });

  console.log("Formulas:", hf.getSheetSerialized(0));
  console.log("Values:  ", hf.getSheetValues(0));
}

async function readXlsxWorkbookFromFile(filename: string) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filename);
  return workbook;
}

function convertXlsxWorkbookToJavascriptArrays(workbook: Workbook) {
  const workbookData = {} as any;

  workbook.eachSheet((worksheet: Worksheet) => {
    const sheetData = [] as any[];

    worksheet.eachRow((row: Row) => {
      const rowData = [] as any[];

      row.eachCell({ includeEmpty: true }, (cell: Cell) => {
        const cellValue: CellValue = cell.value;
        const cellFormula = cell.formula || (cellValue as any)?.formula;
        const cellData = cellFormula ? `=${cellFormula}` : cellValue;
        rowData.push(cellData);
      });

      sheetData.push(rowData);
    });

    workbookData[worksheet.name] = sheetData;
  });

  return workbookData;
}

generateExcels("sample_file.xlsx");
