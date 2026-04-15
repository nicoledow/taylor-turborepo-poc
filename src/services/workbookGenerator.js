import ExcelJS from "exceljs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const PROJECT_ROOT = path.resolve(__dirname, "..", "..");
const TEMPLATE_PATH = path.join(PROJECT_ROOT, "templates", "workbook_template.xlsm");
const OUTPUT_DIR = path.join(PROJECT_ROOT, "generatedWorkbooks");

export async function generateWorkbook(csExportData, outputFilename) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(TEMPLATE_PATH);

  const sheet = workbook.getWorksheet("CSExport");
  if (!sheet) {
    throw new Error('Sheet "CSExport" not found in template');
  }

  populateCSExport(sheet, csExportData);

  const safeName = outputFilename.replace(/\.xlsm$/i, ".xlsx");
  const outputPath = path.join(OUTPUT_DIR, safeName);
  await workbook.xlsx.writeFile(outputPath);

  return outputPath;
}

function populateCSExport(sheet, data) {
  sheet.getCell("B4").value = data.countsAsOf;
  sheet.getCell("B5").value = data.clientName;
  sheet.getCell("B6").value = data.clientDBA;
  sheet.getCell("B7").value = data.clientNumber;
  sheet.getCell("B8").value = data.readyForClient ? "Yes" : "No";
  sheet.getCell("B9").value = data.benefitContact.name;
  sheet.getCell("B10").value = data.benefitContact.email;
  sheet.getCell("B11").value = data.benefitContact.phone;
  sheet.getCell("B12").value = data.benefitRep.name;
  sheet.getCell("B13").value = data.benefitRep.email;
  sheet.getCell("B14").value = data.benefitRep.phone;
  sheet.getCell("B15").value = data.benefitYear;
}
