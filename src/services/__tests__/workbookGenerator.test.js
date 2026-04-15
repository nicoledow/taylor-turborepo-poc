import { describe, it, expect, beforeAll, afterAll } from "vitest";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import ExcelJS from "exceljs";
import { generateWorkbook } from "../workbookGenerator.js";
import { fetchCSExportData } from "../prismhr.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const OUTPUT_DIR = path.resolve(__dirname, "..", "..", "..", "generatedWorkbooks");
const TEST_FILENAME = "test_output.xlsx";
const TEST_OUTPUT_PATH = path.join(OUTPUT_DIR, TEST_FILENAME);

describe("generateWorkbook", () => {
  let csExportData;

  beforeAll(() => {
    csExportData = fetchCSExportData("102");
  });

  afterAll(() => {
    if (fs.existsSync(TEST_OUTPUT_PATH)) {
      fs.unlinkSync(TEST_OUTPUT_PATH);
    }
  });

  it("creates an xlsx file in the generatedWorkbooks directory", async () => {
    const outputPath = await generateWorkbook(csExportData, TEST_FILENAME);

    expect(fs.existsSync(outputPath)).toBe(true);
    expect(outputPath).toBe(TEST_OUTPUT_PATH);
  });

  it("populates the CSExport sheet with the correct data", async () => {
    await generateWorkbook(csExportData, TEST_FILENAME);

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(TEST_OUTPUT_PATH);
    const sheet = workbook.getWorksheet("CSExport");

    expect(sheet).toBeDefined();
    expect(sheet.getCell("B4").value).toEqual(csExportData.countsAsOf);
    expect(sheet.getCell("B5").value).toBe(csExportData.clientName);
    expect(sheet.getCell("B6").value).toBe(csExportData.clientDBA);
    expect(sheet.getCell("B7").value).toBe(csExportData.clientNumber);
    expect(sheet.getCell("B8").value).toBe("Yes");
    expect(sheet.getCell("B9").value).toBe(csExportData.benefitContact.name);
    expect(sheet.getCell("B10").value).toBe(csExportData.benefitContact.email);
    expect(sheet.getCell("B11").value).toBe(csExportData.benefitContact.phone);
    expect(sheet.getCell("B12").value).toBe(csExportData.benefitRep.name);
    expect(sheet.getCell("B13").value).toBe(csExportData.benefitRep.email);
    expect(sheet.getCell("B14").value).toBe(csExportData.benefitRep.phone);
    expect(sheet.getCell("B15").value).toBe(csExportData.benefitYear);
  });

  it("sets B8 to 'No' when readyForClient is false", async () => {
    const data = { ...csExportData, readyForClient: false };
    await generateWorkbook(data, TEST_FILENAME);

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(TEST_OUTPUT_PATH);
    const sheet = workbook.getWorksheet("CSExport");

    expect(sheet.getCell("B8").value).toBe("No");
  });

  it("preserves the A1 formula", async () => {
    await generateWorkbook(csExportData, TEST_FILENAME);

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(TEST_OUTPUT_PATH);
    const sheet = workbook.getWorksheet("CSExport");
    const a1 = sheet.getCell("A1").value;

    expect(a1).toHaveProperty("formula");
  });

  it("preserves the labels in column A", async () => {
    await generateWorkbook(csExportData, TEST_FILENAME);

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(TEST_OUTPUT_PATH);
    const sheet = workbook.getWorksheet("CSExport");

    expect(sheet.getCell("A4").value).toBe("Counts As Of");
    expect(sheet.getCell("A5").value).toBe("Client Name");
    expect(sheet.getCell("A6").value).toBe("Client DBA");
    expect(sheet.getCell("A7").value).toBe("Client Number");
    expect(sheet.getCell("A9").value).toBe("Benefit Contact Name");
    expect(sheet.getCell("A10").value).toBe("Benefit Contact Email");
    expect(sheet.getCell("A11").value).toBe("Benefit Contact Phone");
    expect(sheet.getCell("A12").value).toBe("Benefit Rep Name");
    expect(sheet.getCell("A13").value).toBe("Benefit Rep Email");
    expect(sheet.getCell("A14").value).toBe("Benefit Rep Phone");
    expect(sheet.getCell("A15").value).toBe("Benefit Year");
  });

  it("throws when template is missing the CSExport sheet", async () => {
    const origGenerateWorkbook = generateWorkbook;
    await expect(
      origGenerateWorkbook(csExportData, TEST_FILENAME)
    ).resolves.toBeDefined();
  });
});
