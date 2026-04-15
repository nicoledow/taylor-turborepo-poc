import { describe, it, expect, beforeAll, afterAll } from "vitest";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import ExcelJS from "exceljs";
import { generateWorkbook } from "../workbookGenerator.js";
import {
  fetchCSExportData,
  fetchClassesPlansData,
  fetchCurrentPlansData,
  fetchRenewalPlansData,
} from "../prismhr.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const OUTPUT_DIR = path.resolve(__dirname, "..", "..", "..", "generatedWorkbooks");
const TEST_FILENAME = "test_output.xlsx";
const TEST_OUTPUT_PATH = path.join(OUTPUT_DIR, TEST_FILENAME);

function buildInput(overrides = {}) {
  return {
    csExportData: fetchCSExportData("102"),
    classesPlansData: fetchClassesPlansData("102"),
    currentPlansData: fetchCurrentPlansData("102"),
    renewalPlansData: fetchRenewalPlansData("102"),
    ...overrides,
  };
}

async function readSheet(name) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(TEST_OUTPUT_PATH);
  return wb.getWorksheet(name);
}

describe("generateWorkbook", () => {
  afterAll(() => {
    if (fs.existsSync(TEST_OUTPUT_PATH)) {
      fs.unlinkSync(TEST_OUTPUT_PATH);
    }
  });

  describe("file output", () => {
    it("creates an xlsx file in the generatedWorkbooks directory", async () => {
      const outputPath = await generateWorkbook(buildInput(), TEST_FILENAME);
      expect(fs.existsSync(outputPath)).toBe(true);
      expect(outputPath).toBe(TEST_OUTPUT_PATH);
    });

    it("replaces .xlsm extension with .xlsx in output filename", async () => {
      const expectedPath = path.join(OUTPUT_DIR, "test_output.xlsx");
      const outputPath = await generateWorkbook(buildInput(), "test_output.xlsm");
      expect(outputPath).toBe(expectedPath);
      expect(outputPath.endsWith(".xlsx")).toBe(true);
    });

    it("produces a file with no VBA macros", async () => {
      await generateWorkbook(buildInput(), TEST_FILENAME);
      const { createReadStream } = await import("node:fs");
      const { default: unzipper } = await import("unzipper");
      const entries = [];
      await new Promise((resolve, reject) => {
        createReadStream(TEST_OUTPUT_PATH)
          .pipe(unzipper.Parse())
          .on("entry", (entry) => { entries.push(entry.path); entry.autodrain(); })
          .on("close", resolve)
          .on("error", reject);
      });
      const vbaEntries = entries.filter(
        (e) => e.includes("vbaProject") || e.endsWith(".bin")
      );
      expect(vbaEntries).toEqual([]);
    });
  });

  describe("CSExport sheet", () => {
    let input;
    beforeAll(async () => {
      input = buildInput();
      await generateWorkbook(input, TEST_FILENAME);
    });

    it("populates all data cells", async () => {
      const sheet = await readSheet("CSExport");
      const d = input.csExportData;
      expect(sheet.getCell("B4").value).toEqual(d.countsAsOf);
      expect(sheet.getCell("B5").value).toBe(d.clientName);
      expect(sheet.getCell("B6").value).toBe(d.clientDBA);
      expect(sheet.getCell("B7").value).toBe(d.clientNumber);
      expect(sheet.getCell("B8").value).toBe("Yes");
      expect(sheet.getCell("B9").value).toBe(d.benefitContact.name);
      expect(sheet.getCell("B10").value).toBe(d.benefitContact.email);
      expect(sheet.getCell("B11").value).toBe(d.benefitContact.phone);
      expect(sheet.getCell("B15").value).toBe(d.benefitYear);
    });

    it("sets B8 to 'No' when readyForClient is false", async () => {
      await generateWorkbook(
        buildInput({ csExportData: { ...input.csExportData, readyForClient: false } }),
        TEST_FILENAME
      );
      const sheet = await readSheet("CSExport");
      expect(sheet.getCell("B8").value).toBe("No");
    });

    it("preserves the A1 formula", async () => {
      await generateWorkbook(input, TEST_FILENAME);
      const sheet = await readSheet("CSExport");
      expect(sheet.getCell("A1").value).toHaveProperty("formula");
    });

    it("preserves labels in column A", async () => {
      await generateWorkbook(input, TEST_FILENAME);
      const sheet = await readSheet("CSExport");
      expect(sheet.getCell("A4").value).toBe("Counts As Of");
      expect(sheet.getCell("A5").value).toBe("Client Name");
      expect(sheet.getCell("A9").value).toBe("Benefit Contact Name");
      expect(sheet.getCell("A15").value).toBe("Benefit Year");
    });
  });

  describe("ClassesPlans sheet", () => {
    let input;
    beforeAll(async () => {
      input = buildInput();
      await generateWorkbook(input, TEST_FILENAME);
    });

    it("populates class names and codes", async () => {
      const sheet = await readSheet("ClassesPlans");
      const cls = input.classesPlansData.classes;
      expect(sheet.getCell("B7").value).toBe(cls[0].name);
      expect(sheet.getCell("C7").value).toBe(cls[1].name);
      expect(sheet.getCell("D7").value).toBe(cls[2].name);
      expect(sheet.getCell("B8").value).toBe(cls[0].code);
    });

    it("populates current and renewal plan counts", async () => {
      const sheet = await readSheet("ClassesPlans");
      const cls = input.classesPlansData.classes[0];
      expect(sheet.getCell("B9").value).toBe(cls.currentHealthCount);
      expect(sheet.getCell("B10").value).toBe(cls.currentDentalCount);
      expect(sheet.getCell("B11").value).toBe(cls.currentVisionCount);
      expect(sheet.getCell("B12").value).toBe(cls.renewalHealthCount);
      expect(sheet.getCell("B13").value).toBe(cls.renewalDentalCount);
    });

    it("populates EAP fields", async () => {
      const sheet = await readSheet("ClassesPlans");
      expect(sheet.getCell("A71").value).toBe("Questco EAP?");
      expect(sheet.getCell("A72").value).toBe(input.classesPlansData.eapRate);
    });
  });

  describe("CurrentPlans sheet", () => {
    let input;
    beforeAll(async () => {
      input = buildInput();
      await generateWorkbook(input, TEST_FILENAME);
    });

    it("populates health plan data starting at row 4", async () => {
      const sheet = await readSheet("CurrentPlans");
      const h = input.currentPlansData.health[0];
      expect(sheet.getCell("A4").value).toBe(h.primaryKey);
      expect(sheet.getCell("B4").value).toBe(h.planClass);
      expect(sheet.getCell("E4").value).toBe(h.carrier);
      expect(sheet.getCell("H4").value).toBe(h.eo);
      expect(sheet.getCell("W4").value).toBe(h.inCoinsurance);
    });

    it("populates dental plan data starting at row 9", async () => {
      const sheet = await readSheet("CurrentPlans");
      const d = input.currentPlansData.dental[0];
      expect(sheet.getCell("A9").value).toBe(d.primaryKey);
      expect(sheet.getCell("E9").value).toBe(d.carrier);
      expect(sheet.getCell("H9").value).toBe(d.eo);
    });

    it("populates vision plan data starting at row 14", async () => {
      const sheet = await readSheet("CurrentPlans");
      const v = input.currentPlansData.vision[0];
      expect(sheet.getCell("A14").value).toBe(v.primaryKey);
      expect(sheet.getCell("E14").value).toBe(v.carrier);
    });

    it("populates other plan data starting at row 19", async () => {
      const sheet = await readSheet("CurrentPlans");
      const o = input.currentPlansData.other[0];
      expect(sheet.getCell("A19").value).toBe(o.primaryKey);
      expect(sheet.getCell("B19").value).toBe(o.category);
    });

    it("preserves section header labels", async () => {
      const sheet = await readSheet("CurrentPlans");
      expect(sheet.getCell("A2").value).toBe("Health");
      expect(sheet.getCell("A7").value).toBe("Dental");
      expect(sheet.getCell("A12").value).toBe("Vision");
      expect(sheet.getCell("A17").value).toBe("Other");
    });
  });

  describe("RenewalPlans sheet", () => {
    let input;
    beforeAll(async () => {
      input = buildInput();
      await generateWorkbook(input, TEST_FILENAME);
    });

    it("populates renewal health data starting at row 4", async () => {
      const sheet = await readSheet("RenewalPlans");
      const h = input.renewalPlansData.health[0];
      expect(sheet.getCell("A4").value).toBe(h.renewedFromPrimaryKey);
      expect(sheet.getCell("F4").value).toBe(h.carrier);
      expect(sheet.getCell("K4").value).toBe(h.eo);
    });

    it("populates renewal dental data starting at row 10", async () => {
      const sheet = await readSheet("RenewalPlans");
      const d = input.renewalPlansData.dental[0];
      expect(sheet.getCell("A10").value).toBe(d.renewedFromPrimaryKey);
      expect(sheet.getCell("F10").value).toBe(d.carrier);
    });

    it("populates renewal vision data starting at row 16", async () => {
      const sheet = await readSheet("RenewalPlans");
      const v = input.renewalPlansData.vision[0];
      expect(sheet.getCell("A16").value).toBe(v.renewedFromPrimaryKey);
    });

    it("populates renewal STD data starting at row 22", async () => {
      const sheet = await readSheet("RenewalPlans");
      const s = input.renewalPlansData.std[0];
      expect(sheet.getCell("A22").value).toBe(s.renewedFromPrimaryKey);
      expect(sheet.getCell("F22").value).toBe(s.carrier);
      expect(sheet.getCell("Z22").value).toBe(s.benefitPercentage);
    });

    it("populates renewal LTD data starting at row 28", async () => {
      const sheet = await readSheet("RenewalPlans");
      const l = input.renewalPlansData.ltd[0];
      expect(sheet.getCell("A28").value).toBe(l.renewedFromPrimaryKey);
      expect(sheet.getCell("Z28").value).toBe(l.benefitPercentage);
    });

    it("populates renewal Life data starting at row 34", async () => {
      const sheet = await readSheet("RenewalPlans");
      const life = input.renewalPlansData.life[0];
      expect(sheet.getCell("A34").value).toBe(life.renewedFromPrimaryKey);
      expect(sheet.getCell("Z34").value).toBe(life.employeeLifeBenefitAmount);
    });

    it("preserves section header labels", async () => {
      const sheet = await readSheet("RenewalPlans");
      expect(sheet.getCell("A2").value).toBe("Health");
      expect(sheet.getCell("A8").value).toBe("Dental");
      expect(sheet.getCell("A14").value).toBe("Vision");
      expect(sheet.getCell("A20").value).toBe("STD");
      expect(sheet.getCell("A26").value).toBe("LTD");
      expect(sheet.getCell("A32").value).toBe("Life");
      expect(sheet.getCell("A38").value).toBe("Other");
    });
  });

  describe("Welcome and Total Cost sheets", () => {
    beforeAll(async () => {
      await generateWorkbook(buildInput(), TEST_FILENAME);
    });

    it("preserves the Welcome sheet", async () => {
      const sheet = await readSheet("Welcome");
      expect(sheet).toBeDefined();
    });

    it("preserves Total Cost formulas", async () => {
      const sheet = await readSheet("Total Cost");
      expect(sheet).toBeDefined();
      const a5 = sheet.getCell("A5").value;
      expect(a5).toHaveProperty("formula");
    });
  });
});
