import ExcelJS from "exceljs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const PROJECT_ROOT = path.resolve(__dirname, "..", "..");
const TEMPLATE_PATH = path.join(PROJECT_ROOT, "templates", "workbook_template.xlsm");
const OUTPUT_DIR = path.join(PROJECT_ROOT, "generatedWorkbooks");

export async function generateWorkbook(
  { csExportData, classesPlansData, currentPlansData, renewalPlansData },
  outputFilename
) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(TEMPLATE_PATH);

  const csSheet = workbook.getWorksheet("CSExport");
  if (!csSheet) throw new Error('Sheet "CSExport" not found in template');
  populateCSExport(csSheet, csExportData);

  const cpSheet = workbook.getWorksheet("ClassesPlans");
  if (!cpSheet) throw new Error('Sheet "ClassesPlans" not found in template');
  populateClassesPlans(cpSheet, classesPlansData);

  const curSheet = workbook.getWorksheet("CurrentPlans");
  if (!curSheet) throw new Error('Sheet "CurrentPlans" not found in template');
  populateCurrentPlans(curSheet, currentPlansData);

  const renSheet = workbook.getWorksheet("RenewalPlans");
  if (!renSheet) throw new Error('Sheet "RenewalPlans" not found in template');
  populateRenewalPlans(renSheet, renewalPlansData);

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

const CLASS_COLUMNS = ["B", "C", "D", "E", "F", "G", "H"];

function populateClassesPlans(sheet, data) {
  data.classes.forEach((cls, i) => {
    if (i >= 7) return;
    const col = CLASS_COLUMNS[i];
    sheet.getCell(`${col}7`).value = cls.name;
    sheet.getCell(`${col}8`).value = cls.code;
    sheet.getCell(`${col}9`).value = cls.currentHealthCount;
    sheet.getCell(`${col}10`).value = cls.currentDentalCount;
    sheet.getCell(`${col}11`).value = cls.currentVisionCount;
    sheet.getCell(`${col}12`).value = cls.renewalHealthCount;
    sheet.getCell(`${col}13`).value = cls.renewalDentalCount;
    sheet.getCell(`${col}14`).value = cls.renewalVisionCount;
    sheet.getCell(`${col}15`).value = cls.renewalSTDCount;
    sheet.getCell(`${col}16`).value = cls.renewalLTDCount;
    sheet.getCell(`${col}17`).value = cls.renewalLifeCount;
  });

  sheet.getCell("A71").value = data.questcoEAP ? "Questco EAP?" : "";
  sheet.getCell("A72").value = data.eapRate;
}

const CURRENT_HEALTH_COLS = [
  "primaryKey", "planClass", "planID", "uniquePlanCode", "carrier",
  "riskTier", "benefitGroup", "eo", "es", "ec", "ek", "fam",
  "eoCount", "esCount", "ekCount", "ecCount", "famCount",
  "eoContrib", "esContrib", "ecContrib", "ekContrib", "famContrib",
  "inCoinsurance", "inDedIndiv", "inDedFam", "inOopMaxIndiv", "inOopMaxFam",
  "officeVisitPCP", "officeVisitSpecialist", "advancedRadiology",
  "hospitalInpatient", "surgeryOutpatient", "emergencyRoom", "urgentCare",
  "pharmacyDedAmt", "pharmacyGeneric", "pharmacyFormulary",
  "pharmacyNonFormulary", "pharmacySpecialty", "pharmacyMailOrder",
  "oonCoinsurance", "oonDedIndiv", "oonDedFam", "oonOopMaxIndiv", "oonOopMaxFam",
  "healthLabs", "healthDiagnosticXray",
];

const CURRENT_DENTAL_COLS = [
  "primaryKey", "planClass", "planID", "uniquePlanCode", "carrier",
  "riskTier", "benefitGroup", "eo", "es", "ec", "ek", "fam",
  "eoCount", "esCount", "ekCount", "ecCount", "famCount",
  "eoContrib", "esContrib", "ecContrib", "ekContrib", "famContrib",
  "inOrthodontiaPctDollar", "inOrthodontiaMax",
  "oonDedIndiv", "oonDedFam", "oonPreventativeCare",
  "oonBasicRestorative", "oonMajorRestorative", "annualDentalMax",
];

const CURRENT_VISION_COLS = [
  "primaryKey", "planClass", "planID", "uniquePlanCode", "carrier",
  "riskTier", "benefitGroup", "eo", "es", "ec", "ek", "fam",
  "eoCount", "esCount", "ekCount", "ecCount", "famCount",
  "eoContrib", "esContrib", "ecContrib", "ekContrib", "famContrib",
  "inFreqFrames", "inSingleVisionLenses", "inBifocalLenses",
  "inTrifocalLenses", "inFrames", "inContactLensesNoFrames",
  "oonCopayExam", "oonCopayContactExam", "oonCopayMaterials",
  "oonSingleVisionLenses", "oonBifocalLenses", "oonTrifocalLenses",
  "oonFrames", "oonContactLensesNoFrames",
];

const CURRENT_OTHER_COLS = [
  "primaryKey", "category", "planClass", "planID", "uniquePlanCode", "carrier",
  "riskTier", "benefitGroup", "eo", "es", "ek", "ec", "fam",
  "eoCount", "ekCount", "esCount", "ecCount", "famCount",
  "eoContrib", "esContrib", "ecContrib", "ekContrib", "famContrib",
];

function writeRows(sheet, startRow, columns, rows) {
  rows.forEach((row, rowIdx) => {
    columns.forEach((key, colIdx) => {
      const val = row[key];
      if (val !== undefined && val !== null) {
        sheet.getCell(startRow + rowIdx, colIdx + 1).value = val;
      }
    });
  });
}

function populateCurrentPlans(sheet, data) {
  writeRows(sheet, 4, CURRENT_HEALTH_COLS, data.health);
  writeRows(sheet, 9, CURRENT_DENTAL_COLS, data.dental);
  writeRows(sheet, 14, CURRENT_VISION_COLS, data.vision);
  writeRows(sheet, 19, CURRENT_OTHER_COLS, data.other);
}

const RENEWAL_COMMON_COLS = [
  "renewedFromPrimaryKey", "planClass", "planID", "renewedFromPlan",
  "uniquePlanCode", "carrier", "riskTier", "benefitGroup",
  "effectiveDate", "contributionMethod",
  "eo", "es", "ec", "ek", "fam",
  "eoCount", "esCount", "ecCount", "ekCount", "famCount",
  "eoContrib", "esContrib", "ecContrib", "famContrib", "ekContrib",
];

const RENEWAL_HEALTH_EXTRA = [
  "inCoinsurance", "inDedIndiv", "inDedFam", "inOopMaxIndiv", "inOopMaxFam",
  "officeVisitPCP", "officeVisitSpecialist", "advancedRadiology",
  "hospitalInpatient", "surgeryOutpatient", "emergencyRoom", "urgentCare",
  "pharmacyDedAmt", "pharmacyGeneric", "pharmacyFormulary",
  "pharmacyNonFormulary", "pharmacySpecialty", "pharmacyMailOrder",
  "oonCoinsurance", "oonDedIndiv", "oonDedFam", "oonOopMaxIndiv", "oonOopMaxFam",
  "labs", "diagnosticXray",
];

const RENEWAL_DENTAL_EXTRA = [
  "inDedIndiv", "inDedFam", "inPreventativeCare", "inBasicRestorative",
  "inMajorRestorative", "inOrthodontiaPctDollar", "inOrthodontiaMax",
  "oonDedIndiv", "oonDedFam", "oonPreventativeCare",
  "oonBasicRestorative", "oonMajorRestorative", "annualDentalMax",
];

const RENEWAL_VISION_EXTRA = [
  "inCopayExam", "inCopayContactExam", "inCopayMaterials",
  "inFreqEyeExam", "inFreqReplacementLenses", "inFreqFrames",
  "inSingleVisionLenses", "inBifocalLenses", "inTrifocalLenses",
  "inFrames", "inContactLensesNoFrames",
  "oonCopayExam", "oonCopayContactExam", "oonCopayMaterials",
  "oonSingleVisionLenses", "oonBifocalLenses", "oonTrifocalLenses",
  "oonFrames", "oonContactLensesNoFrames",
];

const RENEWAL_STD_EXTRA = [
  "benefitPercentage", "maximumWeeklyBenefit", "benefitDuration",
  "illness", "accident", "definitionOfDisability",
  "preExistingLimitation", "teleguard", "returnToWork",
  "rehabServices", "coverageType",
];

const RENEWAL_LTD_EXTRA = [
  "benefitPercentage", "monthlyBenefitMaximum", "monthlyBenefitMinimum",
  "eliminationPeriod", "socialSecurityIntegration", "survivorBenefit",
  "benefitDuration", "preExisting", "mentalNervousSubstance",
  "disabilityDefinition",
];

const RENEWAL_LIFE_EXTRA = [
  "employeeLifeBenefitAmount", "guaranteeIssueMaxBenefit", "addBenefits",
  "acceleratedDeathBenefit", "waiverOfPremium", "portability",
  "conversion", "ageReductions", "commonCarrier",
];

const RENEWAL_OTHER_EXTRA = ["category"];

function populateRenewalPlans(sheet, data) {
  writeRows(sheet, 4, [...RENEWAL_COMMON_COLS, ...RENEWAL_HEALTH_EXTRA], data.health);
  writeRows(sheet, 10, [...RENEWAL_COMMON_COLS, ...RENEWAL_DENTAL_EXTRA], data.dental);
  writeRows(sheet, 16, [...RENEWAL_COMMON_COLS, ...RENEWAL_VISION_EXTRA], data.vision);
  writeRows(sheet, 22, [...RENEWAL_COMMON_COLS, ...RENEWAL_STD_EXTRA], data.std);
  writeRows(sheet, 28, [...RENEWAL_COMMON_COLS, ...RENEWAL_LTD_EXTRA], data.ltd);
  writeRows(sheet, 34, [...RENEWAL_COMMON_COLS, ...RENEWAL_LIFE_EXTRA], data.life);
  writeRows(sheet, 40, [...RENEWAL_COMMON_COLS, ...RENEWAL_OTHER_EXTRA], data.other);
}
