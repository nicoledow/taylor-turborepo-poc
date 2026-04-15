import { describe, it, expect } from "vitest";
import {
  fetchCSExportData,
  fetchClassesPlansData,
  fetchCurrentPlansData,
  fetchRenewalPlansData,
} from "../prismhr.js";

describe("fetchCSExportData", () => {
  it("returns an object with the expected shape", () => {
    const data = fetchCSExportData("102");

    expect(data).toHaveProperty("countsAsOf");
    expect(data.countsAsOf).toBeInstanceOf(Date);
    expect(data).toHaveProperty("clientName");
    expect(data).toHaveProperty("clientDBA");
    expect(data).toHaveProperty("clientNumber");
    expect(data).toHaveProperty("readyForClient");
    expect(data).toHaveProperty("benefitContact");
    expect(data).toHaveProperty("benefitRep");
    expect(data).toHaveProperty("benefitYear");
  });

  it("returns benefit contact with name, email, and phone", () => {
    const data = fetchCSExportData("102");

    expect(data.benefitContact).toEqual({
      name: expect.any(String),
      email: expect.any(String),
      phone: expect.any(String),
    });
  });

  it("returns benefit rep with name, email, and phone", () => {
    const data = fetchCSExportData("102");

    expect(data.benefitRep).toEqual({
      name: expect.any(String),
      email: expect.any(String),
      phone: expect.any(String),
    });
  });

  it("returns a numeric benefit year", () => {
    const data = fetchCSExportData("102");
    expect(typeof data.benefitYear).toBe("number");
  });

  it("returns readyForClient as a boolean", () => {
    const data = fetchCSExportData("102");
    expect(typeof data.readyForClient).toBe("boolean");
  });
});

describe("fetchClassesPlansData", () => {
  it("returns classes array with at least one entry", () => {
    const data = fetchClassesPlansData("102");
    expect(Array.isArray(data.classes)).toBe(true);
    expect(data.classes.length).toBeGreaterThan(0);
  });

  it("each class has the expected count fields", () => {
    const data = fetchClassesPlansData("102");
    const cls = data.classes[0];
    const requiredKeys = [
      "name", "code",
      "currentHealthCount", "currentDentalCount", "currentVisionCount",
      "renewalHealthCount", "renewalDentalCount", "renewalVisionCount",
      "renewalSTDCount", "renewalLTDCount", "renewalLifeCount",
    ];
    for (const key of requiredKeys) {
      expect(cls).toHaveProperty(key);
    }
  });

  it("returns questcoEAP flag and eapRate", () => {
    const data = fetchClassesPlansData("102");
    expect(typeof data.questcoEAP).toBe("boolean");
    expect(typeof data.eapRate).toBe("number");
  });
});

describe("fetchCurrentPlansData", () => {
  it("returns health, dental, vision, and other arrays", () => {
    const data = fetchCurrentPlansData("102");
    expect(Array.isArray(data.health)).toBe(true);
    expect(Array.isArray(data.dental)).toBe(true);
    expect(Array.isArray(data.vision)).toBe(true);
    expect(Array.isArray(data.other)).toBe(true);
  });

  it("health plans have rate and benefit fields as numbers", () => {
    const plan = fetchCurrentPlansData("102").health[0];
    expect(typeof plan.eo).toBe("number");
    expect(typeof plan.inCoinsurance).toBe("number");
    expect(typeof plan.inDedIndiv).toBe("number");
    expect(typeof plan.officeVisitPCP).toBe("number");
  });

  it("dental plans have rate fields as numbers", () => {
    const plan = fetchCurrentPlansData("102").dental[0];
    expect(typeof plan.eo).toBe("number");
    expect(typeof plan.annualDentalMax).toBe("number");
  });

  it("vision plans have rate fields", () => {
    const plan = fetchCurrentPlansData("102").vision[0];
    expect(typeof plan.eo).toBe("number");
    expect(typeof plan.inFrames).toBe("number");
  });
});

describe("fetchRenewalPlansData", () => {
  it("returns health, dental, vision, std, ltd, life, and other arrays", () => {
    const data = fetchRenewalPlansData("102");
    for (const key of ["health", "dental", "vision", "std", "ltd", "life", "other"]) {
      expect(Array.isArray(data[key])).toBe(true);
    }
  });

  it("renewal health plans have effectiveDate and rate fields", () => {
    const plan = fetchRenewalPlansData("102").health[0];
    expect(plan.effectiveDate).toBeInstanceOf(Date);
    expect(typeof plan.eo).toBe("number");
    expect(typeof plan.inCoinsurance).toBe("number");
  });

  it("STD plans have benefit description fields", () => {
    const plan = fetchRenewalPlansData("102").std[0];
    expect(typeof plan.benefitPercentage).toBe("number");
    expect(typeof plan.maximumWeeklyBenefit).toBe("number");
    expect(typeof plan.benefitDuration).toBe("string");
  });

  it("LTD plans have benefit description fields", () => {
    const plan = fetchRenewalPlansData("102").ltd[0];
    expect(typeof plan.benefitPercentage).toBe("number");
    expect(typeof plan.monthlyBenefitMaximum).toBe("number");
    expect(typeof plan.eliminationPeriod).toBe("string");
  });

  it("Life plans have benefit description fields", () => {
    const plan = fetchRenewalPlansData("102").life[0];
    expect(typeof plan.employeeLifeBenefitAmount).toBe("number");
    expect(typeof plan.addBenefits).toBe("string");
  });
});
