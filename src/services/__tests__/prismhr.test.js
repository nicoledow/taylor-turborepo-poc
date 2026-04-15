import { describe, it, expect } from "vitest";
import { fetchCSExportData } from "../prismhr.js";

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
