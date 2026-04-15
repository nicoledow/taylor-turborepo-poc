export function fetchCSExportData(_clientId) {
  return {
    countsAsOf: new Date("2024-08-01"),
    clientName: "QUESTCO TEST",
    clientDBA: "DEMO CLIENT",
    clientNumber: "102",
    readyForClient: true,
    benefitContact: {
      name: "Test Contact",
      email: "test@questco.net",
      phone: "(936) 756-1980",
    },
    benefitRep: {
      name: "",
      email: "",
      phone: "",
    },
    benefitYear: 2025,
  };
}
