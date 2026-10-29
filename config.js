// Konfiguracija aplikacije Registri HCS.
// Upišite vrijednosti jednom; kasnije zamjene datoteke index.html ne diraju ovu datoteku.
window.CONFIG = {
  tenantId:    "08ee12ee-23ed-461f-a6bd-d22bbd18980f",
  clientId:    "f90e3e6b-411d-4dc6-9d66-8b1324c8a65f",
  redirectUri: window.location.origin + window.location.pathname,
  spHost:      "curlinghr.sharepoint.com",
  sitePath:    "",                       // prazno = korijenski site (intranet)
  lists: {
    igraci:  "Registar igrača",
    klubovi: "Registar klubova"
  }
};
