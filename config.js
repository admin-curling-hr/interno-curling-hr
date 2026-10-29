// Konfiguracija internog dijela web mjesta Saveza (interno.curling.hr).
// Upišite vrijednosti jednom; zamjene ostalih datoteka ne diraju ovu datoteku.
// Adresa za povratak nakon prijave računa se automatski (korijen web mjesta).
window.CONFIG = {
  tenantId: "UPISATI-DIRECTORY-TENANT-ID",
  clientId: "UPISATI-APPLICATION-CLIENT-ID",
  spHost:   "curlinghr.sharepoint.com",
  sitePath: "",                         // prazno = korijenski site (intranet)
  lists: {
    igraci:  "Registar igrača",
    klubovi: "Registar klubova",
    ekipe:   "Registar ekipa"
  }
};
