# interno.curling.hr

Interni dio web mjesta Hrvatskog curling saveza, samo za ovlaštene osobe.
Prijava: Microsoft Entra ID (tenant Saveza). Podaci: SharePoint liste na intranetu, preko Microsoft Graph.

- `config.js` – tenantId, clientId, SharePoint site i nazivi lista (ne sadrži tajne)
- `assets/shell.js`, `assets/shell.css` – zajednički header, izbornik, prijava, Graph
- `registri/` – registri igrača, klubova i ekipa

Novi modul: nova mapa sa stranicom koja učitava `../config.js`, MSAL, `../assets/shell.js`
i čeka `Shell.ready`; stavka izbornika dodaje se u popis `MODULES` u `assets/shell.js`.
