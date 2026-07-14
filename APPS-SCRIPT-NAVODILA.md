# Navodila: Google Apps Script — shranjevanje surovih odgovorov

Spletni vprašalnik zdaj poleg povprečij (B/G/Y/R + SN) pošilja še dva nova
parametra: `odgovori` (JSON s 15 postavkami) in `sn_odgovori` (JSON s 4
S/N odgovori). Da se shranita v Google Sheet, posodobi Apps Script (enkratno,
~3 minute):

## Koraki

1. Odpri svoj Google Sheet z odgovori.
2. Meni **Razširitve → Apps Script** (Extensions → Apps Script).
3. V funkciji, ki zapisuje vrstico (`appendRow([...])`), dodaj NA KONEC
   seznama dve vrstici:

   ```js
   e.parameter.odgovori || '',
   e.parameter.sn_odgovori || ''
   ```

   Če želiš, lahko celotno funkcijo zamenjaš s spodnjo (preveri, da vrstni
   red ustreza stolpcem tvojega Sheeta):

   ```js
   function doGet(e) {
     const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
     sheet.appendRow([
       new Date(),
       e.parameter.ime || '',
       e.parameter.email || '',
       e.parameter.podjetje || '',
       e.parameter.delovno_mesto || '',
       e.parameter.spol || '',
       e.parameter.B || '',
       e.parameter.G || '',
       e.parameter.Y || '',
       e.parameter.R || '',
       e.parameter.SN || '',
       e.parameter.odgovori || '',
       e.parameter.sn_odgovori || ''
     ]);
     return ContentService.createTextOutput('OK');
   }
   ```

4. **Deploy → Manage deployments → ✏️ Edit → Version: New version → Deploy.**
   POMEMBNO: uredi OBSTOJEČI deployment (ne »New deployment«), da URL ostane
   isti — sicer stran ne bo več pošiljala na pravi naslov.
5. V Sheetu po želji dodaj naslova stolpcev: `odgovori`, `sn_odgovori`.

## Uvoz v admin

V adminu → **+ Ročni vnos** → zgoraj »⚡ Uvoz iz Google Sheeta«: v Sheetu
označi celo vrstico stranke, kopiraj, prilepi — obrazec se izpolni sam
(ime, email, podjetje, delovno mesto, spol, barve, S/N in surovi odgovori).
Profil z odgovori dobi vse izračunane plošče (zanesljivost, odziv pod
pritiskom ...).

Če Apps Script še ni posodobljen, stran deluje naprej brez napak — le
odgovori se ne shranijo (neznane parametre Sheet ignorira).
