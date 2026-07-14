# Navodila: Google Apps Script — shranjevanje surovih odgovorov

Spletni vprašalnik zdaj poleg povprečij (B/G/Y/R + SN) pošilja še dva nova
parametra: `odgovori` (JSON s 15 postavkami) in `sn_odgovori` (JSON s 4
S/N odgovori). Da se shranita v Google Sheet, posodobi Apps Script (enkratno,
~3 minute):

## Stanje: POSODOBLJENO (jul. 2026)

Skripta v produkciji že vsebuje spodnjo kodo — ta datoteka je referenca,
če bo treba skripto kdaj obnoviti. Sheet: zavihek »Oddaje«, stolpca
M (Odgovori) in N (SN odgovori).

## Celotna skripta (referenca)

```js
const SHEET_ID = '1j_Pkl4q3t_YwmOxj7fA4ZucBrj3NZnqh2F2TyahoISo'
const SHEET_NAME = 'Oddaje'

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents)
    return zapisiVSheet(data)
  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({success: false, error: err.message}))
      .setMimeType(ContentService.MimeType.JSON)
  }
}

function doGet(e) {
  try {
    const data = e.parameter
    if(data.ime) return zapisiVSheet(data)
    return ContentService
      .createTextOutput(JSON.stringify({status: 'ok'}))
      .setMimeType(ContentService.MimeType.JSON)
  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({success: false, error: err.message}))
      .setMimeType(ContentService.MimeType.JSON)
  }
}

function zapisiVSheet(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID)
  let sheet = ss.getSheetByName(SHEET_NAME)

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME)
  }

  const headerRow = sheet.getRange(1, 1, 1, 14).getValues()[0]
  if (!headerRow[0]) {
    sheet.getRange(1, 1, 1, 14).setValues([[
      'Ime', 'Email', 'Spol', 'B', 'G', 'Y', 'R', 'SN', 'Podjetje', 'Delovno mesto', 'Datum', 'Cas', 'Odgovori', 'SN odgovori'
    ]])
    sheet.getRange(1, 1, 1, 14).setFontWeight('bold')
    sheet.setColumnWidth(1, 150)
    sheet.setColumnWidth(2, 200)
    sheet.setColumnWidth(9, 150)
    sheet.setColumnWidth(10, 180)
  }

  // Obstoječi list: dodaj naslova novih stolpcev, če še manjkata
  if (!sheet.getRange(1, 13).getValue()) {
    sheet.getRange(1, 13, 1, 2).setValues([['Odgovori', 'SN odgovori']])
    sheet.getRange(1, 13, 1, 2).setFontWeight('bold')
  }

  const now = new Date()

  sheet.appendRow([
    data.ime || '',
    data.email || '',
    data.spol === 'z' ? 'Ženski' : 'Moški',
    parseFloat(data.B) || 0,
    parseFloat(data.G) || 0,
    parseFloat(data.Y) || 0,
    parseFloat(data.R) || 0,
    parseFloat(data.SN) || 0,
    data.podjetje || '',
    data.delovno_mesto || '',
    Utilities.formatDate(now, 'Europe/Ljubljana', 'dd.MM.yyyy'),
    Utilities.formatDate(now, 'Europe/Ljubljana', 'HH:mm'),
    data.odgovori || '',
    data.sn_odgovori || ''
  ])

  return ContentService
    .createTextOutput(JSON.stringify({success: true}))
    .setMimeType(ContentService.MimeType.JSON)
}
```

## Deploy nove verzije (ob spremembah)

**Deploy → Manage deployments → ✏️ Edit → Version: New version → Deploy.**
POMEMBNO: uredi OBSTOJEČI deployment (ne »New deployment«), da URL ostane isti.

## Uvoz v admin

V adminu → **+ Ročni vnos** → zgoraj »⚡ Uvoz iz Google Sheeta«: v Sheetu
označi celo vrstico stranke, kopiraj, prilepi — obrazec se izpolni sam
(ime, email, podjetje, delovno mesto, spol, barve, S/N in surovi odgovori).
Profil z odgovori dobi vse izračunane plošče (zanesljivost, odziv pod
pritiskom ...).
