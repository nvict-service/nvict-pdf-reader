# NVict Reader - Nederlandse Taalverbeteringen

## Overzicht
Dit rapport bevat aanbevelingen voor Nederlandse taalverbeteringen in NVict_Reader.py. De bestaande taal is over het algemeen van hoge kwaliteit, maar er zijn enkele kleine verbeteringen mogelijk voor consistentie en helderheid.

---

## 1. MENU ITEMS (Prioriteit: Laag → Medium)

### 1.1 "Kopieer tekst" → "Tekst kopiëren"
- **Huide plaats**: Regel 873
- **Hudig**: `edit_menu.add_command(label="Kopieer tekst", ...)`
- **Aanbeveling**: `edit_menu.add_command(label="Tekst kopiëren", ...)`
- **Reden**: Consistentie met andere menu-items die werkwoorden gebruiken ("Zoeken...", "Pagina's exporteren..."). "Tekst kopiëren" volgt het patroon "zelfstandignaamwoord + werkwoord" beter.
- **Impact**: Laag - cosmetisch

### 1.2 "Pagina's exporteren..." → consistentie met layout
- **Hudig plaats**: Regel 876
- **Hudig**: `edit_menu.add_command(label="Pagina's exporteren...", ...)`
- **Status**: ✅ Reeds goed
- **Opmerking**: Dit is consistent en duidelijk

---

## 2. TOOLBAR TOOLTIPS (Prioriteit: Medium)

### 2.1 Tooltip voor "save" knop - VERWIJDERD
- **Hudig plaats**: Regel 1003 in _TOOLTIPS
- **Hudig**: `"save": "Opslaan (Ctrl+S)",`
- **Status**: ✅ Deze knop/tooltip is al verwijderd uit v2.1
- **Opmerking**: Correct - het formulieropslagfunctie is volledig verwijderd

### 2.2 "Zoeken in document" is duidelijk
- **Hudig plaats**: Regel 1005
- **Hudig**: `"search": "Zoeken in document  (Ctrl+F)",`
- **Status**: ✅ Goed
- **Reden**: Zeer duidelijk en specifiek

---

## 3. DIALOOG LABELS EN KNOPPEN (Prioriteit: Medium)

### 3.1 "Standaard App Instellen" dialoog
- **Plaats**: Regel 273
- **Hudig**: `messagebox.askyesno("Standaard App Instellen", ...)`
- **Aanbeveling**: Kan blijven, maar alternatief: `"Als standaard instellen"`
- **Reden**: Beide zijn acceptabel; huiding is duidelijk
- **Impact**: Laag - kan blijven

### 3.2 "Dubbelzijdig Printen" dialoog
- **Plaats**: Regel 2840
- **Hudig**: `messagebox.askyesno("Dubbelzijdig Printen", ...)`
- **Status**: ✅ Goed
- **Opmerking**: Duidelijk en consistent met nieuwe v2.1 functie

### 3.3 Print dialoog knoppen
- **Plaats**: Regel 2872+ (buttons in print dialog)
- **Hudig**: Knop labels zijn in Nederlands (bevestigd)
- **Status**: ✅ Goed

---

## 4. PRINT DIALOOG LABELS (Prioriteit: Medium)

### 4.1 "Document:" label
- **Plaats**: Regel 2545
- **Hudig**: `tk.Label(content_frame, text="Document:", ...)`
- **Aanbeveling**: Kan "Document:" blijven, of "Bestand:" gebruiken
- **Reden**: Beide zijn acceptabel; "Document" is meer formeel
- **Impact**: Laag

### 4.2 "Printers laden..." tekst
- **Plaats**: Regel 2555
- **Hudig**: `printer_var = tk.StringVar(value="Printers laden...")`
- **Status**: ✅ Goed
- **Opmerking**: Duidelijk feedback aan gebruiker

### 4.3 Print dialoog secties - alle goed benoemd
- **Pagina's:** ✅ Correct
- **Printer:** ✅ Correct
- **Aantal kopieën:** ✅ Correct en duidelijk
- **Dubbelzijdig printen:** ✅ Correct (nieuw v2.1)
- **Oriëntatie:** ✅ Correct met "Staand/Liggend" (v2.1)
- **Kwaliteit:** ✅ Correct met "Snel/Normaal/Hoog" (v2.1)

### 4.4 "Alle pagina's (1-X)" vs "Alle pagina's"
- **Plaats**: Regel 2607
- **Hudig**: `text=f"Alle pagina's (1-{len(tab.pdf_document)})"`
- **Status**: ✅ Goed - duidelijk bereik
- **Opmerking**: Helpt gebruiker begrijpen welke pagina's

### 4.5 "Huidge pagina" label
- **Plaats**: Regel 2625
- **Hudig**: `text=f"(pagina {tab.current_page + 1})"`
- **Status**: ✅ Goed

---

## 5. FOUTMELDINGEN (Prioriteit: Laag)

### 5.1 Foutmelding terminologie - consistent
- **Plaats**: Verschillende locaties (2813, 2824, 2840, 2857, etc.)
- **Patronen**:
  - `messagebox.showerror("Fout", ...)`
  - `messagebox.showerror("Ongeldige pagina's", ...)`
  - `messagebox.showerror("Invoer Fout", ...)`
  - `messagebox.showerror("Print Fout", ...)`
- **Opmerking**: Foutmeldingen zijn over het algemeen goed benoemd
- **Suggestie**: Inconsistentie in "Fout" vs "Print Fout" - kan uniformer:
  - Optie 1: Alle gebruiken specifieke titel ("Print Fout", "Invoer Fout", etc.)
  - Optie 2: Alle generiek houden als "Fout"
  - **Huiding**: Specifieke titels zijn beter voor UX

### 5.2 "Geen PDF" bericht
- **Plaats**: Regel 1151
- **Hudig**: `messagebox.showinfo("Geen PDF", "Open eerst een PDF...")`
- **Aanbeveling**: Kan ook "Geen document geopend" zijn
- **Status**: ✅ Acceptabel

---

## 6. STATUSBAR EN NAVIGATIE (Prioriteit: Laag)

### 6.1 Pagina indicator teksten
- **Voorbeelden**:
  - Regel 1255: `f"Pagina {self.fs_page + 1} / {total}  ·  Escape om te sluiten"`
  - Regel 1527: `f"Zoom: {int(tab.zoom_level * 100)}%"`
  - Regel 3187-3189: "Pagina..." en "Pagina's..."
- **Status**: ✅ Alle correct

---

## 7. DIALOG TITELS (Prioriteit: Laag → Medium)

### 7.1 Exporteren dialog
- **Plaats**: Regel 3397
- **Hudig**: `dialog.title("Pagina's Exporteren")`
- **Status**: ✅ Goed

### 7.2 Roteren dialog
- **Plaats**: Regel 3538
- **Hudig**: `rotate_dialog.title("Pagina's Roteren")`
- **Status**: ✅ Goed

### 7.3 Extraheren dialog
- **Plaats**: Regel 3673
- **Hudig**: `dialog.title("Pagina's extraheren")`
- **Aanbeveling**: "Pagina's Extraheren" (hoofdletter E) voor consistentie
- **Impact**: Laag - esthetisch

---

## 8. WELKOMST BERICHT (Prioriteit: Laag)

### 8.1 Welkomst tekst completeness
- **Plaats**: Regel 299-301
- **Hudig**: Duidelijk Nederlands
- **Status**: ✅ Goed
- **Opmerking**: Toon is vriendelijk en professioneel

---

## 9. MENU ONDERDELEN - SAMENVATTENDE EVALUATIE

| Menu Item | Hudig | Status | Opmerking |
|-----------|-------|--------|-----------|
| Bestand > Openen... | Openen... | ✅ Goed | Duidelijk |
| Bestand > Recente bestanden | Recente bestanden | ✅ Goed | Consistent |
| Bestand > Afdrukken... | Afdrukken... | ✅ Goed | Duidelijk |
| Bestand > Sluiten | Sluiten | ✅ Goed | Consistent |
| Bestand > Afsluiten | Afsluiten | ✅ Goed | Correct |
| Bewerken > Kopieer tekst | Kopieer tekst | ⚠️ | Consider "Tekst kopiëren" |
| Bewerken > Zoeken... | Zoeken... | ✅ Goed | Tooltip: "Zoeken in document" |
| Bewerken > Pagina's exporteren... | Pagina's exporteren... | ✅ Goed | Duidelijk |
| Bewerken > PDF's samenvoegen... | PDF's samenvoegen... | ✅ Goed | Duidelijk |
| Bewerken > Pagina's roteren... | Pagina's roteren... | ✅ Goed | Duidelijk |
| Beeld > Zoom in | Zoom in | ✅ Goed | Consistent |
| Beeld > Zoom uit | Zoom uit | ✅ Goed | Consistent |
| Beeld > Pasbreedte | Pasbreedte | ✅ Goed | Correct Nederlands |
| Beeld > Eerste pagina | Eerste pagina | ✅ Goed | Duidelijk |
| Beeld > Vorige pagina | Vorige pagina | ✅ Goed | Consistent |
| Beeld > Volgende pagina | Volgende pagina | ✅ Goed | Consistent |
| Beeld > Laatste pagina | Laatste pagina | ✅ Goed | Duidelijk |
| Beeld > Volledig scherm | Volledig scherm | ✅ Goed | Correct |
| Beeld > Pagina's paneel | Pagina's paneel | ✅ Goed | Consistent met tooltip |
| Beeld > Boek-modus | Boek-modus | ✅ Goed | Duidelijk Nederlands |
| Beeld > Markeermodus | Markeermodus | ✅ Goed | Consistent |
| Instellingen > Als standaard... | Instellen als standaard PDF viewer | ✅ Goed | Duidelijk |
| Help > PDF Info | PDF Info | ✅ Goed | Acceptabel |
| Help > Controleer op updates... | Controleer op updates... | ✅ Goed | Duidelijk |
| Help > Over NVict Reader | Over NVict Reader | ✅ Goed | Correct |

---

## 10. SAMENVATTING VAN AANBEVELINGEN

### Kritiek (Implementatie aangeraden) ⚠️ **GEEN**
- Alle kritieke items zijn al opgelost in v2.1

### Belangrijk (Overwegen te implementeren) ⚠️
1. **"Kopieer tekst" → "Tekst kopiëren"** (Regel 873)
   - Voor consistentie met ander menu-items
   - Impact: Laag, maar verbetert UX

### Optioneel (Fijn tuning)
1. "Pagina's extraheren" → "Pagina's Extraheren" (Regel 3673) - hoofdletter E
2. "Document:" → "Bestand:" in print dialoog (optioneel)

---

## 11. KWALITEIT EVALUATIE

**Algemeen Nederlands nivo**: ⭐⭐⭐⭐⭐ (5/5)
- Alle UI-text is correct Nederlands
- Terminologie is consistent en professioneel
- Geen onduidelijke of vertaalde termen
- v2.1 features (Dubbelzijdig, Staand/Liggend, Kwaliteit) zijn allemaal in goed Nederlands

**Specifieke sterktes**:
- ✅ Duidelijk Nederlands in foutmeldingen
- ✅ Professionele dialoogtitels
- ✅ Consistente menu-organisatie
- ✅ Goed gekozen Nederlandse termen ("Markeermodus", "Boek-modus", "Pagina's paneel")
- ✅ Tooltips zijn informatief en compleet

**Aanbevelingen voor toekomstige releases**:
- Overweeg "Kopieer tekst" → "Tekst kopiëren" voor v2.2
- Blijf huiding volgen voor nieuwe features
- Behou the huiding Staand/Liggend in plaats van Portret/Landschap

---

## CONCLUSIE

NVict Reader heeft een **zeer hoog niveau van Nederlandse lokalisatie**. Het programma is goed vertaald en maakt gebruik van professionele Nederlandse termen. De meeste interface-elementen zijn duidelijk en begrijpelijk.

**Aanbevolen actie**: Geen onmiddellijke wijzigingen nodig. De applicatie is gereed voor Nederlands-sprekende gebruikers. Voor toekomstige versies kan overwogen worden "Kopieer tekst" naar "Tekst kopiëren" te wijzigen voor betere consistentie.

---

**Rapport gegenereerd**: 2026-03-24
**Applicatie versie**: 2.1
**Onderzochte bestanden**: NVict_Reader.py (4500+ regels)
