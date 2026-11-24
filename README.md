# Sydbyggarna Resursplanering

Delad planering av hantverkare och inhyrd personal per projekt.

- `resursplanering/index.html` – själva verktyget (publiceras som Claude-artifact med delad databas).
- `resursplanering/seed/` – startdata från den tidigare Excel-planeringen (v. 44–46 2026).

Vyer: **Planering** (person × dag, dra för att planera flera rutor), **Projektöversikt** (bemanning per projekt och dag, lag per projekt), **Personal & projekt** (register, färger, logotyp). PDF-export för 1–3 veckor framåt.

# Sydbyggarna Mötesprotokoll

Gemensam standard och verktyg för samordnings- och byggmöten (projekt 5–50 MSEK).

- `motesprotokoll/index.html` – verktyget (publiceras som Claude-artifact med delad databas).
- `motesprotokoll/seed/import_docx.py` – läser ett protokoll i den gamla Word-mallen och gör om det till punkter med fasta nummer och historik.
- `motesprotokoll/seed/` – startdata från 1039 Ljungbyhedsskolan, SM 27 (2026-10-21).

Vyer: **Protokoll** (punkter per avsnitt, noteringar per möte, klarmarkering, fastställ), **Närvaro & sändlista**, **Åtgärdslista** (öppna punkter per ansvarig), **Projektets flöde** (tidslinje och sök i hela historiken), **Projekt & avsnitt**, **Vår standard**. PDF av protokoll och åtgärdslista.
