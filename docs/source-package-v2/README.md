# PAP, RADs, HMVs
## Historical pathways and the current landscape

**Version 2 · September 19, 2026**  
**Audience:** pulmonary and critical care fellows  
**Format:** 26 main slides with 40 minutes of nominal content, plus 5 minutes of discussion  
**Appendix:** 12 hidden slides; 38 slides in the complete PowerPoint and PDF

## Files

| File | Contents |
|---|---|
| `PAP_RAD_HMV_Pathways_v2.pptx` | Editable PowerPoint. Main slides appear in pathway order; the appendix is hidden from the default slideshow. Every slide has factual notes and references. |
| `PAP_RAD_HMV_Pathways_v2.pdf` | Rendered preview of all 38 slides, including the appendix. |
| `PAP_RAD_HMV_Companion_v2.md` | Approximately 17,700 words of slide-matched factual context, clinical evidence, historical details, policy criteria, source qualifications, and references. |
| `build/content.py` | Source of the slide order, display facts, detailed context, reference links, and duration allocations. |
| `build/slides.json` | Generated slide manifest with stable slide IDs. |
| `build/references.json` | Reference database with consultation scope. |
| `build/make_deck.py` | PowerPoint renderer. Diagrams and tables are native editable objects. |
| `build/render_qa.py` | Rendered slide images and contact sheets from the all-slide PDF. |
| `build/validate.py` | Checks slide counts, hidden flags, duration allocations, notes, anchors, physical boundaries, and rendered text. |
| `build/requirements.txt` | Python build dependencies. |
| `assets/iron_lung_hind2017.jpeg` | Attributed historical illustration from the supplied Hind review; reuse qualification below. |
| `qa/validation.json` | Structural and rendered-text validation results. |

## Main sequence

| Section | Slides | Content |
|---|---:|---|
| Title | 1 | Lecture title and the four-part sequence |
| I. Home mechanical ventilation | 2–4 | Polio-related respiratory paralysis; long-term care outside hospital; portability, continuity, caregiving, financing, and the servicing-based equipment category |
| II. Sleep apnea and PAP | 5–7 | Obstructive sleep apnea as the original clinical context; nasal CPAP in 1981; home treatment; Medicare's sleep-apnea-specific coverage pathway |
| III. Respiratory assist devices | 8–11 | Bilevel pressures and backup breaths; expansion beyond obstruction; the 1992–2006 coding and payment chronology; diagnosis-specific RAD criteria |
| IV. Current integrated landscape | 12–26 | Overlapping modes and equipment; nocturnal and multifactorial hypoventilation; contemporary clinical uses; trial evidence; historical claims growth and inappropriate-billing indicators; Medicare terminology; NCD 240.9; other coverage routes; DME supply and follow-up |

The histories are connected. The sections are not an assertion that three independent technologies emerged without interaction, or that patients belong to three mutually exclusive groups.

## Changes from the preceding version

The prior patient-first, exercise-based sequence has been replaced by the requested historical pathway sequence. The multifactorial and nocturnal physiology content remains, mainly within the integration section. Historical noninvasive ventilator claims growth and the OIG findings now have dedicated slides before the coverage update.

Slide text consists of dates, mechanisms, clinical populations, quantitative results, equipment distinctions, and coverage facts. The earlier learner prompts, debriefs, presenter instructions, and scripted narration are absent from the deck and companion. The companion is a factual reference document, not a teaching script. Detailed context also appears in the PowerPoint notes.

The README, slide IDs, source manifest, PowerPoint, and Markdown companion use the same organization. References retain stable `Rxx` identifiers across versions. Unused references are omitted from the companion bibliography.

## Evidence and policy scope

**Reference date:** September 19, 2026. The national COPD coverage policy is NCD 240.9, effective June 9, 2025. Its subsequent revision history and selected associated documentation sources were reviewed for this revision. NCD 240.4, LCD L33718, LCD L33800, NCD 280.1, and the cited coding and documentation sources are identified separately.

The October 28, 2026 addition of E0466, E0467, and E0468 to the face-to-face encounter / written-order-prior-to-delivery list is labeled forthcoming. Face-to-face requirements, written-order requirements, prior authorization, and substantive coverage criteria are different administrative elements.

The main slides summarize selected criteria. The companion and appendix retain the conditions and distinctions that are abbreviated on the main slides. This package is not a complete payer checklist and contains no verified local supplier directory, Intermountain order set, or code-specific Utah Medicaid authorization guide. Medicare is not used as a synonym for Medicaid or commercial coverage.

The historical OIG figures distinguish claims, beneficiaries, expenditures, and billing indicators. The increase in claims is not presented as a measured prevalence of clinical overuse or as proof that a specific earlier policy caused the increase. ONMAP's reported access barriers are attributed separately. The account of early coverage intent for a declining polio population is identified as a retrospective, uncited account in the source review rather than as archival proof.

The bibliography states the material actually consulted: supplied full papers, primary abstracts, official policies, institutional histories, or the published excerpt of *The Autumn Ghost*. No claim is made that every cited book or primary trial was newly reviewed in full for this revision.

## Visuals and reuse

The iron-lung photograph is Figure 1 from Hind, Polkey, and Simonds, *AJRCCM* 2017, DOI `10.1164/rccm.201702-0285CI`. It was extracted from the supplied PDF. The source caption does not identify the photograph's date, location, or subject. The slide identifies it as illustrative, not a verified Copenhagen image. Attribution is retained. Permission for unrestricted online redistribution has not been established.

All other diagrams, waveform illustrations, and tables are newly assembled native PowerPoint objects. Waveforms and CO₂ courses are conceptual, not patient recordings. The claims chart uses attributed historical values and a zero baseline. Timeline spacing is schematic and explicitly labeled not to scale. No font files or complete copyrighted source articles are included.

## Build

The build uses Python 3.11 or newer, `python-pptx`, Pillow, and PyMuPDF. LibreOffice supplies the PDF conversion. Fontconfig supplies font metrics for the Calibri-compatible font installed in the build environment. The renderer used Carlito for local Calibri substitution; fonts are not packaged.

From the package root:

```bash
python -m pip install -r build/requirements.txt
python build/content.py
python build/make_deck.py

mkdir -p qa/pdf
soffice -env:UserInstallation=file:///tmp/pap_rad_hmv_v2_lo \
  --headless --convert-to pdf --outdir qa/pdf \
  qa/PAP_RAD_HMV_Pathways_v2_all_slides.pptx
cp qa/pdf/PAP_RAD_HMV_Pathways_v2_all_slides.pdf PAP_RAD_HMV_Pathways_v2.pdf

python build/render_qa.py
python build/validate.py
```

`make_deck.py` creates the presentation with hidden appendix slides and an all-visible intermediate in `qa/` for complete PDF export. A rebuild replaces generated files in this package directory. Manual changes to the PowerPoint are not imported back into `content.py` or the renderer.

The historical files in the preceding package are unchanged. This version was generated as a separate local package; no GitHub or Google Drive files were modified.

## Validation

All 38 slides were rendered with LibreOffice, followed by contact-sheet review and selected full-size inspection. Layout corrections were applied before the final export. Automated checks found no missing rendered text, no shapes outside the slide boundary, no missing slide/reference anchors, and no incorrect appendix visibility flags. The main duration allocations sum to 40 minutes.

These checks do not constitute an actual timed rehearsal or a Microsoft PowerPoint rendering test on the presentation computer. Current presenter disclosures were not supplied and have not been invented. The source documents and relevant plan remain authoritative for actual coverage decisions.
