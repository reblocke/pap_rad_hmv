# PAP, RADs, and Home Mechanical Ventilators

Working materials for a 45-minute talk for pulmonary and critical care medicine (PCCM) fellows on positive airway pressure (PAP), respiratory assist devices (RADs), and home mechanical ventilators (HMVs).

## Status

This is a working draft. The current PowerPoint contains 24 slides; the outline targets 22 slides with approximately 40 minutes of content and 5 minutes of discussion. The outline and deck may evolve together as the talk is refined.

The outline retains its references and its **Claims to verify or soften before presenting** section. Its stated policy-verification date is September 18, 2026. Repository publication does not constitute a new clinical or policy review; resolve the outstanding notes and check applicable coverage requirements before presenting.

## Learning objectives

- Explain how the engineering histories of home ventilation and PAP converged in capability while their coverage categories remained distinct.
- Distinguish treatment mode, device capabilities, and billing or coverage category.
- Select and justify home respiratory support from clinical need and the consequences of interruption.

## Files

| File | Purpose |
| --- | --- |
| [PAP_RAD_HMV_talk_outline.md](PAP_RAD_HMV_talk_outline.md) | Talk structure, teaching points, references, and outstanding verification notes. |
| [PAP_RAD_HMV_draft.pptx](PAP_RAD_HMV_draft.pptx) | Current editable draft presentation, including speaker notes. |
| [PAP_RAD_HMV_deck_generator.js](PAP_RAD_HMV_deck_generator.js) | JavaScript source for generating the presentation with PptxGenJS. |

## Build the presentation

Install Node.js 22 or newer, including npm, then run:

```sh
git clone https://github.com/reblocke/pap_rad_hmv.git
cd pap_rad_hmv
npm ci
npm run build
```

The build uses PptxGenJS 4.0.1 with dependencies recorded in `package-lock.json`. It writes `PAP_RAD_HMV_draft.pptx` beside the generator, **replacing the existing file**. Commit or save any manual PowerPoint edits before rebuilding; those edits are not incorporated into the JavaScript source.

The presentation uses Cambria and Calibri. Viewing software may substitute other fonts if they are unavailable, so check the rendered slides on the presentation computer.

Dependency note: at initial publication, `npm audit` reports high-severity denial-of-service advisories in the transitive `image-size` dependency ([ICNS](https://github.com/advisories/GHSA-w3rx-r6r6-pgpr), [JXL and HEIF](https://github.com/advisories/GHSA-5p2g-fcmc-qvqq)). This generator does not load images. Review dependency updates before adding image inputs.

## Working on the talk

Update the outline as the talk develops. For slide changes that should survive regeneration, edit the generator and rebuild. Review the resulting PowerPoint and speaker notes before committing an updated deck.

The initial repository preserves the supplied draft PowerPoint. Build verification uses a separate checkout so it does not overwrite that artifact.
