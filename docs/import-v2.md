# Version 2 repository import

Imported September 20, 2026 from `PAP_RAD_HMV_Pathways_v2_Package.zip`, supplied as the latest talk materials. The download ZIP was retained.

Source ZIP SHA-256:

```text
67cd5d5c4cb58cbeb30e2667aae044ce0b945164010540d93317787427310c7f
```

## File placement

| Source | Repository destination |
| --- | --- |
| Version 2 PowerPoint, PDF, and Markdown companion | Repository root, retaining the supplied filenames. |
| Package `build/`, `assets/`, and `qa/` | Corresponding root directories, preserving the relative paths used by the build scripts. |
| Package `README.md` and `SHA256.json` | [source-package-v2/](source-package-v2/README.md), unchanged as source records. |
| Initial deck, outline, JavaScript generator, npm manifest, and lockfile | `archive/v1/`, unchanged. |
| Initial repository README | `archive/v1/README.md`, updated to identify the archive and run npm from that directory. |

The root README now identifies version 2, links to the current artifacts and archived version, documents the Python build and system dependencies, and separates supplied validation claims from import checks. `.gitignore` now excludes Python environments and caches, temporary slide images, and PDF-conversion intermediates.

The source checksum manifest uses the original package's relative paths. Its `README.md` entry corresponds to `docs/source-package-v2/README.md`; its other entries correspond to the same paths at the repository root. The root README is repository documentation derived from the package README and is intentionally different.

## Checks completed

- ZIP integrity and member-path checks passed; all 15 entries in the supplied checksum manifest matched. Both preserved source-record files also match their ZIP entries exactly.
- The original five non-README files match their contents in repository commit `09e7ecb928d1c675a28ac748d233aaeed22258fc` byte for byte after relocation.
- The supplied PowerPoint has 38 slides and 38 speaker-note sections: 26 main slides and 12 hidden appendix slides. Its ZIP and XML structure parsed successfully.
- The supplied PDF has 38 pages. Slide IDs, companion anchors, reference IDs, nominal main duration of 40 minutes, and shape boundaries passed the import checks.
- All four Python scripts parsed successfully. In a temporary copy, `build/content.py` reproduced the companion and both JSON manifests byte for byte.
- In that temporary copy, `build/make_deck.py` completed using the supplied `python-pptx` and Pillow versions. All 38 slide texts matched after whitespace normalization; speaker notes and visibility flags matched exactly. Fontconfig selected Verdana instead of Calibri, producing 30 preliminary text-height warnings. This generated PowerPoint was not copied into the repository.
- Markdown file links resolved, the archived JavaScript generator passed `node --check`, and repository documentation changes passed `git diff --check`.

## Verification limits

The supplied presentation, PDF, companion, build sources, image, and QA records were retained unchanged. No new clinical or policy review, PDF rendering, or visual acceptance was performed during import. PyMuPDF was not installed in the inspected runtime, so the supplied rendered-text validator was not rerun. Before accepting a newly generated deck, install the pinned requirements, use appropriate font metrics, and render and inspect that new output.

`qa/validation.json` is the report supplied with version 2. The supplied validator writes a static visual-review description; it does not itself conduct that review. The source package's photograph attribution and redistribution qualification remain in the current README and companion.
