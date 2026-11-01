# Literature Review Chapter Change Record

## Delivery scope

Chapter 2 has been written in English and incorporated into `masar-master-project-report.docx` using the university report's existing styles and page geometry. The chapter is a critical narrative literature review for a **master's project in Information Systems**, not a master's thesis or a completed evaluation of Masar.

The verified layout contains 27 chapter pages, numbered 15–41. The complete report contains 50 pages because the original preliminary material, introduction, later chapter placeholders, and appendices remain in the university template. The Markdown chapter source contains approximately 8,264 whitespace-delimited words, including headings and table content. There are five comparative tables and one original conceptual figure. The report-wide bibliography contains 32 unique sources: the six existing sources and 26 additions.

## Branch and authorship

- Base: `develop`, commit `9445ba3` (`Add introduction chapter v.0.1`).
- Working branch: `chapter/literaturereview`.
- Git author configuration: `Mahmood-Anaam <eng.mahmood.anaam@gmail.com>`.
- Requested commit message: `Add literature review chapter v.0.1 `.
- Pull request destination: `develop`.

## Research materials used

The report introduction and existing `references.bib` established the contribution and its limits. Paperguide **Documents** supplied the existing research draft and associated source identifiers. My Notebooks content was not accessed or used. The user's subsequent instruction to use Documents and exclude My Notebooks was followed.

The existing Paperguide Literature Review and Introduction documents were consulted as research inputs. Neither source document was overwritten. Account-specific document links and internal source-manager identifiers are excluded from this public repository.

The project reference folder was used as a discovery inventory rather than treated as a set of fully appraised papers. A targeted subset was examined using paper abstracts, bibliographic records, available full-text passages, official standards descriptions, and original software documentation. `chapter-2-source-register.json` records the evidence basis and canonical source information for each cited source. It identifies the sources inspected only through abstracts and distinguishes official product descriptions from independent research evidence.

The full extracted text of *A Theory of Service Dependency* was read in two parts. Relevant full-text material from the saved DGA guideline was checked, particularly printed pages 116 and 118. Publisher and institutional records were used to verify selected claims and publication details. The accepted manuscript of the integrated continuity and disaster-recovery study and the full-text resilience-framework paper supplied additional checks. Access to an abstract or selected passages is not described as reading an entire paper.

## Academic restructuring

The existing Paperguide draft was treated as research input rather than copied into the report. Its abstract, journal-style results/discussion structure, generated PRISMA diagram, unsupported screening counts, generic evidence grades, and irrelevant biological-mechanism discussion were not carried into the chapter.

The chapter now follows a general-to-specific structure:

| Section | Purpose |
| --- | --- |
| 2.1 Introduction | Establish the review's argument and relation to Chapter 1 |
| 2.2 Review Approach and Evidence Boundaries | Describe the actual narrative-review approach and access limitations |
| 2.3 Business Continuity and Organizational Resilience | Establish organizational context and uncertainty |
| 2.4 Business Impact Analysis and Recovery Requirements | Connect services, time-dependent impacts, and recovery parameters |
| 2.5 Methods for Recovery Prioritization | Compare rules, multicriteria analysis, simulation, and optimization |
| 2.6 Dependencies and the Interpretation of Criticality | Examine direction, shared resources, and inheritance assumptions |
| 2.7 Decision Support Systems and Comparable Software | Compare research prototypes, commercial platforms, and technical tools |
| 2.8 Critical Synthesis and Positioning of TDRP and Masar | Define a bounded contribution and evaluation implications |
| 2.9 Chapter Summary | Connect the evidence to the methodology chapter |

The review distinguishes findings reported in a study, the interpretation developed in the chapter, and proposed behaviour of TDRP/Masar. No project implementation, practitioner study, or improvement in actual recovery performance is represented as already completed.

## Substantive decisions

1. **Respect the service-first scope.** Service criticality, MTPD, RTO, and relative recovery priority are the proposed derived outputs. RPO is reviewed as an adjacent concept, but its derivation is not added to the project scope.
2. **Separate priority from scheduling.** A service ranking is not described as a feasible recovery sequence, resource allocation, or orchestration plan.
3. **Separate requirements from capability.** A target RTO does not establish that the organization's technical or operational arrangements can achieve it.
4. **Treat the dependency maximum rule as a proposal.** The highest-supported-service rule is an explicit ordinal inheritance rule subject to evaluation, not a universal requirement attributed to a standard.
5. **Avoid unsupported novelty claims.** Existing BIA methods, commercial platforms, and dependency-aware tools are recognized. The contribution is the specification, implementation, and evaluation of the particular transparent procedure pursued in this project.
6. **Handle temporal evidence carefully.** Discrete intervals, threshold equality, missing responses, no observed threshold, and ties are discussed without inventing final methodological decisions that belong in Chapter 3.
7. **Qualify product comparisons.** Official descriptions of Riskonnect, Fusion, Veeam, ReaR, and Borg support a functional comparison. No installation, price comparison, benchmark, or product superiority is claimed.
8. **Use an honest review designation.** The chapter is a focused critical narrative review. It does not claim systematic-review completeness or PRISMA compliance without a screening audit trail.

## Reference and metadata work

The six existing bibliography numbers remain unchanged. New numbers 7–32 follow first appearance after those already introduced in Chapter 1. The original six BibTeX records were retained, and the additional sources were appended to `references.bib`.

Final issue years were checked where Paperguide's online-first year differed from the journal citation:

| Source | Year used in the chapter |
| --- | --- |
| Steen et al. resilience framework | 2024 |
| Aghabegloo et al. physical-asset criticality | 2024 |
| Laugé et al. dynamic infrastructure dependencies | 2015 |
| Henry and Ramírez-Márquez resilience metrics | 2012 |

The Merz et al. page range was completed as 156–174. DRAMM and automated BIA were represented as conference papers instead of copying their misleading generic report classifications. Institutional and corporate authors were retained for standards and product documentation. Unknown publication dates were not invented. Public web documentation carries an access date of 1 November 2026.

One narrow correction was made in Chapter 1: the sentence specifically discussing the Digital Government Authority's guidance now cites source [2], including the field metadata and hyperlink target, rather than source [1]. The remaining introductory prose and the original Zotero library identities were preserved.

## Word fields and citation navigation

The existing IEEE Zotero document preference and bibliography field were retained. The final DOCX contains 53 unique Zotero citation fields, including 45 additions. New fields contain structured `CSL_CITATION` data and cached IEEE display text. The existing six items retain their original Zotero URIs. New sources use embedded CSL item data and canonical DOI or source URLs; no fictional Zotero library item identifiers were assigned.

Blue in-text citations link to corresponding bibliography bookmarks. The new bibliography entries are included within the cached result of the existing Zotero bibliography field. External DOI and website links in the added bibliography entries are active. New figure/table captions include `STYLEREF` and `SEQ` fields, and the contents and lists retain their field structure with materialized page-number results.

The [ZoteroCiteLinker README](https://github.com/sBaydin/ZoteroCiteLinker) was read to understand its citation linking and colour behaviour. Equivalent bookmark links and blue formatting were written into the document. **The VBA add-in was not executed**, because native Microsoft Word and Zotero were not available in this environment. New embedded sources are not claimed to have been imported into or synchronized with the user's local Zotero library. Native Word/Zotero refresh and the add-in's handling of those embedded sources have not been verified. The delivered file already contains the rendered citation text, references, navigation links, and populated indexes; no refresh is required merely to read this version.

## Tables figure and typography

The university A4 page dimensions, margins, Calibri body typography, multilevel chapter numbering, and existing introductory layout were retained. New table headers use a restrained light-blue fill; body rows alternate white and a pale neutral background. Table rows expand with content and repeat header rows when a table spans pages. Captions stay with their associated tables or figures. Paragraph spacing was checked at table boundaries.

Bibliography entries were normalized to a consistent 11-point hanging-indent format, with individual entries kept together across page boundaries. Their original citation identities and the first six reference numbers remain intact.

The original conceptual relationships were explored with FigJam. Account-specific design links are excluded from the public repository. A compact SVG layout of the same relationships was prepared for publication and embedded as a high-resolution PNG. The repository includes both vector and raster versions so the figure remains reproducible. Its arrows indicate derivation and information relationships, not predicted failure propagation.

## Validation and its limits

The DOCX was structurally checked for ZIP integrity, well-formed XML, balanced field boundaries, unique citation identifiers, unique bookmark identifiers, resolved internal hyperlink targets, and correspondence between numbered citations and the 32 bibliography entries. Citation tokens such as `[@key]` do not remain in Word display text. Python helper files were syntax-checked.

The packaged Word renderer was attempted on the original live-field template. Its LibreOffice PDF-export path exhausted memory, including after an ODT fallback. Layout QA therefore used a **separate rendering copy** that removes field instructions and controls while preserving their displayed results, formatting, hyperlinks, bookmarks, and page geometry. The final DOCX retains its fields. The rendered PDF remains a layout check, not a native Microsoft Word or Zotero compatibility test.

Page images were reviewed across the complete report, with detailed checks of contents, abbreviations, all five tables, the conceptual figure, bibliography, and chapter boundaries. Populated page caches were checked against the rendered pagination. The verified Chapter 2 range is 15–41, and Chapter 3 begins on page 42. The existing repository PDF was synchronized with the reviewed visible-content rendering. Later uncompleted chapters and preliminary placeholders remain as supplied because this change concerns Chapter 2.

The review is academically edited and source-qualified; it is not represented as having received university examiner approval, independent peer review, or an empirical validation of TDRP. The remaining native Word/Zotero interoperability limitation is stated above rather than hidden behind a publication-readiness claim.

## Files and maintenance

| File | Change |
| --- | --- |
| `masar-master-project-report.docx` | Chapter 2, citation fields and links, bibliography, navigation caches, abbreviations, narrow DGA correction |
| `masar-master-project-report.pdf` | Synchronized visible-content rendering |
| `chapter-2-literature-review.md` | Reviewable English chapter source with citation keys |
| `references.bib` | Added 26 sources while retaining existing records |
| `chapter-2-source-register.json` | CSL metadata, IEEE numbers, evidence basis, public verification links |
| `chapter-2-layout-map.json` | Verified heading and caption page locations |
| `figures/literature-concept-map.svg` | Reusable vector figure |
| `figures/literature-concept-map.png` | High-resolution Word figure |
| `tools/` | Content-building, figure-rendering, field-preserving layout preparation, page-cache and validation helpers |

The helpers require Python with `python-docx`, `lxml`, `PyMuPDF`, and the packaged document renderer. Rebuilding starts from the unmodified `develop` Word document, uses the Markdown and source register, then requires layout rendering and page-cache materialization. The delivered DOCX has already passed those steps. Rendering-only files and page images are QA intermediates and are not committed.
