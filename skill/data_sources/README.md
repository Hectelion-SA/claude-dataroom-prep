# DD reference lists (not bundled)

Step 5bis of the pipeline (Top 50 missing documents) needs a per-country DD
reference list to compare the built dataroom against. These lists are
firm-specific methodology, so they are **not shipped in this public repo** —
drop your own here to enable the step; without them, 5bis is skipped
gracefully and the rest of the pipeline runs normally.

Expected files, one workbook per country:

- `Listes DD - FRANCE.xlsx`
- `Listes DD - SUISSE.xlsx`

Each workbook has one tab per sector (BTP, Industrie, Services & Conseil,
Tech, Distribution, Santé, Immobilier, Hôtellerie-Restauration, Transport &
Logistique, Agroalimentaire — or your own set), with these columns:

| Column | Description |
|---|---|
| Catégorie | N1 dataroom folder this requirement maps to |
| Réf. | Requirement reference/ID |
| Document à demander | The document/information to request from the client |
| Priorité | Haute / Moyenne / Basse |
| Période | Relevant period, if applicable |
| Base légale | Legal/contextual basis for the requirement |

Pass `--country FR|CH --sector <tab name> --dd-dir <this folder>` to
`run_pipeline.py` (or answer the corresponding questions in `/dataroom-prep`)
once your files are in place.
