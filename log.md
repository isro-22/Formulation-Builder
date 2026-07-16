# Change Log

## 2026-07-16

| Issue / Bug | Action | Update |
| --- | --- | --- |
| Material table borders disappeared when input rows exceeded the original template capacity. | Added a final material-table border pass across active rows for all rendered phases. | Borders are retained for expanded phase tables, including main phases and premix phases. |
| Phase tables could appear without a visible separator, especially around premix sections. | Added a standardized one-row separator and removed unused rows instead of hiding them. | Rendered phase tables now keep one visible blank row between sections. |
| Empty premix tables were rendered even when no materials existed for that premix phase. | Removed empty premix sections from the generated workbook layout. | Only premix phases with material rows are shown. |
| Online Material Lookup formulas could appear as `=@IFNA(...)` or use compatibility-prefixed lookup syntax. | Changed lookup formula generation to `=IFNA(XLOOKUP(...), fallback)` and added formula normalization for `=@IFNA` to `=IFNA`. | Generated online lookup formulas no longer contain `=@IFNA`, `@IFNA`, or `_xlfn.XLOOKUP`. |
| Approval layout could become unstable after dynamic section resizing. | Hardened merge cleanup and regenerated the approval block after section layout updates. | Approval columns and row heights remain stable after large dynamic inserts. |
| Sensory metadata cells `J13:J16` raised Excel numeric-format warnings. | Added numeric coercion and numeric format for `Impact`, `Flavor Aroma`, `Irritation`, and `Cooling` cells. | Numeric inputs in `J13:J16` are saved as numeric cells with `0.00` format. |
| Tests relied on static row numbers that changed after dynamic row deletion. | Updated regression tests to use runtime phase positions and section ranges. | Tests now validate dynamic workbook layout instead of fixed template row assumptions. |

### Verification

- Focused layout and lookup tests passed.
- Generated workbook XML was scanned and confirmed free of `=@IFNA`, `@IFNA`, and `_xlfn.XLOOKUP`.
- Full generator suite still has three pre-existing failures related to Excel date values being read back as `datetime` instead of `date` or string.
