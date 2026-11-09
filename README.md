# FlexCreds

Planning tool for estimating Workday Flex Credit consumption and API call metering in client workshops. Single self-contained page: open `index.html` in a browser. No build step.

**Use with review.** This is a planning aid, not a quote. It follows Workday's Flex Credits rate card, Flex Credits and Platform Entitlement Policy and Platform Entitlements overview (all dated 5 Nov 2026), plus public vendor documentation for system API patterns. System-level API call estimates are modelled from typical designs, not measured. Review every figure against the client's contract and Platform Consumption Console before sharing it.

## Contents
- `index.html`: the app (logic, styles and logos inline).
- `validation/`: independent check of the page's calculation.
  - `scen.py`: independently transcribed rate card and 15 test scenarios.
  - `harness.js`: runs the page's own calculation code for each scenario (`node harness.js`, writes `results.json`).
  - `build.py`: builds `Flex_Credit_Validation_Model.xlsx`, an all-formula Excel model that compares against the page results (needs `openpyxl`; recalculate with LibreOffice).

## Before client use
- Replace logos with approved brand assets and get Deloitte brand/IP sign-off on the name.
- Remove the Google Fonts link for internal hosting.
- Have a second person check the rate card transcription, and cross-check against Workday's Flex Credit Estimator.
