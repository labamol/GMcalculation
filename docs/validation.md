# Validation evidence

Tested in the cloud workspace using Python 3.12.14.

- Automated suite: `python -m pytest -q`. Includes Streamlit AppTest loading, input validation, and rendering all generated table types.
- Synthetic demonstration: `python examples/create_demo.py /tmp/gm-demo`, followed by all three processing functions. All stages PASS; 40 reconciliation/conversion checks in the margin report. Banking / Account A projects 6 INR Crores quarterly cost and 40% GM against both 10-crore revenue measures. MetLife's genuine zero cost remains valid.
- Spreadsheet integration: `python tools/verify_excel.py` with LibreOfficeDev 26.8.0.0.alpha0. Unfiltered Banking subtotal and grand total both equal INR 20,000,000 (no double counting). With Banking rows filtered/hidden, both become zero. The filter criteria and hidden-row state are saved programmatically; desktop interactive dropdown behavior was not manually exercised.
- Dependency consistency: `python -m pip check`.
- Streamlit server started headlessly; local health endpoint responds successfully.

Excel desktop itself and real business workbooks were not available for acceptance testing. Input preservation is semantic for supported .xlsx features, not a byte-identical package round trip. Formula-cache freshness requires user confirmation after external recalculation. Reports surface that limitation even though synthetic subtotal behavior was checked with LibreOffice.

The integration check is optional and separate from CI because LibreOffice is not a runtime dependency. The application never invokes it on uploaded workbooks. It only processes synthetic fixtures in the integration script.
