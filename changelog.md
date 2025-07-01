# 📦 Changelog – SpellCurrencyUniversal

All notable changes to this project are documented here.

## [v1.0.0] – 2025-07-01

### Added
- Core function: `SpellCurrencyUniversal(cell)`
- ISO-based and Unicode-based currency detection
- Modular subfunctions:
  - `ExtractNumericAmount`
  - `DetectISOFromCell`, `DetectISOFromDisplayText`
  - `ConvertToIndianWords`, `ConvertToWesternWords`
  - `GetCurrencyMetadataByISO`
- Smart validator for invalid/junk inputs
- Metadata support for 12+ currencies, including zero-decimal formats
- `README.md` with installation and extension guide
- Test dataset with 100 currency samples

### Notes
- Add-in exported as `.xlam` and `.bas` for reuse
- VBA project optionally password-protected
