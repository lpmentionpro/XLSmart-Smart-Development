# Changelog

All notable changes to XLSmart are documented here.

Format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/).

## [Unreleased]

### Added
- Professional VS Code development workspace structure
- PowerShell scripts for VBA export, import, comparison and release build
  - `scripts/Export-VBA.ps1` - Export VBA components from Excel workbook
  - `scripts/Import-VBA.ps1` - Import changed VBA components into workbook
  - `scripts/Compare-VBA.ps1` - Compare source files vs workbook components
  - `scripts/Build-Release.ps1` - Validate and package release artifacts
- Documentation and comparison reports
- Phase 4 PowerDashboard comparison report (local source vs. embedded xlam copy): `docs/PHASE4_POWERDASHBOARD_COMPARISON.md` with evidence in `docs/phase4-evidence/`

### Added
- Automatic update feature in `modHelpMenu` (2026-11-08)
  - Replaced browser hyperlink redirect with background download and install
  - Downloads latest xlam from GitHub raw URL directly
  - Verifies SHA-256 hash before installation
  - Creates backup of existing version before overwrite
  - Reloads add-in and prompts Excel restart
- Updated `modVersion` to point to XLSmart-Smart-Development repository
  - `UPDATE_JSON_URL` updated from old `xlsmart-release` repo
  - `APP_VERSION` bumped to 1.0.3

### Fixed
- Released version 1.0.4 so Excel detects and installs the verified Power Dashboard update (2026-11-09)
  - Restores browser dashboard controls by fixing the generated JavaScript parse error
  - Uses a modeless, two-step range picker so Ctrl+Shift+Down/Right selects directly on the worksheet
  - Refreshes the downloaded add-in SHA-256 in `installer/version.json`
- Replaced truncated `modDashTemplate` with self-contained `BuildDashboardHTML` in `modPowerDashboardEngine` (2026-11-09)
  - Root cause: `modDashTemplate.bas` in the repository was incomplete/truncated at line 3209 inside `Sec08_HeaderLogo_14()`, causing compile error "Sub or Function not defined" on `Sec08_HeaderLogo_15()`
  - Restored full, self-contained HTML dashboard builder in `modPowerDashboardEngine.bas` (306 lines) with `gridlinesShow`, `smartLock`, multi-currency support, and UTF-8 encoding
  - Completely removed broken `modDashTemplate.bas` from source and workbook
- Enabled manual range selection in `frmDashboard` with Ctrl+Shift+Down/Right (2026-11-09)
  - Updated `PickSingleTableRange` to allow natural manual selection via keyboard (`Ctrl+Shift+Down` and `Ctrl+Shift+Right`) or mouse drag
  - Removed forced single-cell auto-expansion and pre-filled defaults that interfered with keyboard selection
- Fixed `frmDashboard` PowerDashboard compile error (2026-11-08)
  - `btnCreateCommandButton_Click` was calling `CreateDashboard` with 7 arguments
  - Updated `modPowerDashboardEngine` requires 8 parameters (added `gridlinesShow`)
  - Added missing `gridlinesShow` as 8th argument
  - Compile error "Argument not optional" now resolved
  - Form imported into workbook via `Import-VBA.ps1`

### Changed
- Updated `modPowerDashboardEngine` to self-contained engine with gridlinesShow support (2026-11-09)
  - Added gridlinesShow parameter to CreateDashboard signature
  - Fixed currencyCode parameter name (was currencyCodeCode in old version)
  - Added UTF-8 file writer (WriteUtf8 helper)
  - Added error code +104 for file write validation
  - Workbook SHA-256: `D961FE5AAFC03D90F8E7824C05A1D8CDB1FA2563D8DB3AB6320A7727FC1C19EE`

## [1.0.2] - 2026-07-30

### Changed
- Updated installer download URL to GitHub release repo

### Notes
- `installer/version.json` points to `https://raw.githubusercontent.com/lpmentionpro/XLSmart-Smart-Development/main/workbook/XLSmart_User.xlam`

## [1.0.1] - Unknown

- See git history after repository initialization for details.

## [1.0.0] - Unknown

- Initial release of XLSmart Excel VBA Add-in.

[Unreleased]: https://github.com/lpmentionpro/XLSmart-Smart-Development/compare/v1.0.2...HEAD
[1.0.2]: https://github.com/lpmentionpro/XLSmart-Smart-Development/releases/tag/v1.0.2
