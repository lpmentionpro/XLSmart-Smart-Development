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
- Fixed `GetDashboardTemplate` missing module compile error (2026-11-08)
  - Root cause: `modDashTemplate.bas` was present in `src/Modules/` but never imported into `XLSmart_User.xlam`
  - Resulted in compile error: "Sub or Function not defined" at `s = GetDashboardTemplate()`
  - Imported `modDashTemplate` into `workbook/XLSmart_User.xlam` (3208 lines)
- Enhanced range picker in `frmDashboard` with Ctrl+Shift+Down+Right behavior (2026-11-08)
  - Auto-detects data table from `ActiveCell.CurrentRegion` upon form load
  - Pre-fills `Application.InputBox(Default:=...)` with full table address so users don't have to drag
  - Clicking a single cell or header automatically expands to full data region (Ctrl+Shift+Down+Right)
  - Replaced hardcoded `gridlinesShow` with dynamic `cmbGridlinesComboBox` selection
- Fixed `frmDashboard` PowerDashboard compile error (2026-11-08)
  - `btnCreateCommandButton_Click` was calling `CreateDashboard` with 7 arguments
  - Updated `modPowerDashboardEngine` requires 8 parameters (added `gridlinesShow`)
  - Added missing `gridlinesShow = True` as 8th argument
  - Compile error "Argument not optional" now resolved
  - Form imported into workbook via `Import-VBA.ps1`

### Changed
- Updated `modPowerDashboardEngine` to latest version with gridlinesShow support (2026-11-08)
  - Added gridlinesShow parameter to CreateDashboard signature
  - Fixed currencyCode parameter name (was currencyCodeCode in old version)
  - Added UTF-8 file writer (WriteUtf8 helper)
  - Added error code +104 for file write validation
  - Workbook SHA-256: `aaa784311b9e620124a5eea38d8d5e336df83fdaca178bed4095d191176be7e0`
  - Previous SHA-256: `d95094b24b7975a4fc1f6843b7f4f8ccd140ac18fb443ca6a4b20e32f2aff294`

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
