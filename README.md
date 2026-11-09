# XLSmart Development Workspace

Professional development workspace for the XLSmart Excel VBA Add-in.

## Project Structure

```
xlsmart-development/
├── src/                # VBA source code (modules, classes, forms)
│   ├── Modules/        # 20 standard modules (.bas)
│   ├── Classes/        # 2 class modules (.cls)
│   └── Forms/          # 11 UserForms (.frm + .frx pairs)
├── workbook/           # Working copy of XLSmart_User.xlam
├── installer/          # XLSmart_Installer.xlsm + version.json
├── release/            # Release artifacts for GitHub
├── scripts/            # PowerShell automation (export/import/compare VBA)
├── backups/            # Timestamped backups
└── docs/               # Documentation and comparison reports
```

## Source Code Inventory

| Type       | Count | Location        |
|------------|-------|-----------------|
| Modules    | 19    | `src/Modules/`  |
| Classes    | 2     | `src/Classes/`  |
| UserForms  | 11    | `src/Forms/`    |
| `.frx`     | 11    | `src/Forms/`    |
| **Total**  | **43** | `src/`         |

## Key Files

- **Original add-in** (never modify): `%APPDATA%\Microsoft\AddIns\XLSmart_User.xlam`
- **VBA export source of truth**: `%APPDATA%\Microsoft\AddIns\VBA_Export`
- **Working copy**: `workbook/XLSmart_User.xlam`
- **Installer**: `installer/XLSmart_Installer.xlsm`
- **Version**: `installer/version.json` (v1.0.4)
- **GitHub Repository**: https://github.com/lpmentionpro/XLSmart-Smart-Development

## Scripts

| Script              | Purpose                                      |
|---------------------|----------------------------------------------|
| `Export-VBA.ps1`    | Export VBA modules/classes/forms from Excel  |
| `Import-VBA.ps1`    | Import VBA source into a workbook            |
| `Compare-VBA.ps1`   | Compare workbook VBA against `src/`          |

All scripts use Excel COM with proper cleanup. Requires Excel installed.

## Rules

1. Never modify the original `XLSmart_User.xlam` in AddIns.
2. Never delete `VBA_Export` — it is the VBA source of truth.
3. Do not push to main repository without testing and verification.
4. GitHub PowerDashboard code is old; local export is authoritative.
5. Preserve `.frx` files alongside their `.frm` counterparts.
6. Create timestamped backups before any workbook modification.
7. Installer XLM is not a VBA source module — do not export it as one.
