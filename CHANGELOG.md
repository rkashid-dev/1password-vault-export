# Changelog

All notable changes to this project will be documented in this file.

The format follows [Keep a Changelog](https://keepachangelog.com/en/1.0.0/).

---

## [3.0.0] — 2026-03-16

### Added
- Full coverage of all 23 official 1Password item categories
- Per-category fixed column definitions sourced directly from `op item template get <category>`
- Dynamic/custom field capture with no-data-drop policy — duplicates preserved under `_raw` or numeric suffixes
- TOTP/OTP fields exported as both raw `otpauth://` URI and live code at export time
- Embedded file attachment download via `op document get <fileId>`
- Document body download for `DOCUMENT` category items
- Auto re-authentication with single retry on session/auth errors
- `$INCLUDE_ARCHIVE` flag to optionally include archived items
- Timestamped output folder — safe to run multiple times without overwriting
- Structured `export_log.txt` with `[INFO]`, `[SUCCESS]`, `[WARN]`, `[ERROR]` levels
- `Other_<CATKEY>.csv` fallback for unknown or future categories

### Changed
- Refactored field extraction into `Extract-CategoryFields` function with `$consumed` HashSet to prevent duplicate column output
- Replaced ScriptBlock re-auth pattern with explicit `string[]` argument passing via call operator (`&`) — fixes silent variable capture failures in PowerShell closures
- `Write-CategoryCsv` now uses `[ref]` for written count to avoid pipeline pollution

### Fixed
- Fields consumed by fixed extraction no longer duplicated in dynamic extras columns
- Notes field (`notesPlain`) correctly collapsed to single line for CSV compatibility

---

## [2.0.0] — 2026-02-01

### Added
- Support for 15 additional item categories beyond Login/Password/Secure Note
- Per-vault output subdirectories

### Changed
- Switched from `Export-Csv` pipeline to explicit column list for consistent column ordering

---

## [1.0.0] — 2026-01-10

### Added
- Initial release
- Login, Password, and Secure Note export
- Basic attachment download
- Single-vault and all-vault modes
