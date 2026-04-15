# Changelog

## [5.2] - 2026-04-15

### Critical Fixes
- AdmissionDate field now required in Students sheet.
- `sendEmailReceipt` fixed with proper column mapping.
- `getStudentFeeSummary` now handles missing dates gracefully.

### Architecture
- Confirmed dynamic fee calculation pattern: pending and total amounts never stored, computed from Payments sheet on-demand.

### Improvements
- Better error handling in `getMonthlyFee`.
- Email receipts use normalized field lookups.
- Seed data includes complete admission dates.

### Migration
- Run `setupSystem()` to add AdmissionDate column to existing sheets.

### Test
- Verify fee calculations work correctly for students with various admission dates.
