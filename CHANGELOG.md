# Changelog

## [1.0.1.0] - 2025-09-09
### Added
- Toggle for suppressing card selections (`allowSelections` in `stateSettings`); clears cross-filters when disabled.
- Default selection logic now triggers on toggle enable (clears existing selections if needed).
- Cursor set to "default" for the entire visual to prevent pointer on hover.
- Drill-through capability: Users can now right-click a KPI card to access drill-through options, navigating to configured report pages with category context (requires report setup).
- Enhanced the wrapText function to properly handle UNICHAR(10) (newline characters) in "Comparison Value" and measure names, ensuring all text after newlines is rendered with correct spacing, including support for empty lines.

### Fixed
- Selection state persistence when toggling `allowSelections` off.
- Default selection not applying after toggling `selectedDefault` on.
- Resolved issue with delta calculations for "Comparison Value" when using measures with fixed decimal formatting (e.g., 2 decimals), ensuring absolute and relative deltas display correctly.
- Corrected text string handling in "Comparison Value" to preserve non-numeric content (e.g., custom text or multi-line strings) when "Delta Type" is set to "None", restoring functionality lost due to numeric parsing changes.

## [1.0.0.8] - 2025-09-04
### Added
- Single KPI support (when no category is added)
- No-data and max-cards (1000) handling with warnings.

## [1.0.0.7] - 2025-09-02
### Added
- Initial KPI grid visual with measure, category, and comparison support.