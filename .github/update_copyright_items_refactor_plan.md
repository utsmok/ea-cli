# Refactoring Plan for `update_copyright_items` Function

## Overview
The `update_copyright_items` function in `easy_access/db/update.py` is a 400+ line monolithic function with convoluted logic for merging copyright items. This plan outlines a step-by-step refactoring to improve maintainability, testability, and integration with the Settings system.

## Current Issues
- **Size**: 400+ lines, hard to understand and maintain.
- **Nested Functions**: `change` and `compare_fields` are closures, difficult to test.
- **Hardcoded Logic**: Field definitions and merge rules embedded in the function.
- **Repetitive Code**: Type casting and comparison logic duplicated.
- **Poor Error Handling**: Broad exception catching masks issues.
- **Tight Coupling**: Direct DB interactions and logging mixed with business logic.

## Integration with Settings
The `Settings` class (from `easy_access/settings.py`) contains field definitions that can be leveraged:
- `data_settings.data_entry_cols`: List of `ColInfo` with field metadata (names, types, dropdowns).
- `data_settings.new_fields`: Dictionary of new field configurations.
- `classification_options`: List of valid classification values.
- Use these to dynamically populate `added_fields` and `changeable_fields` in `merge_rules.py`.

## Detailed Refactoring Steps

### Phase 1: Preparation (1-2 days)
1. **Analyze Current Usage**: Map all call sites of `update_copyright_items` to understand dependencies.
   - **Call Sites Identified:**
     - `easy_access/db/update.py:887`: `await update_copyright_items(settings=settings, data=update_df, overwrite=True)` - Used for derived field calculations
     - `easy_access/sheets/analysis.py:185`: `await update_copyright_items(settings, df)` - Used for sheet analysis processing
     - `easy_access/old_main.py:620-623`: `await update_copyright_items(settings=self.settings, data=update_df)` - Used for faculty sheet processing
     - `easy_access/old_main.py:1061`: `await update_copyright_items(settings=self.settings, data=with_file_exists)` - Used for file existence updates
     - `dashboard/data.py:628-633`: `await update_copyright_items(SETTINGS, full_data_list, update_relations=False, overwrite=True, user_info=user_info)` - Used for dashboard data updates
   - **Common Pattern**: All calls pass `settings` first, `data` second. Optional parameters: `update_relations`, `overwrite`, `user_info`.

2. **Create Test Data**: Use available raw data, faculty sheet data, and DB items to create comprehensive test cases.
   - **Available Data Sources:**
     - `copyright_data_with_pdf.parquet`: 2426 rows × 58 columns with all relevant fields
     - `main_data.parquet`: Additional dataset
     - `sample_dataset.xlsx` & `sample_dataset_full.xlsx`: Excel test data
     - `db.sqlite3`: Database with existing processed items
     - `raw_copyright_data/` & `faculty_sheets/`: Raw input data
   - **Test Strategy**: Create fixtures with subsets of real data for unit/integration tests.

3. **Document Merge Rules**: Write clear documentation for ranking logic, field priorities, and edge cases.
   - **Added Fields (script-added, prioritized):**
     - `workflow_status`: [Done > InProgress > ToDo] - Higher priority wins
     - `retrieved_from_copyright_on`: None (no priority) - Latest date wins
     - `possible_fine`: None (no priority) - Higher value wins
     - `infringement`: [YES > NO > UNDETERMINED] - Higher priority wins
     - `file_exists`: [False, 0, True, 1] - True values always update
   - **Changeable Fields (checker-editable, prioritized):**
     - `manual_classification`: [OPEN_ACCESS > KORTE_OVERNAME > MIDDELLANGE_OVERNAME > LANGE_OVERNAME > EIGEN_MATERIAAL_* > ONBEKEND > LICENTIE_BESCHIKBAAR > NIET_GEANALYSEERD > IN_ONDERZOEK > VERWIJDERVERZOEK_VERSTUURD] - Higher priority wins
     - `manual_identifier`: None - Longer string wins
     - `remarks`: None - Longer string wins
     - `scope`: None - Longer string wins
   - **Core Fields**: Not updated (material_id, title, author, etc.)
   - **Comparison Logic**: Rankings use list index (lower = higher priority). For non-ranked fields: longer strings win, higher numbers/dates win.

### Phase 2: Extract and Modularize (3-5 days)
4. **Extract Nested Functions**:
   - Move `change` to `record_field_change(db_item, field, new_value, old_value, reason)`.
   - Move `compare_fields` to `compare_and_update_fields(new_item, db_item, fielddict, changes)`.
   - Make them pure functions by passing all dependencies as parameters.

5. **Externalize Field Definitions**:
   - Move `added_fields` and `changeable_fields` to `easy_access/merge_rules.py`.
   - Create a function `build_merge_rules_from_settings(settings: Settings)` to populate these from Settings data.
   - Integrate with `data_settings.data_entry_cols` for dynamic field lists.

6. **Create Type Casting Helpers**:
   - `cast_datetime_value(value)`: Handle datetime parsing with fallbacks.
   - `cast_enum_value(value, enum_class)`: Safe enum casting.
   - `cast_numeric_value(value, target_type)`: Unified int/float casting with rounding.
   - `normalize_file_exists(value)`: Handle file_exists boolean normalization.

### Phase 3: Break Down Main Function (3-4 days)
7. **Split into Smaller Functions**:
   - `preprocess_input_data(data, settings)`: Handle DataFrame vs list input, standardize, separate new vs update.
   - `process_new_copyright_items(new_items, settings)`: Create new CopyrightItem instances.
   - `process_existing_copyright_items(update_items, merge_rules, settings)`: Handle updates with comparison logic.
   - `execute_bulk_database_operations(new_objects, updates, changelist, settings)`: Perform bulk creates/updates.

8. **Simplify Comparison Logic**:
   - Implement strategy pattern: `FieldComparisonStrategy` with subclasses for ranking, length, magnitude.
   - Add early returns for unchanged fields.
   - Replace magic numbers (e.g., `len(changes) >= 3`) with named constants like `MIN_CHANGE_KEYS = 3`.

### Phase 4: Improve Error Handling and Testing (2-3 days)
9. **Enhance Error Handling**:
   - Define custom exceptions: `MergeConflictError`, `InvalidFieldValueError`, `DatabaseOperationError`.
   - Replace broad `except Exception` with specific catches.
   - Add logging with structured context (material_id, field, values).

10. **Add Comprehensive Tests**:
    - Unit tests for each helper function using pytest fixtures.
    - Integration tests with real data: load raw data, process through refactored function, verify DB state.
    - Edge case tests: invalid types, missing fields, ranking conflicts.
    - Performance tests: benchmark before/after refactoring.

### Phase 5: Integration and Validation (1-2 days)
11. **Integrate with Settings**:
    - Modify `update_copyright_items` to accept `settings: Settings` parameter.
    - Use `build_merge_rules_from_settings(settings)` to get field definitions.
    - Ensure backward compatibility with existing calls.

12. **Final Validation**:
    - Run full pipeline with refactored function.
    - Verify no regressions in data processing.
    - Update documentation and add changelog entry.

## Benefits
- **Maintainability**: Smaller functions are easier to modify and debug.
- **Testability**: Isolated components can be unit-tested thoroughly.
- **Reliability**: Better error handling prevents silent failures.
- **Performance**: Optimized comparisons and bulk operations.
- **Flexibility**: Dynamic field definitions from Settings allow runtime configuration.

## Risks and Mitigations
- **Breaking Changes**: Ensure all call sites are updated; add deprecation warnings if needed.
- **Performance Impact**: Profile changes; optimize if bulk operations slow down.
- **Data Integrity**: Extensive testing with real data to catch merge logic errors.

## Timeline
- Total: 10-16 days
- Phase 1: 1-2 days
- Phase 2: 3-5 days
- Phase 3: 3-4 days
- Phase 4: 2-3 days
- Phase 5: 1-2 days

## Dependencies
- Available test data: raw data, faculty sheets, DB items.
- Settings integration: `data_settings.data_entry_cols`, `classification_options`.
- Existing helpers: `safe_*` functions from `utils.py`.
