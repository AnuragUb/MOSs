# Exception Marking Feature

## Overview
The exception marking feature allows users to mark specific rows as exceptions, which prevents automatic filling of Film/Album Title and Title Prefix fields during export.

## How to Use

### 1. Mark Rows as Exception
1. Select one or more rows using the checkboxes
2. Click the "Mark Row" button
3. Choose "Mark Exception" from the modal
4. In the exception settings modal, select which fields should be manual:
   - **Film/Album Title**: When checked, prevents auto-filling the Film/Album Title field
   - **Title Prefix**: When checked, prevents auto-filling the Title Prefix field
5. Click "Apply Exception"

### 2. Visual Indicators
- Exception rows are marked with a grey background (`#e9ecef`)
- A grey dot appears next to the checkbox
- A grey left border is added to the row
- Input fields in exception rows have a slightly different styling

### 3. Export Behavior
When exporting data:
- Rows marked as exception will **not** have their Film/Album Title auto-filled (if that exception is enabled)
- Rows marked as exception will **not** have their Title Prefix auto-filled (if that exception is enabled)
- All other rows will continue to use the normal auto-fill behavior based on export settings

## Technical Implementation

### Data Storage
- Exception settings are stored in `localStorage` as `exceptionSettings`
- Format: `{ rowIndex: { filmTitle: boolean, titlePrefix: boolean } }`

### Key Functions
- `markSelectedRowsWithException(color, exceptionConfig)`: Marks rows with exception settings
- `shouldAutoFillField(rowIndex, fieldName)`: Checks if a field should be auto-filled for a specific row
- Export functions in `export_settings.js` check exception settings before applying auto-fill

### CSS Classes
- `.marked-exception`: Applied to exception rows
- `.mark-btn.exception`: Styling for the exception button
- Exception modal styles for the settings interface

## Example Use Cases
1. **Different Film Titles**: When some tracks belong to a different film/album than the main show
2. **Custom Titles**: When certain tracks need custom titles without the series prefix
3. **Mixed Content**: When importing content from multiple sources with different naming conventions

## Removing Exceptions
- Click the "×" button next to the grey dot to remove the exception marking
- Or use the "Mark Row" button and select "Unmark" if available 