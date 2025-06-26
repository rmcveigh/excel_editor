# Excel Editor for Drupal

A comprehensive Drupal module that provides advanced Excel file editing capabilities directly in the browser.

## Features

### Core Functionality
- **File Upload**: Support for Excel (.xlsx, .xls) and CSV files up to 10MB
- **Interactive Editing**: In-browser editing of Excel data with real-time updates
- **Column Management**: Show/hide columns, set default visibility
- **Advanced Filtering**: Multi-column filtering with various operators
- **Row Selection**: Select individual rows or all visible rows for bulk operations
- **Draft System**: Save and restore work sessions with auto-save capability
- **Export Options**: Export selected rows or all data back to Excel format

### User Interface
- **Modern Design**: Clean, responsive interface using Bulma CSS framework
- **Drag & Drop**: Intuitive file upload with drag-and-drop support
- **Real-time Feedback**: Loading indicators and status messages
- **Keyboard Shortcuts**: Ctrl+S to save drafts

### Technical Features
- **Security**: CSRF protection and proper user permissions
- **Performance**: Optimized for large datasets with column visibility limits
- **Accessibility**: Semantic markup and proper contrast ratios
- **Browser Support**: Works with modern browsers supporting ES6+

## Installation

### Requirements
- Drupal 9 or 10
- PHP 8.0+
- Modern web browser with JavaScript enabled

### Steps
1. Download and extract the module to your Drupal modules directory:
   ```bash
   cd /path/to/drupal/modules/custom
   git clone [your-repo-url] excel_editor
   ```

2. Enable the module:
   ```bash
   drush en excel_editor
   ```
   Or via the Drupal admin interface at `/admin/modules`

3. Set permissions at `/admin/people/permissions`:
   - `use excel editor` - For users who should access the editor
   - `administer excel editor` - For administrators who can configure settings

4. Configure the module at `/admin/config/content/excel-editor`

## Usage

### Basic Workflow
1. Navigate to `/excel-editor`
2. Upload an Excel file by dragging/dropping or clicking "Choose File"
3. Edit data directly in the table interface
4. Use filters to find specific data
5. Select rows for bulk operations
6. Save your work as a draft or export the results

### Editable Columns
The module automatically adds three editable columns to any uploaded file:
- **new_barcode**: For entering new barcode values
- **notes**: For adding comments or notes
- **actions**: Dropdown for workflow actions (Relabel, Pending, Discard)

### Filtering
Click the "Filter" link in any column header to:
- Select specific values to show/hide
- Use text search within the filter
- Apply multiple filters simultaneously
- Clear individual or all filters

### Column Visibility
Use the "Show/Hide Columns" button to:
- Hide irrelevant columns for focused work
- Show only editable columns
- Apply default visibility settings
- Optimize performance for large datasets

### Draft Management
- **Auto-save**: Drafts are automatically saved every 5 minutes (if enabled)
- **Manual save**: Use Ctrl+S or the "Save Draft" button
- **Load drafts**: Access saved drafts from the sidebar
- **Persistent state**: Filters, selections, and column visibility are preserved

## Configuration

Access configuration at `/admin/config/content/excel-editor`:

### General Settings
- **Enable Autosave**: Automatically save drafts every 5 minutes

### Column Visibility
- **Default Visible Columns**: Specify which columns should be visible by default
- **Column Hiding Behavior**: Choose whether to hide non-specified columns
- **Always Visible Columns**: Columns that are always shown regardless of settings

### Performance
- **Maximum Visible Columns**: Limit concurrent visible columns for better performance

## API Endpoints

The module provides REST-like endpoints for AJAX operations:

- `POST /excel-editor/save-draft` - Save or update a draft
- `GET /excel-editor/load-draft/{id}` - Load a specific draft
- `DELETE /excel-editor/delete-draft/{id}` - Delete a draft
- `GET /excel-editor/drafts` - List user's drafts

## Development

### File Structure
```
excel_editor/
├── config/
│   ├── install/excel_editor.settings.yml
│   └── schema/excel_editor.schema.yml
├── css/excel-editor.css
├── js/excel-editor.js
├── src/
│   ├── Controller/ExcelEditorController.php
│   ├── DraftManager.php
│   └── Form/ExcelEditorConfigForm.php
├── templates/
│   ├── components/
│   └── excel-editor-page.html.twig
├── excel_editor.info.yml
├── excel_editor.install
├── excel_editor.libraries.yml
├── excel_editor.module
├── excel_editor.permissions.yml
├── excel_editor.routing.yml
└── excel_editor.services.yml
```

### Key Components
- **ExcelEditor JavaScript Class**: Main frontend application
- **DraftManager Service**: Database operations for draft management
- **ExcelEditorController**: REST API endpoints
- **ExcelEditorConfigForm**: Admin configuration interface

### Dependencies
- **SheetJS**: Excel file parsing and generation
- **Bulma CSS**: UI framework
- **FontAwesome**: Icons
- **jQuery**: DOM manipulation and AJAX

## Security

- All API endpoints require proper user permissions
- CSRF tokens protect state-changing operations
- User isolation ensures drafts are only accessible by their owners
- Input validation prevents XSS and injection attacks
- File upload restrictions limit size and format

## Browser Compatibility

- Chrome 60+
- Firefox 60+
- Safari 12+
- Edge 79+

## Troubleshooting

### Common Issues

**"Failed to load XLSX library"**
- Check your internet connection for CDN access
- Verify Bulma and FontAwesome CSS are loading

**"No data found in file"**
- Ensure the Excel file contains actual data
- Check that the first row contains headers
- Verify file is not corrupted

**"Permission denied"**
- Check user permissions at `/admin/people/permissions`
- Ensure user has "use excel editor" permission

**Poor performance with large files**
- Adjust "Maximum Visible Columns" in settings
- Hide unnecessary columns
- Consider splitting large files

### Debug Mode
Add `?debug=1` to the URL or set `localStorage.excel_editor_debug = 'true'` in browser console to enable debug logging.

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes with appropriate tests
4. Submit a pull request

## License

This module is licensed under the GPL-2.0+ license, consistent with Drupal core.

## Support

For issues, feature requests, or questions:
- Create an issue in this repository
- Check the Drupal.org project page (if published)
- Review the module's help page at `/admin/help/excel_editor`
