# Angular Web Spreadsheet

A powerful, Excel-like spreadsheet application built with Angular that provides comprehensive spreadsheet functionality in the browser. This web application offers features similar to traditional spreadsheet software while maintaining a modern, responsive interface.

## Features

### Basic Functionality
- Multiple worksheet support with tabbed navigation
- Cell editing and formatting
- Formula support with real-time calculations
- Cut, copy, and paste functionality
- Undo/redo capabilities

### Grid Features
- Row and column resizing
- Drag and drop for rows and columns
- Cell range selection
- Context menu for quick actions
- Column sorting capabilities
- Custom column widths and row heights

### Formula Support
Built-in functions include:
- `SUM(range)`: Calculate sum of cells
- `AVG(range)`: Calculate average of cells
- `MAX(range)`: Find maximum value
- `MIN(range)`: Find minimum value
- `COUNT(range)`: Count non-empty cells
- `TRIM(text)`: Remove extra whitespace
- `UPPER(text)`: Convert text to uppercase
- `LOWER(text)`: Convert text to lowercase
- `REMOVE_DUPLICATES(range)`: Remove duplicate values
- `FIND_AND_REPLACE(range, find, replace)`: Replace text in range

### File Operations
- Create new spreadsheets
- Open existing spreadsheets (XLSX, CSV)
- Save spreadsheets (XLSX, CSV)
- Automatic data persistence using localStorage

## Technical Stack
- Angular (Standalone Components)
- TypeScript
- HyperFormula for formula calculations
- XLSX library for file operations
- Custom formula parser for spreadsheet functions

## Installation

1. Clone the repository:
```bash
git clone [repository-url]
```

2. Install dependencies:
```bash
npm install
```

3. Run the development server:
```bash
ng serve
```

4. Open your browser and navigate to `http://localhost:4200`

## Project Structure

- `grid.component.ts`: Main spreadsheet grid functionality
- `toolbar.component.ts`: Application toolbar and menu actions
- `formula-functions.ts`: Custom formula parser and calculations

## Usage

### Basic Operations
1. **Cell Editing**: Click any cell to begin editing
2. **Formatting**: Use the toolbar buttons for basic text formatting
3. **Formulas**: Start with '=' to enter formula mode
4. **Multi-cell Selection**: Click and drag to select multiple cells

### Formula Examples
```
=SUM(A1:A10)        // Sum values in range A1 to A10
=AVG(B1:B5)         // Average of values in range B1 to B5
=MAX(C1:C20)        // Maximum value in range C1 to C20
```

### Keyboard Shortcuts
- `Ctrl + C`: Copy
- `Ctrl + X`: Cut
- `Ctrl + V`: Paste
- Arrow keys: Navigate between cells
- Enter: Complete cell editing
- F2: Edit cell

## Development

### Adding New Features
1. Create new components in the appropriate directory
2. Use the standalone component architecture
3. Update the toolbar component for new functionality
4. Add new formula functions to the FormulaParser class

### Building for Production
```bash
ng build --prod
```

## Contributing
1. Fork the repository
2. Create a feature branch
3. Commit your changes
4. Push to the branch
5. Create a Pull Request

## License
This project is licensed under the MIT License - see the LICENSE file for details.

## Acknowledgments
- HyperFormula for formula processing
- SheetJS for file handling
- Angular team for the framework

## Support
For support, please open an issue in the repository's issue tracker.
