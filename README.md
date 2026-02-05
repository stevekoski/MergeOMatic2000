# Merge-o-matic 2000 🤖

A retro-techno styled web application for merging and time-aligning CSV and Excel data files.

## Features

- **File Upload**: Drag & drop or browse for CSV, XLS, and XLSX files
- **Flexible Header Detection**: Automatically detects where your data actually starts
- **Column Selection**: Choose which columns to include in the combined output
- **Missing Data Handling**: Multiple strategies for handling gaps in your data
  - Fill with nearest value
  - Linear interpolation
  - Delete rows
  - Fill with zero
- **Duplicate Timestamp Handling**: Average, max, or min when timestamps repeat
- **Time Alignment**: Resample data to consistent intervals (1s to 1 day)
- **Visualization**: Preview your combined data with interactive Plotly charts
- **Excel Output**: Download your merged data as an Excel file

## Usage

### Local Development

Simply open `index.html` in a web browser. For full functionality (especially template loading), you may need to serve the files via a local server:

```bash
# Using Python
python -m http.server 8000

# Using Node.js
npx serve .
```

Then open http://localhost:8000 in your browser.

### GitHub Pages Deployment

1. Push this repository to GitHub
2. Go to Settings → Pages
3. Select "Deploy from a branch" and choose `main` (or your default branch)
4. Your app will be available at `https://yourusername.github.io/repository-name/`

## File Structure

```
/
├── index.html              # Main HTML page
├── css/
│   └── style.css           # Retro techno styling
├── js/
│   ├── app.js              # Main application logic
│   ├── fileHandlers.js     # CSV/Excel parsing
│   └── dataProcessing.js   # Data cleaning & alignment
├── assets/
│   └── Analysis Template.xlsx  # Excel template (optional)
└── README.md
```

## Adding the Excel Template

To use a custom Excel template:

1. Place your `Analysis Template.xlsx` file in the `assets/` folder
2. The application will load it and insert data starting at row 10
3. If no template is found, a basic Excel file will be created

## Libraries Used

- [Papa Parse](https://www.papaparse.com/) - CSV parsing
- [SheetJS](https://sheetjs.com/) - Excel file handling
- [Plotly.js](https://plotly.com/javascript/) - Interactive charting

## Browser Support

Works in all modern browsers (Chrome, Firefox, Safari, Edge).

## License

MIT
