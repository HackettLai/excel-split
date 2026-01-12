# 📊 Data Split Tool

A lightweight web-based tool to split multi-value cells in Excel/CSV files into separate rows. Perfect for cleaning data where multiple values are stored in a single cell.

![Version](https://img.shields.io/badge/version-1.2.1-blue)
![License](https://img.shields.io/badge/license-MIT-green)

## 🎯 Features

- **📁 Multi-format Support**: Works with Excel (.xlsx, .xls) and CSV files
- **✂️ Smart Splitting**: Automatically detects common separators (commas, semicolons, slashes, etc.)
- **🎨 Visual Feedback**: Highlights split rows for easy identification
- **🌐 Multi-language**: Supports both Western and Chinese separators (`,` `，` `;` `；` `/` `／`)
- **🛡️ Special Term Protection**: Preserves terms like "N/A" without splitting
- **💾 Excel Export**: Download results with formatting and highlighting preserved
- **📱 Responsive Design**: Works on desktop and mobile devices
- **🚀 No Backend Required**: Runs entirely in your browser

## 🚀 Demo

[[Live Demo]](https://excel-split.hackettlai.com)

## 📸 Screenshots

### Upload & Configure
Upload your file and select which column to split using either column names or index numbers.

### Results Preview
View split data with visual highlighting showing which rows were created from multi-value cells.

### Download
Export your cleaned data with formatting preserved.

## 🔧 How It Works

1. **Upload** your Excel or CSV file
2. **Configure** header settings (check if first row contains headers)
3. **Select** the column to split by name or index number
4. **Process** - the tool automatically detects separators and splits values
5. **Download** your cleaned data as an Excel file

### Supported Separators

The tool automatically detects and splits on these separators:
- Comma: `,` `，`
- Semicolon: `;` `；`
- Forward slash: `/` `／`
- Ampersand: `&`
- Newline: `\n`

### Protected Terms

These terms are **NOT** split even if they contain separators:
- `N/A` `n/a` `N\A` `n\a`

## 💡 Use Cases

- **Email Lists**: Split cells like `john@email.com, jane@email.com` into separate rows
- **Tags/Categories**: Separate multiple tags stored in one cell
- **Multi-value Fields**: Clean up any data where multiple values were incorrectly stored together
- **Data Migration**: Prepare data for import into systems that require one value per row
- **Survey Data**: Split multiple-choice responses into individual entries

## 🛠️ Installation

### Option 1: Use Online (Recommended)
Simply visit the [live demo](https://excel-split.hackettlai.com) - no installation needed!

### Option 2: Run Locally

1. Clone this repository:
```bash
git clone https://github.com/yourusername/data-split-tool.git
cd data-split-tool
```

2. Open `index.html` in your browser:
```bash
# On macOS
open index.html

# On Linux
xdg-open index.html

# On Windows
start index.html
```

That's it! No build process or dependencies required.

## 📦 Dependencies

The tool uses these CDN-hosted libraries (no installation required):
- [SheetJS (xlsx)](https://cdn.sheetjs.com/) - Excel file processing
- Modern browser with JavaScript enabled

## 🎨 Customization

### Color Scheme
Edit the CSS variables in `style.css`:
```css
:root {
  --gradient: linear-gradient(to right, #948e99, #2e1437);
  --btn: linear-gradient(315deg, #948e99, #2e1437);
  --colorAccent: #6f5a89;
  --text: #6b5e71;
}
```

### Separators
Add or modify separators in `split.js`:
```javascript
const separators = [',', ';', '/', '&', '\n', '，', '；', '／'];
```

### Protected Terms
Edit the special terms list in `split.js`:
```javascript
const specialTerms = ['n/a', 'N/A', 'n\\a', 'N\\A'];
```

## 🤝 Contributing

Contributions are welcome! Here's how you can help:

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

### Ideas for Contributions
- [ ] Add support for custom separators
- [ ] Support for multiple column splitting
- [ ] Preview before processing
- [ ] Undo/redo functionality
- [ ] Dark mode
- [ ] Additional export formats (CSV, JSON)

## 📝 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 👤 Author

**Hackett Lai**

- GitHub: [@HackettLai](https://github.com/HackettLai)
- Website: [Your Website](https://hackettlai.com)

## 🙏 Acknowledgments

- Built with [SheetJS](https://sheetjs.com/) for Excel processing
- Inspired by the need for simple, browser-based data cleaning tools
- Thanks to all contributors and users who provide feedback!

## 📊 Version History

### v1.2.1 (Current)
- ✨ Added clickable column badges for easy selection
- 🛡️ Protected special terms (N/A) from splitting
- 🎨 Improved visual feedback with success messages
- 📱 Enhanced mobile responsiveness

### v1.2.0
- ✨ Added header checkbox option
- 🔧 Support for both column name and index input
- 🎨 Visual column list display

### v1.1.0
- ✨ Multi-language separator support
- 💾 Excel export with formatting

### v1.0.0
- 🎉 Initial release
- ✂️ Basic split functionality

## 🐛 Known Issues

- Very large files (>5MB) may take longer to process
- Some special characters in headers may need manual encoding

## 📮 Support

Having issues? Please [open an issue](https://github.com/HackettLai/data-split-tool/issues) on GitHub.

---

<div align="center">
  Made with ❤️ by Hackett Lai
  <br>
  ⭐ Star this repo if you find it useful!
</div>
```
