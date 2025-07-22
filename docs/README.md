# 📍 Excel Address Geocoder

A powerful web application that transforms address data from Excel files into precise GPS coordinates using Google Maps API.

## ✨ Features

- **🔄 Drag & Drop Upload**: Easy Excel file uploading
- **👁️ Data Preview**: See your data before processing
- **🎯 Smart Column Selection**: Choose which column contains addresses
- **🌍 Real-time Geocoding**: Convert addresses to latitude/longitude coordinates
- **📊 Results Preview**: View processed data with coordinates
- **⬇️ Instant Download**: Get your geocoded Excel file immediately
- **🔒 Secure**: API key is never stored or shared

## 🚀 Live Demo

**Try it now:** [https://yourusername.github.io/excel-address-geocoder](https://yourusername.github.io/excel-address-geocoder)

## 📋 How to Use

1. **Get Google Maps API Key**
   - Go to [Google Cloud Console](https://console.cloud.google.com/)
   - Enable Geocoding API
   - Create an API key

2. **Upload Your Excel File**
   - Drag & drop or browse for your `.xlsx` file
   - Maximum file size: 200MB

3. **Process Your Data**
   - Enter your Google Maps API key
   - Select the column containing addresses
   - Click "Fetch Coordinates"

4. **Download Results**
   - Preview your geocoded data
   - Download the Excel file with latitude/longitude columns added

## 🛠️ Technical Details

- **Frontend**: Vanilla HTML, CSS, JavaScript
- **Excel Processing**: SheetJS library
- **API**: Google Maps Geocoding API
- **Hosting**: GitHub Pages
- **No Backend Required**: Everything runs in your browser

## 📁 File Structure

```
excel-address-geocoder/
├── index.html          # Main application file
├── README.md          # This file
├── LICENSE           # MIT License
└── .gitignore       # Git ignore rules
```

## 🔐 Privacy & Security

- Your data never leaves your browser
- API key is only used for geocoding requests
- No data is stored on any servers
- Fully client-side processing

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## 📜 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🆘 Support

If you encounter any issues or have questions:
- Open an [Issue](https://github.com/yourusername/excel-address-geocoder/issues)
- Check the [Wiki](https://github.com/yourusername/excel-address-geocoder/wiki)

## ⭐ Show Your Support

If this project helped you, please give it a ⭐ star on GitHub!

---

**Made with ❤️ for the developer community**
