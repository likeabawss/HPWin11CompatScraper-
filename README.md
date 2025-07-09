# HPWin11CompatScraper

A simple PowerShell script that scrapes the list of HP business PCs tested for Windows 11 from HP’s official support site and outputs the results as a structured JSON file.

## Features
- **Automatic Chrome/Chromedriver Setup:** Handles Chrome and ChromeDriver installation and version matching automatically.
- **Robust Scraping:** Uses Selenium to load dynamic content and save the full HTML page source.
- **HTML Extraction:** Parses the saved HTML to extract compatibility tables and outputs a clean JSON file.
- **No CSV Output:** Only JSON output is produced for easy integration with other tools.

## Requirements
- Windows 10/11
- PowerShell 5.1 or later
- Internet connection
- [Selenium PowerShell module](https://www.powershellgallery.com/packages/Selenium/)
- Google Chrome (the script will attempt to install if not found)

## Usage
1. Open PowerShell as Administrator (recommended for first run).
2. Clone or download this repository.
3. Run the script:

    ```powershell
    cd path\to\HPWin11CompatScraper-
    # Run with default options
    powershell -ExecutionPolicy Bypass -File .\Get-HPWin11CompatitibilityExtract.ps1

    # Optional parameters:
    # -ShowProgress: Show detailed scraping progress
    # -HeadlessBrowser: Run Chrome in headless mode
    # -WaitTimeSeconds: Adjust wait time for page load (default 15)
    powershell -ExecutionPolicy Bypass -File .\Get-HPWin11CompatitibilityExtract.ps1 -ShowProgress -HeadlessBrowser -WaitTimeSeconds 20
    ```

## Output
- `selenium_page_source.html`: The raw HTML page source as loaded by Selenium.
- `HP_Win11_Compatibility.json`: Structured JSON file containing the extracted compatibility data.

## Example JSON Output
```json
{
  "ExtractedDate": "2025-07-10 14:23:01",
  "SourceUrl": "https://support.hp.com/us-en/document/ish_4890350-4890415-16",
  "TableTitle": "Business PCs and workstations tested with Windows 11",
  "Headers": ["Model", "Product Number", "Windows 11 Support"],
  "DataCount": 123,
  "Data": [
    {
      "Model": "HP EliteBook 840 G8",
      "Product Number": "1234ABCD",
      "Windows 11 Support": "Yes"
    }
    // ... more entries ...
  ]
}
```

## Troubleshooting
- If Chrome is not installed, the script will attempt to download and install it.
- If you encounter issues, try running with `-ShowProgress` for more detailed output.
- Increase `-WaitTimeSeconds` if the page is slow to load.
- Check `selenium_page_source.html` for the actual HTML captured.


