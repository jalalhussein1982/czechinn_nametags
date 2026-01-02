# Hotel Nametag Generator

Automated nametag generator for Czech Inn hostel. Extracts guest data from PDF arrival reports and generates printable Word documents with 12 nametags per A4 page (6 rows x 2 columns).

## Features

- **PDF Parsing**: Automatically extracts guest information from hotel arrival PDF reports
- **Smart Multi-Guest Handling**: Identifies rooms with multiple guests and lets you choose how many nametags to print
- **Sequential ID Generation**: Assigns sequential IDs (000, 001, 002...) without gaps in the final output
- **Matchcode Display**: Shows booking reference codes for easy identification during review
- **WiFi QR Code**: Automatically generates QR codes for WiFi connection on each nametag
- **GDPR Compliant**: Web app processes data in memory only; all temporary files are deleted when session ends
- **12 Nametags per Page**: Optimized A4 layout (2 columns x 6 rows)
- **Empty Nametag Templates**: Fills remaining cells on the last page with blank templates for manual use

## Installation

### Prerequisites
- Python 3.7 or higher

### Install Dependencies

```bash
pip install pdfplumber python-docx qrcode pillow streamlit
```

Or use the requirements file:

```bash
pip install -r requirements.txt
```

## Usage

### Web Application (Recommended)

The Streamlit web app provides the easiest user experience with a graphical interface.

```bash
streamlit run streamlit_app.py
```

This will open the app in your default web browser.

**Steps:**
1. Upload your hotel arrivals PDF file
2. Click "Parse PDF" to extract guest data
3. For rooms with multiple guests, select how many nametags to print for each
4. Click "Generate Nametags" to create the Word document
5. Download the generated DOCX file
6. Click "End Session" when done to clear all data (GDPR compliance)

### GUI Mode (Desktop)

For Windows/Mac users who prefer a native desktop experience:

```bash
python hotel_nametag_generator.py
```

This opens file dialogs to select input PDF and output location.

### Command Line Mode

For batch processing or automation:

```bash
python hotel_nametag_generator.py input.pdf output.docx
```

## Nametag Layout

Each nametag contains the following information:

| Row | Content |
|-----|---------|
| 1 | CZECHINN (18pt) + # ______ (room number blank for manual fill) |
| 2 | Reservation name: LASTNAME (bold) |
| 3 | QR Code (left) + From/To dates, Check-out time, WiFi code (right) |
| 4 | Sequential ID number (e.g., 000, 001, 002) |

### Layout Specifications
- **Page Size**: A4 (21cm x 29.7cm)
- **Nametags per Page**: 12 (2 columns x 6 rows)
- **Cell Dimensions**: 10.5cm x 4.80cm
- **Borders**: Dashed lines for easy cutting

## PDF Input Format

The application expects a PDF with a table containing these columns:
1. **Matchcode/Reg.No.** - Format: `LASTNAME_FIRSTNAME_BOOKINGCODE` (e.g., `AHMAD_WASEEM_XXD/2521/1`)
2. **Room** - Room number with guest count (e.g., `102(4)`, `501-A(2)`, `-1(1)`)
3. **Arrival** - Arrival date (DD.MM.YY format)
4. **Departure** - Departure date (DD.MM.YY format)
5. **Guests** - Number of guests in the reservation

## Configuration

### WiFi Settings

To change WiFi credentials, edit these values in `hotel_nametag_generator.py`:

```python
WIFI_SSID = "Czech-Inn-Public-NEW"
WIFI_PASSWORD = "iczechedinn"
```

### Nametag Height

To adjust nametag height, modify:

```python
ROW_HEIGHT_CM = 4.80  # Height per row in centimeters
```

## File Structure

```
hotel_nametag_generator/
├── hotel_nametag_generator.py  # Main script (GUI + CLI)
├── streamlit_app.py            # Web application
├── requirements.txt            # Python dependencies
├── README.md                   # This file
├── CLAUDE.md                   # Developer reference
└── sample_output.docx          # Example output
```

## Troubleshooting

### "GuestRecord object has no attribute" error
Restart the Streamlit server to reload the updated modules:
1. Press `Ctrl+C` in the terminal
2. Run `streamlit run streamlit_app.py` again

### QR code not appearing
Ensure `pillow` is installed: `pip install pillow`

### PDF parsing fails
- Verify the PDF contains a properly formatted table
- Check that the Matchcode column follows the expected format

### Content clipped in nametags
Adjust `ROW_HEIGHT_CM` in the generator to increase cell height.

## License

MIT License

Copyright (c) 2024 Jalal Hussein

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

## Author

**Jalal Hussein**

- Email: jalalhussein@gmail.com
- GitHub: [https://github.com/jalalhussein1982](https://github.com/jalalhussein1982)

## Acknowledgments

- Built for Czech Inn Hostel, Prague
- Uses [pdfplumber](https://github.com/jsvine/pdfplumber) for PDF extraction
- Uses [python-docx](https://python-docx.readthedocs.io/) for Word document generation
- Uses [qrcode](https://github.com/lincolnloop/python-qrcode) for WiFi QR code generation
- Web interface powered by [Streamlit](https://streamlit.io/)
