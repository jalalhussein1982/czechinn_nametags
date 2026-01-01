# Hotel Nametag Generator

Generates printable nametags from hotel arrival PDF lists.

## Layout (per nametag)
| Row | Content |
|-----|---------|
| 1 | CZECHINN (18pt, Geo Sans Light) | # room_number (16pt, bold) |
| 2 | Reservation name: (10pt) LASTNAME (12pt, bold) |
| 3 | [QR Code] | From: XX To: XX / Check-out: 12:00 / wifi: iczechedinn |
| 4 | ID value (e.g., 000) |

## Features
- QR code generated dynamically from WiFi password
- 12 nametags per A4 page (2×6 grid)
- No external image assets required

## Installation
```bash
pip install pdfplumber python-docx qrcode pillow
```

## Usage
**GUI Mode (Windows/Mac):**
```bash
python hotel_nametag_generator.py
```

**CLI Mode:**
```bash
python hotel_nametag_generator.py input.pdf output.docx
```
