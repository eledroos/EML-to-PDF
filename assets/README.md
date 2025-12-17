# Application Icons

This folder should contain the application icons for each platform:

- `icon.icns` - macOS icon (1024x1024, Apple Icon Image format)
- `icon.ico` - Windows icon (256x256 multi-resolution)
- `icon.png` - Linux/general icon (512x512 or 1024x1024 PNG)

## Creating Icons

### Option 1: Use the generate script

Run the icon generator script (requires Pillow):

```bash
pip install pillow
python assets/generate_icons.py
```

### Option 2: Create manually

1. Create a square image (1024x1024 recommended) with your design
2. Use online tools or image editors to convert:
   - For macOS: Use `iconutil` or online .icns converter
   - For Windows: Use online .ico converter (include 16, 32, 48, 256 sizes)
   - For Linux: Save as PNG at 512x512 or 1024x1024

### Suggested Design

For an EML to PDF converter, consider:
- An envelope icon transforming into a PDF document
- Email icon with PDF badge
- Simple document with @ symbol and PDF indicator

### Note

If icons are missing, the build will proceed without custom icons (using system defaults).
