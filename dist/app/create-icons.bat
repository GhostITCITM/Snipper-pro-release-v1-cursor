@echo off
echo Creating placeholder icons for Excel add-in...

REM Create simple colored squares as placeholder icons
REM In production, replace these with proper icon files

echo ^<svg width="16" height="16" xmlns="http://www.w3.org/2000/svg"^>^<rect width="16" height="16" fill="#0078d4"/^>^<text x="8" y="12" text-anchor="middle" fill="white" font-size="10"^>S^</text^>^</svg^> > 16.svg
echo ^<svg width="32" height="32" xmlns="http://www.w3.org/2000/svg"^>^<rect width="32" height="32" fill="#0078d4"/^>^<text x="16" y="22" text-anchor="middle" fill="white" font-size="18"^>S^</text^>^</svg^> > 32.svg
echo ^<svg width="64" height="64" xmlns="http://www.w3.org/2000/svg"^>^<rect width="64" height="64" fill="#0078d4"/^>^<text x="32" y="42" text-anchor="middle" fill="white" font-size="36"^>S^</text^>^</svg^> > 64.svg
echo ^<svg width="80" height="80" xmlns="http://www.w3.org/2000/svg"^>^<rect width="80" height="80" fill="#0078d4"/^>^<text x="40" y="52" text-anchor="middle" fill="white" font-size="44"^>S^</text^>^</svg^> > 80.svg

echo Placeholder SVG icons created. Convert to PNG for production use.
echo Install ImageMagick and run: magick convert 16.svg 16.png
pause