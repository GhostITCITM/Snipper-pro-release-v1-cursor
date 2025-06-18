# Third-Party Licenses

This project depends on several open source components. Their original licenses apply. No license text has been modified. Copies of the upstream license texts are linked below.

| Component | Version | License | Source |
|-----------|---------|---------|--------|
| PdfiumViewer | 2.13.0 | Apache-2.0 | https://github.com/pvginkel/PdfiumViewer |
| PdfiumViewer.Native (x86/x64) | 2018.4.8.256 | BSD-3-Clause (PDFium) | https://github.com/pvginkel/PdfiumViewer |
| Tesseract (C# wrapper) | 5.2.0 | Apache-2.0 | https://github.com/charlesw/tesseract |
| Newtonsoft.Json | 13.0.3 | MIT | https://github.com/JamesNK/Newtonsoft.Json |
| Microsoft.Office.Interop.Excel | 15.0.4795.1001 | Proprietary (Microsoft) | https://learn.microsoft.com/ |

## Attribution

- **Apache License 2.0** components include PdfiumViewer and Tesseract. Their license terms require preservation of the license notice and inclusion of a copy of the license when distributing binaries or source. This project includes unmodified copies of these libraries via NuGet.
- **MIT License** component is Newtonsoft.Json. The MIT License requires inclusion of the copyright notice and license when distributing. A copy of the license text is included in this repository via the link above.
- **BSD-3-Clause** code from PDFium is bundled as native binaries via PdfiumViewer.Native packages. The license notice from Google is preserved.

No third-party code listed above has been modified. All libraries are consumed through official NuGet packages. Redistributing compiled binaries of this project should include this file and copies of the licenses to satisfy the requirements of the respective licenses.


