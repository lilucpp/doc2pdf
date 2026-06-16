# doc2pdf

> A command-line tool to convert WPS / Microsoft Office documents (.doc/.docx) to PDF.

[![Platform](https://img.shields.io/badge/Platform-Windows-blue)](https://github.com/lilucpp/doc2pdf)
[![Language](https://img.shields.io/badge/Language-C++11-00599C)](https://isocpp.org/)
[![Build](https://img.shields.io/badge/Build-CMake_%2F_MSVC-6DB33F)](https://cmake.org/)
[![License](https://img.shields.io/badge/License-MIT-green)](LICENSE)

## Features

- **No printer driver required** — Uses COM automation to call the native export API of WPS / Office
- **Auto-detection** — Prefers WPS, falls back to Microsoft Word
- **Page range support** — Specify start/end page numbers
- **Lightweight** — Single exe, no external runtime dependencies (static link /MT)
- **32/64-bit compatible** — 32-bit build recommended for maximum WPS/Office compatibility

## How It Works

Uses **COM automation** to invoke the application's `ExportAsFixedFormat` API, no middleware or virtual printer needed.

## Compatibility

| Application | ProgID | Requirements |
|-------------|--------|-------------|
| **WPS Office** (recommended) | `KWps.Application` | WPS Office installed |
| **Microsoft Office Word 2007** | `Word.Application` | Requires **SP2** or [Save as PDF add-in](http://www.microsoft.com/downloads/details.aspx?familyid=4d951911-3e7e-4ae6-b059-a2e79ed87041) |
| **Microsoft Office Word 2010+** | `Word.Application` | Native support, no additional components |
| **Microsoft Office 365** | `Word.Application` | Native support |

> Startup order: tries WPS first → falls back to Office Word → exits with error if both fail.

## Usage

```
doc2pdf <input file> [-o output file] [-f from page] [-t to page]
```

| Argument | Description |
|----------|-------------|
| `-o` / `--output` | Output PDF path (default: input file name with .pdf extension) |
| `-f` / `--from` | Start page (default: page 1) |
| `-t` / `--to` | End page (default: last page; requires `-f`) |

### Examples

```powershell
doc2pdf 1.docx                      # All pages
doc2pdf 1.docx -f 2 -t 5           # Pages 2-5
doc2pdf 1.docx -f 3                # From page 3 to end
doc2pdf 1.docx -o D:\out.pdf       # Specify output path
```

## Build

### Requirements

- Visual Studio 2022 Build Tools (or full VS)
- CMake 3.10+

### Build (32-bit, static link /MT)

```powershell
cmake -G "Visual Studio 17 2022" -A Win32 -B build
cmake --build build --config Release
```

Output at `build\Release\doc2pdf.exe`, no VC runtime dependencies.

### Architecture Notes

The tool is compiled as 32-bit and uses COM to call WPS/Office. A 32-bit tool can only call 32-bit WPS/Office (the most common installation on 64-bit systems).

To support native 64-bit Office, also build a 64-bit version with `-A x64`:

```powershell
cmake -G "Visual Studio 17 2022" -A x64 -B build64
cmake --build build64 --config Release
```

## Keywords

`doc2pdf` `docx2pdf` `word to pdf` `wps to pdf` `office to pdf` `command line pdf converter`

## License

[MIT](LICENSE)
