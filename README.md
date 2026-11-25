# Лист Таблицы (List Spreadsheet)

An Android spreadsheet editor built with Python and Kivy. The project is in early development.

## Current prototype
- Spreadsheet grid with selectable cells.
- Enter values and formulas.
- Formula engine: arithmetic, cell references, ranges, SUM, AVERAGE, MIN, MAX, COUNT, COUNTA, IF, AND, OR, NOT, ABS, ROUND, SQRT, LEN, UPPER, LOWER and CONCATENATE.
- Basic XLSX import and export using OpenPyXL.
- GitHub Actions workflow for building a Debug APK.

## Build locally
Requires Python 3.11, Java 17, Android SDK/NDK and Buildozer 1.5.0.

```bash
python -m pip install -r requirements.txt
pip install buildozer==1.5.0 cython==0.29.36
buildozer android debug
```

The APK is generated in `bin/`.

## Build with GitHub Actions
Open the **Actions** tab and run **Android Debug APK** with **Run workflow**. Pushes affecting Python sources or Android build configuration also trigger a build. Download the APK from the completed workflow run's Artifacts section.

## Documents
- [Privacy policy](PRIVACY_POLICY.md)
- [Terms of use](TERMS_OF_USE.md)
- [MIT License](LICENSE)

## Project status
This is an initial prototype, not an Excel replacement yet. Formula compatibility, formatting, charts, sorting, filtering, CSV support, file picker integration and broader XLSX compatibility are planned. OpenPyXL reads and writes formulas but does not calculate them; the app's own formula engine evaluates the supported subset.

## Contributions
Bug reports and feature suggestions are welcome through GitHub Issues.
