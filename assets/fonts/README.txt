Place font files here to bundle with the Windows build.

Recommended:
- NotoSansKhmer-Regular.ttf (OFL license)

The GitHub Actions workflow bundles this directory via PyInstaller:
--add-data "assets\fonts;assets\fonts"
