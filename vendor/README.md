# Vendor Dependencies

This folder contains third-party binary dependencies required to run and package the application.

## qpdf

- **Version**: 12.4.1 (mingw64 official release)
- **Archive filename**: `qpdf-12.4.1-mingw64.zip`
- **Expected SHA-256**: `6a47eeddc8ff712a6e003314daae25569402c6e904ba82b7b9181d7b0301b689`
- **Official Download**: [https://github.com/qpdf/qpdf/releases/tag/v12.4.1](https://github.com/qpdf/qpdf/releases/tag/v12.4.1)

If the zip file is not found locally, `build.ps1` will automatically download and verify it from GitHub releases.
For offline builds, download `qpdf-12.4.1-mingw64.zip` manually and place it in this directory.
