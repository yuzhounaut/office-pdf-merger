Bundled runtime components
==========================
qpdf 12.4.1, unchanged official mingw64 build:
https://github.com/qpdf/qpdf/releases/tag/v12.4.1
qpdf.exe, qpdf30.dll: Apache License 2.0; see qpdf-LICENSE.txt and qpdf-NOTICE.md.

The official binary release also supplies:
libgcc_s_seh-1.dll and libstdc++-6.dll: GCC GPLv3 with GCC Runtime Library Exception.
See GCC-GPLv3.txt and GCC-Runtime-Exception.txt.
Source: https://gcc.gnu.org/ and https://github.com/gcc-mirror/gcc

libwinpthread-1.dll: mingw-w64 runtime licensing; see mingw-w64-COPYING.txt.
Source: https://github.com/mingw-w64/mingw-w64

Runtime components are embedded without modification. The build script verifies
the complete official qpdf zip SHA-256 before extracting the executable and DLLs.
