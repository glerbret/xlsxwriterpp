# Xlsxwriter++

> [!IMPORTANT]
> Xlsxwriter++ is still a work in progress.
>
> For the moment, it's a partial, direct, quick and dirty C++ port.
> No idiomatic C++, no rich C++ API, ...
>
> Future commit will change architecture, implementation and API without backward compatibility.

Xlsxwriter++ is a C++20 library for writing Excel XLSX files.

This is a C++ port of [libxlsxwriter](https://libxlsxwriter.github.io/).

## Requirement

- C++20
- ZLIB
- Openssl
- Boost.Test for unit tests

## Limitations

- No support of Windows 32bits.
- No in-memory generation of XLSX file.
