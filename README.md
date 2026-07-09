# ABSA Converter Tool

A Windows desktop utility that converts proprietary card-data (`.inp`) files into Microsoft Access database (`.mdb`) files, one output database per input file.

## Overview

The tool reads every file in a chosen folder, splits each one into individual card records, extracts fields such as PAN, expiry date, cardholder name, CVV2, track 1/2 data, and EMV chip data, and writes the results into a Microsoft Access `Cards` table. Each input file produces its own `<InputFileName>_Output.mdb` file in the same folder.

## Features

- Batch processing of an entire folder in a single run
- Per-file Access database output (no merging of records across files)
- Chip data (binary) and text fields decoded from the same source record
- Per-file/per-record error reporting that does not halt the rest of the batch
- Simple two-step workflow: **Choose Folder** → **Process Files**

## Requirements

- Windows 10/11
- [.NET 8 SDK](https://dotnet.microsoft.com/download) (target framework: `net8.0-windows`)
- Visual Studio 2022 (or another install that provides the full-framework MSBuild) — the project uses a COM reference (`ADOX`) that the `dotnet build` CLI cannot resolve; building/running requires the Visual Studio version of MSBuild
- Microsoft Access Database Engine, **32-bit (x86)**, providing the `Microsoft.ACE.OLEDB.12.0` provider — required to create `.accdb`/`.mdb` files. The project is built for the `x86` platform to match this provider's bitness.

## Building

Open `Absa.sln` in Visual Studio 2022 and build/run from there (recommended), or use the full-framework MSBuild directly:

```powershell
& "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" Absa.sln -p:Configuration=Debug
```

> Running `dotnet build Absa.sln` will fail with `MSB4803: The task "ResolveComReference" is not supported on the .NET Core version of MSBuild` — this is expected due to the ADOX COM reference.

## Usage

1. Launch **AbsaConverterTool**.
2. Click **Choose Folder** and select a folder containing the input files. The tool parses every file in that folder and reports:
   - Successfully parsed records in the log window
   - Any files/records that failed to parse, listed separately
3. Click **Process Files**. For each successfully parsed input file, the tool writes an Access database named `<InputFileName>_Output.mdb` into the same folder, containing that file's records in a `Cards` table.
4. Review the summary in the log window — each input file is reported as **Created**, **Skipped** (no valid records parsed), or **Failed** (with the underlying error).

### Input file format

Input files contain one or more records separated by an `#END#` delimiter. The first record in each file is treated as a header and discarded. Within each record:

| Field | Extracted from |
|---|---|
| PAN | Text between `*` and `@` |
| Expiry (`IDWEXP`) | Fixed-length text after `$`, offset 13, length 5 |
| Name | Text between `) ` and `@` |
| CVV2 | Text between `:` and `@` |
| Track 1 / Track 2 | Text between `"`/`@`, then split on `%`/`?` and `;`/`?` |
| Chip data (`IDWChip1`–`IDWChip4`) | Binary block following `{` + a 7-digit length prefix, decoded to hex and split into 4 chunks |

### Output schema

Each generated `.mdb` file contains a single `Cards` table with the columns: `IDWAutoNumber`, `JobNumber`, `IDWPAN`, `IDWEXP`, `IDWNAME`, `IDWCVV2`, `IDWTrack1`, `IDWTrack2`, `IDWChip1`, `IDWChip2`, `IDWChip3`, `IDWChip4`.

## Project structure

```
AbsaConverterTool/
├── MainWindow.xaml / .cs   UI and workflow orchestration
├── Helper/FileHelper.cs    File parsing, field extraction, and Access DB creation
└── App.xaml / .cs          Application entry point
```

## Known limitations

- The input folder scan reads *every* file in the folder (no extension filter). Re-running **Choose Folder** on a folder that already contains previously generated `*_Output.mdb` files will attempt to parse those too; they simply fail with a per-file parse error and do not affect the rest of the batch.
- If two input files share the same base name (e.g. `Card001.inp` and `Card001.dat`), the tool automatically appends a numeric suffix (`Card001_Output_2.mdb`) to avoid overwriting the first file's output.
