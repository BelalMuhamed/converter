# ABSA Converter Tool (v3)

A Windows desktop utility that converts proprietary card-data (`.inp`) files into Microsoft Access database (`.mdb`) files, split **per product**, one output database per product found in each input file.

## Overview

The tool reads every file in a chosen folder, splits each one into individual card records, and for each record extracts the Product Code, PAN, expiry date, cardholder name, CVV2, track 1/2 data, and EMV chip data. Records are grouped by product and written into per-product `Cards` tables. Records that cannot be assigned to a known product, or whose chip data is too large for the schema, are **not** written to any `.mdb` — they are collected and reported in a `Failures.xlsx` file instead. Processing never stops because of one bad file or record.

## Features

- Batch processing of an entire folder in a single run
- **Product-based output splitting**: one `.mdb` per product found in each input file (a file containing 3 products produces 3 `.mdb` files)
- Chip data (binary) and text fields decoded from the same source record
- Per-file/per-record error reporting that does not halt the rest of the batch
- Invalid/unmappable records are captured in an `Failures.xlsx` report instead of being silently dropped
- Simple two-step workflow: **Choose Folder** → **Process Files**

## Requirements

- Windows 10/11
- [.NET 8 SDK](https://dotnet.microsoft.com/download) (target framework: `net8.0-windows`)
- Visual Studio 2022 (or another install that provides the full-framework MSBuild) — the project uses a COM reference (`ADOX`) that the `dotnet build` CLI cannot resolve; building/running requires the Visual Studio version of MSBuild
- Microsoft Access Database Engine, **32-bit (x86)**, providing the `Microsoft.ACE.OLEDB.12.0` provider — required to create `.accdb`/`.mdb` files. The project is built for the `x86` platform to match this provider's bitness.
- NuGet package `ClosedXML` (added in v3) for generating the `Failures.xlsx` report — restored automatically on build.

## Building

Open `Absa.sln` in Visual Studio 2022 and build/run from there (recommended), or use the full-framework MSBuild directly:

```powershell
& "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" Absa.sln -p:Configuration=Debug -p:Platform=x86
```

> Running `dotnet build Absa.sln` will fail with `MSB4803: The task "ResolveComReference" is not supported on the .NET Core version of MSBuild` — this is expected due to the ADOX COM reference.

A convenience `apply_changes.bat` script is included that restores packages, locates the correct MSBuild, and builds the solution.

## Usage

1. Launch **AbsaConverterTool**. The title bar reads **"ABSA Tool v03"**.
2. Click **Choose Folder** and select a folder containing the input files. The tool parses every file in that folder and reports:
   - Successfully parsed, product-valid records in the log window (per file: valid record count and failed record count)
   - Any files that failed to parse outright, listed separately (unchanged from v2 — a hard parse exception, e.g. a byte/text record-count mismatch)
3. Click **Process Files**. For each input file, the tool:
   - Groups that file's valid records by Product Code
   - Writes one Access database per product group, named `<InputFileName>_<ProductName>.mdb`, into the same folder
   - Writes `Failures.xlsx` into the same folder if any record across the batch could not be placed into a product bucket
4. Review the summary in the log window — each generated `.mdb` is reported as **Created**, **Skipped** (no valid records for that file), or **Failed** (with the underlying error), alongside which product(s) it contains.

### Input file format

Input files contain one or more records separated by an `#END#` delimiter. The first record in each file is treated as a header and discarded. Within each record:

| Field | Extracted from |
|---|---|
| Product Code | The 4 characters immediately following the first `'` character in the record |
| PAN | Text between `*` and `@` |
| Expiry (`IDWEXP`) | Fixed-length text after `$`, offset 13, length 5 |
| Card Holder Name | Text starting immediately after `)` (up to the next `@`, or end of record), trimmed, capped at **26 characters** |
| CVV2 | Text between `:` and `@` |
| Track 1 / Track 2 | Text between `"`/`@`, then split on `%`/`?` and `;`/`?` |
| Chip data (`IDWChip1`–`IDWChip4`) | Binary block following `{` + a 7-digit length prefix, decoded to hex and split sequentially into 4 chunks |

### Product mapping

| Product | Product Code |
|---|---|
| Business Credit | `AGVR` |
| Business Debit | `AGVG` |
| Classic Credit | `AGVK` |
| Platinum Credit | `AGVL` |
| Signature Credit | `AGVQ` |
| Infinite Credit | `AGVN` |
| Signature Debit | `AGVQ` |
| Infinite Debit | `AGVP` |
| International Debit | `AGVB` |
| Personal Debit | `AGVA` |
| Prepaid | `BBGP` |
| Prestige Debit | `AGVC` |
| Premier Debit | `AGVE` |

**`AGVQ` is shared** by Signature Credit and Signature Debit. Records for both go into **one** output `.mdb` (they are never split into two files because they share a code), and that file's result line displays both product names: `Signature Credit + Signature Debit`.

A record whose Product Code is **missing** or **not in this table** is not written to any `.mdb`. It is recorded in `Failures.xlsx` with reason `missing product code` or `unrecognized product code: <code>`, and processing continues with the rest of the batch.

### Output file naming

```
<InputFileName>_<ProductName>.mdb
```

`ProductName` has spaces removed (e.g. `BusinessCredit`); for the shared `AGVQ` bucket the token is `SignatureCredit_SignatureDebit`. If a name collision occurs (e.g. two input files with the same base name and the same product), a numeric suffix is appended (`..._2.mdb`, `..._3.mdb`, ...), same de-duplication behavior as v2.

### Output schema

Each generated `.mdb` file contains a single `Cards` table with the columns: `IDWAutoNumber`, `JobNumber`, `IDWPAN`, `IDWEXP`, `IDWNAME`, `IDWCVV2`, `IDWTrack1`, `IDWTrack2`, `IDWChip1`, `IDWChip2`, `IDWChip3`, `IDWChip4`. `IDWAutoNumber` restarts at 1 within each generated file, same as v2's per-file numbering.

**Chip columns are intentionally `TEXT` (255 characters each), unchanged from v2.** This is a required, load-bearing schema constraint for a downstream personalization system and is not a bug. The combined chip string (`{` + 7-digit length + hex) therefore has a hard limit of **1020 characters** (4 × 255). A record whose chip data would exceed 1020 characters is **not truncated and not inserted** — it is failed with reason `chip data exceeds 1020-char schema limit` and reported in `Failures.xlsx`.

### `Failures.xlsx`

Generated in the selected folder whenever at least one record across the batch fails validation. One row per failed record, with exactly these columns:

| Column | Content |
|---|---|
| Masked PAN | First 6 and last 4 digits of the PAN, middle digits replaced with `*` (PANs of 10 digits or fewer are fully masked) |
| File Name | The input file the record came from |
| Record Index | 1-based position of the record within its input file (matches the old `IDWAutoNumber` numbering) |
| Product Code | The extracted code, or blank if none could be extracted |
| Reason | `missing product code`, `unrecognized product code: <code>`, `chip data exceeds 1020-char schema limit`, or `parse error: <message>` |

The full PAN, CVV2, track data, and chip data are never written to this report.

## Project structure

```
AbsaConverterTool/
├── MainWindow.xaml / .cs   UI and workflow orchestration (product grouping, per-product .mdb naming)
├── Helper/FileHelper.cs    Parsing, product mapping, field extraction, Access DB creation, Excel report
└── App.xaml / .cs          Application entry point
```

## Known limitations

- The input folder scan reads *every* file in the folder (no extension filter). Re-running **Choose Folder** on a folder that already contains previously generated `*_Output.mdb`/`Failures.xlsx` files will attempt to parse those too; a non-card file typically fails as a parse error (shown in the file-level error list) and does not affect the rest of the batch.
- Card Holder Name extraction assumes the name field ends at the next `@` character after the opening `)` (consistent with how every other field in this format is delimited). If a particular file's layout doesn't follow that convention, the extracted name may need review.
- Chip data over 1020 combined characters is a hard failure (by design — see above), not a truncation.
