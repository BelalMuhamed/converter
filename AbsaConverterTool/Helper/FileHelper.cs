using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using ADOX;
using ClosedXML.Excel;

using System.Text;

namespace AbsaConverterTool.Helper
{
    public class FileProcessingResult
    {
        public string FilePath { get; set; }
        public List<byte[]> RecordsInBytes { get; set; } = new();
        public List<string> RecordsInText { get; set; } = new();
        public string ErrorMessage { get; set; } = null;
    }

    /// <summary>
    /// One product entry from the confirmed mapping table. The mapping is keyed by
    /// code and grouped generically: if two product names ever share a code, one
    /// ProductInfo carries both display names and both resolve to one output bucket.
    /// As of the latest mapping, Signature Credit (AGVO) and Signature Debit (AGVQ)
    /// have distinct codes and are no longer combined — this is not special-cased
    /// anywhere; it falls out of the code-keyed grouping automatically.
    /// </summary>
    public class ProductInfo
    {
        public string Code { get; set; }
        public List<string> DisplayNames { get; } = new();

        /// <summary>Human-readable name(s) for UI/log display, e.g. "Signature Credit + Signature Debit"
        /// if a code is ever shared again, or just "Signature Credit" for a single-name bucket.</summary>
        public string DisplayName => string.Join(" + ", DisplayNames);

        /// <summary>Token used inside generated file names, e.g. "SignatureCredit" or, for a shared code,
        /// "SignatureCredit_SignatureDebit".</summary>
        public string FileNameToken => string.Join("_", DisplayNames.Select(n => n.Replace(" ", string.Empty)));
    }

    /// <summary>
    /// Confirmed Product Code -> Product mapping. Grouping by code is generic: it only
    /// combines two product names into one bucket if they actually share a code. Update
    /// this table and the grouping/naming keep working without further code changes.
    /// </summary>
    public static class ProductMapping
    {
        public static readonly IReadOnlyDictionary<string, ProductInfo> ByCode = Build();

        private static Dictionary<string, ProductInfo> Build()
        {
            // (Code, DisplayName) pairs exactly as confirmed. Signature Credit = AGVO,
            // Signature Debit = AGVQ — distinct codes, so they land in separate .mdb files.
            var raw = new (string Code, string Name)[]
            {
                ("AGVR", "Business Credit"),
                ("AGVG", "Business Debit"),
                ("AGVK", "Classic Credit"),
                ("AGVL", "Platinum Credit"),
                ("AGVO", "Signature Credit"),
                ("AGVN", "Infinite Credit"),
                ("AGVQ", "Signature Debit"),
                ("AGVP", "Infinite Debit"),
                ("AGVB", "International Debit"),
                ("AGVA", "Personal Debit"),
                ("BBGP", "Prepaid"),
                ("AGVC", "Prestige Debit"),
                ("AGVE", "Premier Debit"),
            };

            var map = new Dictionary<string, ProductInfo>(StringComparer.OrdinalIgnoreCase);
            foreach (var (code, name) in raw)
            {
                if (!map.TryGetValue(code, out var info))
                {
                    info = new ProductInfo { Code = code };
                    map[code] = info;
                }
                info.DisplayNames.Add(name);
            }

            return map;
        }

        public static bool TryGet(string code, out ProductInfo info)
        {
            if (string.IsNullOrEmpty(code))
            {
                info = null;
                return false;
            }

            return ByCode.TryGetValue(code, out info);
        }
    }

    /// <summary>A single successfully-parsed, product-valid card record.</summary>
    public class CardRecord
    {
        public int RecordIndex { get; set; }
        public string ProductCode { get; set; }
        public string PAN { get; set; }
        public string EXP { get; set; }
        public string Name { get; set; }
        public string CVV2 { get; set; }
        public string Track1 { get; set; }
        public string Track2 { get; set; }

        /// <summary>The full combined chip string ("{" + 7-digit length + uppercase hex), not yet
        /// split into IDWChip# columns. Splitting happens in BuildCardsDataTable, once the number
        /// of chip columns needed for the whole output file is known.</summary>
        public string ChipData { get; set; }
    }

    /// <summary>A record that could not be placed into any output .mdb, with the reason why.</summary>
    public class FailedRecord
    {
        public string FileName { get; set; }
        public int RecordIndex { get; set; }
        public string MaskedPan { get; set; }
        public string ProductCode { get; set; }
        public string Reason { get; set; }
    }

    /// <summary>Result of parsing one input file: valid records (ready for grouping/MDB), failed records
    /// (destined for the Excel report), and any hard parse exceptions (kept for the existing error UI).</summary>
    public class FileParseResult
    {
        public string FilePath { get; set; }
        public List<CardRecord> Records { get; } = new();
        public List<FailedRecord> FailedRecords { get; } = new();
        public List<string> ParseErrors { get; } = new();
    }

    public static class FileHelper
    {
        /// <summary>Capacity of a single IDWChip# TEXT column. The combined chip string is split
        /// into sequential chunks of this size; as many IDWChip# columns are created as needed
        /// (minimum 4, matching the original fixed schema) so chip data is never truncated.</summary>
        private const int ChipColumnSize = 255;

        /// <summary>Minimum number of IDWChip# columns always present, matching the original schema.</summary>
        private const int MinChipColumns = 4;

        /// <summary>Confirmed fixed length of the Card Holder Name field.</summary>
        private const int CardHolderNameMaxLength = 26;

        public static List<FileProcessingResult> GetAllFilesRecords(string folderPath)
        {
            var results = new List<FileProcessingResult>();

            if (!Directory.Exists(folderPath))
                throw new DirectoryNotFoundException($"Folder not found: {folderPath}");

            var files = new DirectoryInfo(folderPath).GetFiles("*", SearchOption.TopDirectoryOnly);

            var delimiterBytes = Encoding.ASCII.GetBytes("#END#");

            foreach (var file in files)
            {
                var result = new FileProcessingResult { FilePath = file.FullName };
                try
                {
                    byte[] allBytes = File.ReadAllBytes(file.FullName);

                    // تقسيم الـ bytes
                    result.RecordsInBytes = SplitRecords(allBytes, delimiterBytes);

                    // تحويل للـ text
                    string fileContent = Encoding.ASCII.GetString(allBytes); // استخدم ASCII أو UTF8 حسب الحاجة
                    result.RecordsInText = fileContent
                        .Split(new string[] { "#END#" }, StringSplitOptions.RemoveEmptyEntries)
                        .ToList();


                    if (result.RecordsInBytes.Count != result.RecordsInText.Count)
                    {
                        result.ErrorMessage = $"Mismatch between bytes and text count. Bytes: {result.RecordsInBytes.Count}, Text: {result.RecordsInText.Count}";
                    }
                    else
                    {

                        if (result.RecordsInBytes.Count > 0)
                        {
                            result.RecordsInBytes.RemoveAt(0);
                            result.RecordsInText.RemoveAt(0);
                        }
                    }
                }
                catch (Exception ex)
                {
                    result.ErrorMessage = ex.Message;
                }

                results.Add(result);
            }

            return results;
        }

        public static List<byte[]> SplitRecords(byte[] fileBytes, byte[] delimiter)
        {
            var records = new List<byte[]>();
            int start = 0;

            for (int i = 0; i <= fileBytes.Length - delimiter.Length; i++)
            {
                bool match = true;
                for (int j = 0; j < delimiter.Length; j++)
                {
                    if (fileBytes[i + j] != delimiter[j])
                    {
                        match = false;
                        break;
                    }
                }

                if (match)
                {
                    int length = i - start;
                    if (length > 0)
                    {
                        byte[] record = new byte[length];
                        Array.Copy(fileBytes, start, record, 0, length);
                        records.Add(record);
                    }
                    start = i + delimiter.Length;
                    i = start - 1;
                }
            }

            // آخر record
            if (start < fileBytes.Length)
            {
                byte[] record = new byte[fileBytes.Length - start];
                Array.Copy(fileBytes, start, record, 0, record.Length);
                records.Add(record);
            }

            return records;
        }

        /// <summary>
        /// Finds the raw chip byte block after "{" + a 7-digit length prefix.
        /// Refactored: the 7-digit length prefix is now returned instead of being
        /// stashed in a static field, so this method has no shared mutable state.
        /// </summary>
        private static (string LengthText, byte[] Data) ExtractRawChipData(byte[] recordBytes)
        {
            int braceIndex = Array.IndexOf(recordBytes, (byte)'{');
            if (braceIndex == -1 || braceIndex + 8 > recordBytes.Length)
                return (null, Array.Empty<byte>());

            int startIndex = braceIndex + 1;

            string lengthText = Encoding.ASCII.GetString(recordBytes, startIndex, 7);
            if (!int.TryParse(lengthText, out int byteCount))
                throw new Exception("Invalid length format in chip data");

            int dataStartIndex = startIndex + 7;

            if (dataStartIndex + byteCount > recordBytes.Length)
                throw new Exception("Chip data length exceeds record size");

            byte[] chipData = new byte[byteCount];
            Array.Copy(recordBytes, dataStartIndex, chipData, 0, byteCount);

            return (lengthText, chipData);
        }

        /// <summary>
        /// Rebuilds the "{" + 7-digit length + uppercase-hex string. Byte-for-byte identical
        /// output to the original implementation; the length prefix is now passed in explicitly
        /// instead of being read from a static field.
        /// </summary>
        private static string ReadBinaryChipData(string lengthText, byte[] binaryData)
        {
            if (string.IsNullOrEmpty(lengthText))
                return string.Empty;

            try
            {
                StringBuilder stringBuilder = new StringBuilder((binaryData.Length * 2) + 8);
                foreach (byte b in binaryData)
                {
                    stringBuilder.Append($"{b:X2}");
                }
                stringBuilder.Insert(0, lengthText);

                stringBuilder.Insert(0, "{");
                return stringBuilder.ToString();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error reading Binary Data: {ex.Message}");
                return string.Empty;
            }
        }

        /// <summary>
        /// Extract fixed-length substring from text between two delimiters.
        /// If the extracted text is shorter than expected, it pads with spaces.
        /// </summary>
        public static string ExtractBetweenWithFixedLength(string text, char startDelimiter, char endDelimiter, int offset, int length)
        {
            if (string.IsNullOrEmpty(text))
                return new string(' ', length);

            int startIndex = text.IndexOf(startDelimiter);
            if (startIndex == -1) return new string(' ', length);

            startIndex += 1; // نبدأ بعد الـ delimiter

            int endIndex = text.IndexOf(endDelimiter, startIndex);
            if (endIndex == -1) endIndex = text.Length;

            string between = text.Substring(startIndex, endIndex - startIndex);

            // الآن ناخد الـ offset و length المطلوبين
            if (between.Length <= offset)
                return new string(' ', length);

            int maxLength = Math.Min(length, between.Length - offset);
            string result = between.Substring(offset, maxLength);

            // pad إذا أقصر من المطلوب
            if (result.Length < length)
                result = result.PadRight(length, ' ');

            return result;
        }

        public static string ExtractLine(string text, string prefix, int length)
        {
            var index = text.IndexOf(prefix);
            if (index == -1) return null;

            var start = index + prefix.Length;
            if (start + length > text.Length) return null;

            return text.Substring(start, length).Trim();
        }

        public static string ExtractBetween(string text, string startDelim, string endDelim)
        {
            var start = text.IndexOf(startDelim);
            if (start == -1) return null;

            start += startDelim.Length;

            if (endDelim == null)
                return text.Substring(start).Trim();

            var end = text.IndexOf(endDelim, start);
            if (end == -1) return null;

            return text.Substring(start, end - start).Trim();
        }

        /// <summary>
        /// Confirmed rule: the Product Code is the 4 characters immediately following the
        /// first ' character in the record. Uses the ' as the anchor rather than a fixed
        /// offset, so it keeps working if the leading record counter's width changes.
        /// </summary>
        public static string ExtractProductCode(string record)
        {
            if (string.IsNullOrEmpty(record))
                return null;

            int quoteIndex = record.IndexOf('\'');
            if (quoteIndex == -1)
                return null;

            int start = quoteIndex + 1;
            if (start + 4 > record.Length)
                return null;

            return record.Substring(start, 4);
        }

        /// <summary>
        /// Confirmed rule: Card Holder Name starts immediately after ')'. The field's end is
        /// not separately specified beyond "fixed 26-character field", so — consistent with
        /// every other delimited field in this file format — extraction stops at the next '@'
        /// (or end of record if none exists) before trimming and applying the 26-char cap.
        /// If this end boundary turns out to be wrong for some records, it's isolated to this
        /// one method.
        /// </summary>
        public static string ExtractCardHolderName(string record)
        {
            if (string.IsNullOrEmpty(record))
                return string.Empty;

            int startIndex = record.IndexOf(')');
            if (startIndex == -1)
                return string.Empty;

            startIndex += 1;
            if (startIndex >= record.Length)
                return string.Empty;

            int endIndex = record.IndexOf('@', startIndex);
            string raw = endIndex == -1
                ? record.Substring(startIndex)
                : record.Substring(startIndex, endIndex - startIndex);

            string trimmed = raw.Trim();

            return trimmed.Length > CardHolderNameMaxLength
                ? trimmed.Substring(0, CardHolderNameMaxLength)
                : trimmed;
        }

        /// <summary>
        /// Masks a PAN to first-6 + last-4 for the failure report (never store/display the
        /// full PAN in the report). Short/unusable values are fully masked rather than shown.
        /// </summary>
        public static string MaskPan(string pan)
        {
            if (string.IsNullOrWhiteSpace(pan))
                return string.Empty;

            string digits = pan.Trim();
            if (digits.Length <= 10)
                return new string('*', digits.Length);

            string first6 = digits.Substring(0, 6);
            string last4 = digits.Substring(digits.Length - 4);
            string middleMask = new string('*', digits.Length - 10);
            return $"{first6}{middleMask}{last4}";
        }

        public static void CreateEmptyAccdb(string path)
        {
            string folder = Path.GetDirectoryName(path);
            if (!Directory.Exists(folder))
                Directory.CreateDirectory(folder);

            if (!File.Exists(path))
            {
                try
                {
                    Catalog catalog = new Catalog();
                    string connStr = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={path};Jet OLEDB:Engine Type=5;";
                    catalog.Create(connStr);
                }
                catch (Exception ex)
                {
                    throw new Exception($"Failed to create ACCDB file at {path}. Error: {ex.Message}");
                }
            }
        }

        public static void CreateMdbFromDataTable(DataTable table, string mdbPath, string tableName = "Cards")
        {
            if (table == null || table.Rows.Count == 0)
                throw new ArgumentException("DataTable is empty.");

            string folder = Path.GetDirectoryName(mdbPath);
            if (!Directory.Exists(folder))
            {
                Directory.CreateDirectory(folder); // أنشئ المجلد إذا لم يكن موجود
            }

            if (!File.Exists(mdbPath))
            {
                try
                {
                    CreateEmptyAccdb(mdbPath);
                }
                catch (Exception ex)
                {
                    throw new Exception($"Failed to create ACCDB file at {mdbPath}. Error: {ex.Message}");
                }
            }
            string connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={mdbPath};Persist Security Info=False;";

            using (OleDbConnection conn = new OleDbConnection(connectionString))
            {
                conn.Open();

                using (OleDbTransaction transaction = conn.BeginTransaction())
                {
                    try
                    {
                        // Create table
                        using (OleDbCommand cmd = new OleDbCommand(BuildCreateTableQuery(table, tableName), conn, transaction))
                        {
                            cmd.ExecuteNonQuery();
                        }

                        // Prepare insert command once
                        using (OleDbCommand cmd = new OleDbCommand(BuildInsertQuery(table, tableName), conn, transaction))
                        {
                            foreach (DataRow row in table.Rows)
                            {
                                cmd.Parameters.Clear();
                                for (int i = 0; i < table.Columns.Count; i++)
                                {
                                    object value = row[i] ?? DBNull.Value;

                                    if (value is byte[] bytes)
                                        cmd.Parameters.Add("?", OleDbType.Binary).Value = bytes;
                                    else
                                        cmd.Parameters.AddWithValue("?", value);
                                }
                                cmd.ExecuteNonQuery();
                            }
                        }

                        transaction.Commit();
                    }
                    catch
                    {
                        transaction.Rollback();
                        throw;
                    }
                }
            }
        }

        /// <summary>
        /// Builds the "Cards" DataTable from an already-grouped list of valid CardRecords for
        /// a single product bucket. IDWAutoNumber restarts at 1 per generated .mdb, matching the
        /// original per-file numbering.
        ///
        /// Chip columns are dynamic: this file's table always has at least IDWChip1..IDWChip4
        /// (matching the original fixed schema), and gains IDWChip5, IDWChip6, ... only if some
        /// record's combined chip string is longer than 4 x 255 = 1020 characters. Every record
        /// in this table shares the same column count — it's sized once, up front, to the widest
        /// chip data in the group — so a single long outlier widens the whole output file, not
        /// just its own row. No chip data is ever truncated or dropped.
        /// </summary>
        public static DataTable BuildCardsDataTable(List<CardRecord> records)
        {
            var table = new DataTable("Cards");

            table.Columns.Add("IDWAutoNumber", typeof(int));
            table.Columns.Add("JobNumber", typeof(int));
            table.Columns.Add("IDWPAN", typeof(string));
            table.Columns.Add("IDWEXP", typeof(string));
            table.Columns.Add("IDWNAME", typeof(string));
            table.Columns.Add("IDWCVV2", typeof(string));
            table.Columns.Add("IDWTrack1", typeof(string));
            table.Columns.Add("IDWTrack2", typeof(string));

            int maxChipLength = records.Count > 0 ? records.Max(r => r.ChipData?.Length ?? 0) : 0;
            int chipColumnCount = Math.Max(MinChipColumns, (int)Math.Ceiling(maxChipLength / (double)ChipColumnSize));

            for (int c = 1; c <= chipColumnCount; c++)
                table.Columns.Add($"IDWChip{c}", typeof(string));

            int autoNumber = 1;
            foreach (var r in records)
            {
                var row = table.NewRow();
                row["IDWAutoNumber"] = autoNumber++;
                row["JobNumber"] = 1;
                row["IDWPAN"] = (object)r.PAN ?? DBNull.Value;
                row["IDWEXP"] = (object)r.EXP ?? DBNull.Value;
                row["IDWNAME"] = (object)r.Name ?? DBNull.Value;
                row["IDWCVV2"] = (object)r.CVV2 ?? DBNull.Value;
                row["IDWTrack1"] = (object)r.Track1 ?? DBNull.Value;
                row["IDWTrack2"] = (object)r.Track2 ?? DBNull.Value;

                var chunks = SplitChipData(r.ChipData, ChipColumnSize);
                for (int c = 1; c <= chipColumnCount; c++)
                {
                    row[$"IDWChip{c}"] = c <= chunks.Count ? (object)chunks[c - 1] : DBNull.Value;
                }

                table.Rows.Add(row);
            }

            return table;
        }

        /// <summary>
        /// Splits the combined chip string ("{" + 7-digit length + hex) into sequential,
        /// fixed-size chunks of up to <paramref name="chunkSize"/> characters each — e.g. for
        /// a 1040-char string and chunkSize 255: four 255-char chunks (1020 chars) plus one
        /// final 20-char chunk. This replaces the original proportional 4-way split
        /// (chunkSize = ceil(total/4)); the trade-off is an intentional one, confirmed alongside
        /// the move to dynamic columns: with a fixed 255-char chunk, a record's chip data can
        /// leave later baseline columns (IDWChip2-4) null even when the total is under 1020,
        /// instead of spreading it evenly across all 4 as before. No data is ever lost.
        /// </summary>
        private static List<string> SplitChipData(string chipData, int chunkSize)
        {
            var chunks = new List<string>();
            if (string.IsNullOrEmpty(chipData))
                return chunks;

            for (int offset = 0; offset < chipData.Length; offset += chunkSize)
            {
                int length = Math.Min(chunkSize, chipData.Length - offset);
                chunks.Add(chipData.Substring(offset, length));
            }

            return chunks;
        }

        /// <summary>
        /// Parses every record of one input file into valid CardRecords (product-known) and
        /// FailedRecords (missing/unknown product code, or a parse exception). Chip data is never
        /// a failure reason — its column count grows dynamically instead, see BuildCardsDataTable.
        /// Never throws for a single bad record — the record is captured as a FailedRecord instead
        /// so the rest of the file continues.
        /// </summary>
        public static FileParseResult ParseFileRecords(List<string> records, List<byte[]> recordsInBytes, string filePath)
        {
            var result = new FileParseResult { FilePath = filePath };
            string fileName = Path.GetFileName(filePath);

            for (int i = 0; i < records.Count; i++)
            {
                int recordIndex = i + 1; // 1-based, matches previous IDWAutoNumber numbering
                string rawRecord = records[i];
                string productCode = null;
                string maskedPan = string.Empty;

                try
                {
                    productCode = ExtractProductCode(rawRecord);

                    string IDWPAN = ExtractBetween(rawRecord, "*", "@");
                    maskedPan = MaskPan(IDWPAN);

                    if (string.IsNullOrEmpty(productCode))
                    {
                        result.FailedRecords.Add(new FailedRecord
                        {
                            FileName = fileName,
                            RecordIndex = recordIndex,
                            MaskedPan = maskedPan,
                            ProductCode = string.Empty,
                            Reason = "missing product code"
                        });
                        continue;
                    }

                    if (!ProductMapping.TryGet(productCode, out var productInfo))
                    {
                        result.FailedRecords.Add(new FailedRecord
                        {
                            FileName = fileName,
                            RecordIndex = recordIndex,
                            MaskedPan = maskedPan,
                            ProductCode = productCode,
                            Reason = $"unrecognized product code: {productCode}"
                        });
                        continue;
                    }

                    string IDWEXP = ExtractBetweenWithFixedLength(rawRecord, '$', '@', 13, 5);
                    string IDWNAME = ExtractCardHolderName(rawRecord);
                    string IDWCVV2 = ExtractBetween(rawRecord, ":", "@");
                    string track1AndTrack2 = ExtractBetween(rawRecord, "\"", "@");
                    string IDWTrack1 = ExtractBetween(track1AndTrack2, "%", "?");
                    string IDWTrack2 = ExtractBetween(track1AndTrack2, ";", "?");

                    // No length ceiling here by design: ExtractRawChipData already validates the
                    // declared length against the record's own byte length (throws "Chip data
                    // length exceeds record size" for a corrupted/garbage prefix), so chipData
                    // can never exceed the record it came from. Splitting into as many IDWChip#
                    // columns as needed happens later, in BuildCardsDataTable, once every
                    // record's chip length in this product group is known.
                    var (lengthText, rawChip) = ExtractRawChipData(recordsInBytes[i]);
                    string chipData = ReadBinaryChipData(lengthText, rawChip);

                    result.Records.Add(new CardRecord
                    {
                        RecordIndex = recordIndex,
                        ProductCode = productCode,
                        PAN = IDWPAN,
                        EXP = IDWEXP,
                        Name = IDWNAME,
                        CVV2 = IDWCVV2,
                        Track1 = IDWTrack1,
                        Track2 = IDWTrack2,
                        ChipData = chipData
                    });
                }
                catch (Exception ex)
                {
                    result.ParseErrors.Add($"File: {filePath}, Record Index: {i}, Error: {ex.Message}");
                    result.FailedRecords.Add(new FailedRecord
                    {
                        FileName = fileName,
                        RecordIndex = recordIndex,
                        MaskedPan = maskedPan,
                        ProductCode = productCode ?? string.Empty,
                        Reason = $"parse error: {ex.Message}"
                    });
                }
            }

            return result;
        }

        /// <summary>
        /// Writes the confirmed Failures.xlsx report: Masked PAN, File Name, Record Index,
        /// Product Code, Reason. Only ever called when at least one FailedRecord exists.
        /// </summary>
        public static void WriteFailuresReport(List<FailedRecord> failedRecords, string outputPath)
        {
            string folder = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(folder) && !Directory.Exists(folder))
                Directory.CreateDirectory(folder);

            using var workbook = new XLWorkbook();
            var sheet = workbook.Worksheets.Add("Failures");

            sheet.Cell(1, 1).Value = "Masked PAN";
            sheet.Cell(1, 2).Value = "File Name";
            sheet.Cell(1, 3).Value = "Record Index";
            sheet.Cell(1, 4).Value = "Product Code";
            sheet.Cell(1, 5).Value = "Reason";
            sheet.Row(1).Style.Font.Bold = true;

            int row = 2;
            foreach (var failure in failedRecords)
            {
                sheet.Cell(row, 1).Value = failure.MaskedPan;
                sheet.Cell(row, 2).Value = failure.FileName;
                sheet.Cell(row, 3).Value = failure.RecordIndex;
                sheet.Cell(row, 4).Value = failure.ProductCode;
                sheet.Cell(row, 5).Value = failure.Reason;
                row++;
            }

            sheet.Columns().AdjustToContents();
            workbook.SaveAs(outputPath);
        }

        private static string BuildCreateTableQuery(DataTable table, string tableName)
        {
            string query = $"CREATE TABLE [{tableName}] (";

            foreach (DataColumn col in table.Columns)
            {
                query += $"[{col.ColumnName}] {MapType(col.DataType)},";
            }

            query = query.TrimEnd(',') + ")";
            return query;
        }

        private static string BuildInsertQuery(DataTable table, string tableName)
        {
            string columns = "";
            string values = "";

            foreach (DataColumn col in table.Columns)
            {
                columns += $"[{col.ColumnName}],";
                values += "?,";
            }

            columns = columns.TrimEnd(',');
            values = values.TrimEnd(',');

            return $"INSERT INTO [{tableName}] ({columns}) VALUES ({values})";
        }

        // NOTE: chip columns are deliberately kept as Access TEXT (255-char cap each) —
        // confirmed as an intentional, load-bearing pattern for a downstream personalization
        // system. Do not change this to MEMO/Long Text. There is no longer a hard total-length
        // limit: BuildCardsDataTable adds as many IDWChip# TEXT columns as the widest chip
        // string in the group needs, so every column individually still respects the 255-char
        // cap while the total capacity grows with the data.
        private static string MapType(Type type, string columnName = "")
        {

            if (columnName.StartsWith("Chip")) return "MEMO";

            if (type == typeof(string)) return "TEXT";
            if (type == typeof(int)) return "INTEGER";
            if (type == typeof(DateTime)) return "DATETIME";
            if (type == typeof(double) || type == typeof(float)) return "DOUBLE";
            if (type == typeof(byte[])) return "OLEOBJECT";
            return "TEXT";
        }


    }
}
