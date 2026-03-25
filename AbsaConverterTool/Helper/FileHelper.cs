using System.Data;
using System.Data.OleDb;
using System.IO;
using ADOX; 

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
    public static class FileHelper
    {
        static string lengthText;
   
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
        private static byte[] ExtractRawChipData(byte[] recordBytes)
        {
            int braceIndex = Array.IndexOf(recordBytes, (byte)'{');
            if (braceIndex == -1 || braceIndex + 8 > recordBytes.Length)
                return Array.Empty<byte>();

            int startIndex = braceIndex + 1;

            lengthText = Encoding.ASCII.GetString(recordBytes, startIndex, 7);
            if (!int.TryParse(lengthText, out int byteCount))
                throw new Exception("Invalid length format in chip data");

            int dataStartIndex = startIndex + 7;

            if (dataStartIndex + byteCount > recordBytes.Length)
                throw new Exception("Chip data length exceeds record size");

            byte[] chipData = new byte[byteCount];
            Array.Copy(recordBytes, dataStartIndex, chipData, 0, byteCount);

            return chipData;
        }

        private static string ReadBinaryChipData(byte[] binaryData)
        {
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
        
        public static (DataTable Table, List<string> Errors) ParseRecordsOfFileToDataTable(
        List<string> records,
        List<byte[]> recordsInBytes,
        string filePath
    )
        {
            var table = new DataTable("Cards");
            var errors = new List<string>();

            // الأعمدة
            table.Columns.Add("IDWAutoNumber", typeof(int)); // Auto increment
            table.Columns.Add("JobNumber", typeof(int));     // Constant 1
            table.Columns.Add("IDWPAN", typeof(string));
            table.Columns.Add("IDWEXP", typeof(string));
            table.Columns.Add("IDWNAME", typeof(string));
            table.Columns.Add("IDWCVV2", typeof(string));
            table.Columns.Add("IDWTrack1", typeof(string));
            table.Columns.Add("IDWTrack2", typeof(string));
            table.Columns.Add("IDWChip1", typeof(string));
            table.Columns.Add("IDWChip2", typeof(string));
            table.Columns.Add("IDWChip3", typeof(string));
            table.Columns.Add("IDWChip4", typeof(string));

            for (int i = 0; i < records.Count; i++)
            {
                try
                {
                    string IDWPAN = ExtractBetween(records[i], "*", "@");
                    string IDWEXP = ExtractBetweenWithFixedLength(records[i], '$', '@', 13, 5);
                    string IDWNAME = ExtractBetween(records[i], ") ", "@");
                    string IDWCVV2 = ExtractBetween(records[i], ":", "@");
                    string track1AndTrack2 = ExtractBetween(records[i], "\"", "@");
                    string IDWTrack1 = ExtractBetween(track1AndTrack2, "%", "?");
                    string IDWTrack2 = ExtractBetween(track1AndTrack2, ";", "?");

                    // قراءة البيانات الثنائية
                    string chipData = ReadBinaryChipData(ExtractRawChipData(recordsInBytes[i]));

                    // قسمها على 4 أعمدة
                    int chunkSize = (int)Math.Ceiling(chipData.Length / 4.0);
                    string IDWchip1 = chipData.Length >= 1 ? chipData.Substring(0, Math.Min(chunkSize, chipData.Length)) : null;
                    string IDWchip2 = chipData.Length > chunkSize ? chipData.Substring(chunkSize, Math.Min(chunkSize, chipData.Length - chunkSize)) : null;
                    string IDWchip3 = chipData.Length > chunkSize * 2 ? chipData.Substring(chunkSize * 2, Math.Min(chunkSize, chipData.Length - chunkSize * 2)) : null;
                    string IDWchip4 = chipData.Length > chunkSize * 3 ? chipData.Substring(chunkSize * 3, Math.Min(chunkSize, chipData.Length - chunkSize * 3)) : null;

                    table.Rows.Add(
                        i + 1,       // IDWAutoNumber
                        1,           // JobNumber
                        IDWPAN,
                        IDWEXP,
                        IDWNAME,
                        IDWCVV2,
                        IDWTrack1,
                        IDWTrack2,
                        IDWchip1,
                        IDWchip2,
                        IDWchip3,
                       IDWchip4
                    );
                }
                catch (Exception ex)
                {
                    errors.Add($"File: {filePath}, Record Index: {i}, Error: {ex.Message}");
                }
            }

            return (table, errors);
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