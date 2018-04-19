using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Security.Cryptography.Xml;
using System.Text;
using System.Threading.Tasks;
using Mono.Options;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;

namespace SimpleXlsxDiff
{
    static class MainProcess
    {
        public static void Execute(string[] args)
        {
            bool h;
            string @base = null;
            string file1 = null;
            string file2 = null;
            string json = null;
            var opetios = new OptionSet()
            {
                {"h|help", val => h = val != null},
                {"b|base=", val => @base = val},
                {"f1|file1=", val => file1 = val},
                {"f2|file2=", val => file2 = val},
                {"j|json=", val => json = val},
            };

            List<string> extra;
            try
            {
                extra = opetios.Parse(args);
            }
            catch (OptionException e)
            {
                Trace.TraceError(e.Message);
                Trace.TraceError("Try 'sxd' --help' for more information.");
                return;
            }


            JObject jObject = null;
            if (json != null)
            {
                string jsonText;

                try
                {
                    jsonText = File.Exists(json) ? File.ReadAllText(json) : json;
                }
                catch (Exception e)
                {
                    Trace.TraceError(e.Message);
                    return;
                }


                try
                {
                    jObject = JObject.Parse(jsonText);
                }
                catch (JsonReaderException e)
                {
                    Trace.TraceError(e.Message);
                    return;
                }
            }


            if (@base == null && jObject != null)
            {
                JToken tmp;
                var jvalue = (jObject.TryGetValue("b", out tmp) ? tmp : jObject.TryGetValue("base", out tmp) ? tmp : null) as JValue;
                @base = jvalue?.Value as string;
            }

            if (file1 == null && jObject != null)
            {
                JToken tmp;
                var jvalue = (jObject.TryGetValue("f1", out tmp) ? tmp : jObject.TryGetValue("file1", out tmp) ? tmp : null) as JValue;
                file1 = jvalue?.Value as string;
            }

            if (file2 == null && jObject != null)
            {
                JToken tmp;
                var jvalue = (jObject.TryGetValue("f2", out tmp) ? tmp : jObject.TryGetValue("file2", out tmp) ? tmp : null) as JValue;
                file2 = jvalue?.Value as string;
            }

            List<RangeObject> ranges = new List<RangeObject>();


            if (jObject != null)
            {
                JToken tmp;
                var jToken = (jObject.TryGetValue("r", out tmp) ? tmp : jObject.TryGetValue("range", out tmp) ? tmp : null);

                if (jToken is JObject jObj)
                {
                    ranges.Add(RangeObject.Create(jObj));
                }
                else if (jToken is JArray jArray)
                {
                    ranges.AddRange(jArray.OfType<JObject>().Select(RangeObject.Create));
                }

            }

            if (!File.Exists(@base))
            {
                Console.WriteLine($"'{@base}' is not found.");
                return;
            }

            if (!File.Exists(file1))
            {
                Console.WriteLine($"'{file1}' is not found.");
                return;
            }

            if (!File.Exists(file2))
            {
                Console.WriteLine($"'{file2}' is not found.");
                return;
            }

            byte[] baseBuffer;
            byte[] file1Buffer;
            byte[] file2Buffer;
            try
            {
                baseBuffer = File.ReadAllBytes(@base);
                file1Buffer = File.ReadAllBytes(file1);
                file2Buffer = File.ReadAllBytes(file2);

            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return;
            }

            try
            {
                using (var baseMemoryStream = new MemoryStream(baseBuffer))
                using (var file1MemoryStream = new MemoryStream(file1Buffer))
                using (var file2MemoryStream = new MemoryStream(file2Buffer))
                using (var outMemoryStream = new MemoryStream(baseBuffer))
                using (var baseExcelPackage = new ExcelPackage(baseMemoryStream))
                using (var file1ExcelPackage = new ExcelPackage(file1MemoryStream))
                using (var file2ExcelPackage = new ExcelPackage(file2MemoryStream))
                using (var outExcelPackage = new ExcelPackage(outMemoryStream))
                {
                    var merges = new List<Tuple<MergeCellObject, ExcelWorksheet>>();
                    var conflictMerges = new List<Tuple<MergeCellObject, ExcelWorksheet>>();

                    foreach (var baseSheet in baseExcelPackage.Workbook.Worksheets)
                    {
                        var sheetName = baseSheet.Name;

                        var file1Sheet = file1ExcelPackage.Workbook.Worksheets[sheetName];
                        if (file1Sheet == null)
                            continue;
                        var file2Sheet = file2ExcelPackage.Workbook.Worksheets[sheetName];
                        if (file2Sheet == null)
                            continue;

                        var outSheet = outExcelPackage.Workbook.Worksheets[sheetName];


                        foreach (var rangeObject in ranges.Where(val => val.SheetName == sheetName && val.Address != null))
                        {
                            var startRow = rangeObject.Address.Start.Row;
                            var startCol = rangeObject.Address.Start.Column;
                            var endRow = rangeObject.Address.End.Row;
                            var endCol = rangeObject.Address.End.Column;

                            for (int row = startRow; row <= endRow; ++row)
                            {
                                for (int col = startCol; col <= endCol; ++col)
                                {
                                    var merge = MergeCellObject.CreateMergeCellObject(baseSheet, file1Sheet, file2Sheet, row, col);

                                    if (merge.MergeTarget == MergeTarget.Base)
                                        continue;

                                    if (merge.IsConflict)
                                        conflictMerges.Add(Tuple.Create(merge, outSheet));
                                    else
                                        merges.Add(Tuple.Create(merge, outSheet));
                                }
                            }
                        }
                    }

                    foreach (var merge in merges)
                    {
                        if (!merge.Item1.IsConflict)
                        {
                            merge.Item1.Merge(merge.Item2);
                            continue;
                        }
                    }

                    var count = conflictMerges.Count;
                    var num = 1;
                    foreach (var merge in conflictMerges)
                    {
                        Console.WriteLine($"conflict {num} / {count} [1 , 2 , b]");

                        while (true)
                        {
                            var readLine = Console.ReadLine();

                            if (readLine == "1")
                            {
                                merge.Item1.Merge(merge.Item2, MergeTarget.File1);
                                break;
                            }
                            else if (readLine == "2")
                            {
                                merge.Item1.Merge(merge.Item2, MergeTarget.File2);
                                break;
                            }
                            else if (readLine == "b")
                            {
                                break;
                            }

                            Console.WriteLine("Please type [1 . 2 , b]");
                        }

                        ++num;
                    }


                    const string outDir = "out";
                    Directory.CreateDirectory(outDir);
                    var filePath = Path.Combine(outDir, Path.GetFileName(@base));

                    using (var fileStream = new FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.None))
                    {
                        outExcelPackage.SaveAs(fileStream);
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return;
            }

        }
    }

    public class RangeObject
    {
        public string SheetName { get; }
        public ExcelAddress Address { get; }

        public static RangeObject Create(JObject jObject)
        {


            JToken jToken;
            string sheetName = null;
            ExcelAddress address = null;
            {
                if (jObject.TryGetValue("sheetName", out jToken) && jToken is JValue jValue)
                {
                    sheetName = jValue.Value as string;
                }
            }
            {
                if (jObject.TryGetValue("address", out jToken) && jToken is JValue jValue)
                {
                    var range = jValue.Value as string;
                    try
                    {
                        address = new ExcelAddress(range);
                    }
                    catch (Exception)
                    {
                        address = null;
                    }

                }
            }


            return new RangeObject(sheetName, address);
        }


        public RangeObject(string sheetName, ExcelAddress address)
        {
            SheetName = sheetName;
            Address = address;
        }
    }

    public enum MergeTarget
    {
        Base,
        File1,
        File2,
        Conflict
    }

    public class MergeCellObject
    {
        private readonly string _baseFormula;
        private readonly string _file1Formula;
        private readonly string _file2Formula;
        private readonly object _baseValue;
        private readonly object _file1Value;
        private readonly object _file2Value;
        private readonly int _row;
        private readonly int _col;
        public MergeTarget MergeTarget { get; private set; }


        public static MergeCellObject CreateMergeCellObject(ExcelWorksheet baseSheet, ExcelWorksheet file1Sheet, ExcelWorksheet file2Sheet, int row, int col)
        {
            var baseCell = baseSheet.Cells[row, col];
            var file1Cell = file1Sheet.Cells[row, col];
            var file2Cell = file2Sheet.Cells[row, col];

            var baseFormula = baseCell.Formula;
            var baseValue = baseCell.Value;
            var file1Formula = file1Cell.Formula;
            var file1Value = file1Cell.Value;
            var file2Formula = file2Cell.Formula;
            var file2Value = file2Cell.Value;

            var file1Diff = baseFormula != file1Formula || (baseValue != null ? file1Value != null ? !baseValue.Equals(file1Value) : true : file1Value != null);
            var file2Diff = baseFormula != file2Formula || (baseValue != null ? file2Value != null ? !baseValue.Equals(file2Value) : true : file2Value != null);


            return new MergeCellObject(baseFormula, file1Formula, file2Formula, baseValue, file1Value, file2Value, row, col)
            {
                IsConflict = file1Diff && file2Diff,
                MergeTarget = file1Diff && file2Diff ? MergeTarget.Conflict : file1Diff ? MergeTarget.File1 : file2Diff ? MergeTarget.File2 : MergeTarget.Base
            };
        }

        public MergeCellObject(string baseFormula, string file1Formula, string file2Formula, object baseValue, object file1Value, object file2Value, int row, int col)
        {
            _baseFormula = baseFormula;
            _file1Formula = file1Formula;
            _file2Formula = file2Formula;
            _baseValue = baseValue;
            _file1Value = file1Value;
            _file2Value = file2Value;
            _row = row;
            _col = col;
        }

        public bool IsConflict { get; private set; }

        public void Merge(ExcelWorksheet dest)
        {
            if (IsConflict)
                throw new InvalidOperationException();

            Merge(dest, MergeTarget);
        }

        public void Merge(ExcelWorksheet dest, MergeTarget target)
        {

            var destCell = dest.Cells[_row, _col];


            if (target == MergeTarget.File1)
            {
                destCell.Formula = _file1Formula;
                if (_file1Formula == "")
                    destCell.Value = _file1Value;

            }
            else if (target == MergeTarget.File2)
            {
                destCell.Formula = _file2Formula;
                if (_file2Formula == "")
                    destCell.Value = _file2Value;
            }
            else 
            {
                destCell.Formula = _baseFormula;
                if (_baseFormula == "")
                    destCell.Value = _baseValue;
            }
        }

    }
}
