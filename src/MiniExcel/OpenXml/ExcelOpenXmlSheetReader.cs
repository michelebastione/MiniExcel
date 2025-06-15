using MiniExcelLibs.OpenXml.Models;
using MiniExcelLibs.Utils;
using MiniExcelLibs.Zip;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml;

namespace MiniExcelLibs.OpenXml
{
    internal partial class ExcelOpenXmlSheetReader : IExcelReader
    {
        private bool _disposed;
        
        private static readonly string[] Ns = { Config.SpreadsheetmlXmlns, Config.SpreadsheetmlXmlStrictns };
        private static readonly string[] RelationshiopNs = { Config.SpreadsheetmlXmlRelationshipns, Config.SpreadsheetmlXmlStrictRelationshipns };
        
        private readonly OpenXmlConfiguration _config;
        private List<SheetRecord> _sheetRecords;
        private ExcelOpenXmlStyles _style;
        private MergeCells _mergeCells;
        
        internal readonly ExcelOpenXmlZip Archive;
        internal IDictionary<int, string> SharedStrings;

        private readonly XmlReaderSettings _xmlSettings = new XmlReaderSettings
        {
            IgnoreComments = true,
            IgnoreWhitespace = true,
            XmlResolver = null
        };

        private ExcelOpenXmlSheetReader(Stream stream, IConfiguration configuration, bool isUpdateMode, bool asynchronous = false)
        {
            Archive = new ExcelOpenXmlZip(stream);
            _config = (OpenXmlConfiguration)configuration ?? OpenXmlConfiguration.DefaultConfig;
            _xmlSettings.Async = asynchronous;
        }

        public static ExcelOpenXmlSheetReader Create(Stream stream, IConfiguration configuration, bool isUpdateMode = true)
        {
            var reader = new ExcelOpenXmlSheetReader(stream, configuration, isUpdateMode, false);
            reader.SetSharedStrings();

            return reader;
        }
        
        public IEnumerable<IDictionary<string, object>> Query(bool useHeaderRow, string sheetName, string startCell)
        {
            return QueryRange(useHeaderRow, sheetName, startCell, "");
        }

        public IEnumerable<T> Query<T>(string sheetName, string startCell, bool hasHeader) where T : class, new()
        {
            sheetName = sheetName ?? CustomPropertyHelper.GetExcellSheetInfo(typeof(T), _config)?.ExcelSheetName;
            return QueryImpl<T>(Query(false, sheetName, startCell), startCell, hasHeader, _config);
        }

        public IEnumerable<IDictionary<string, object>> QueryRange(bool useHeaderRow, string sheetName, string startCell, string endCell)
        {
            // convert to 0-based
            if (!ReferenceHelper.ParseReference(startCell, out var startColumnIndex, out var startRowIndex))
                throw new InvalidDataException($"Value {startCell} is not a valid cell reference.");
            
            startColumnIndex--;
            startRowIndex--;

            // endCell is allowed to be empty to query for all rows and columns
            int? endColumnIndex = null;
            int? endRowIndex = null;
            if (!string.IsNullOrWhiteSpace(endCell))
            {
                if (!ReferenceHelper.ParseReference(endCell, out int cIndex, out int rIndex))
                    throw new InvalidDataException($"Value {endCell} is not a valid cell reference.");

                // convert to 0-based
                endColumnIndex = cIndex - 1;
                endRowIndex = rIndex - 1;
            }

            return QueryRangeInternal(useHeaderRow, sheetName, startRowIndex, startColumnIndex, endRowIndex, endColumnIndex);
        }

        public IEnumerable<T> QueryRange<T>(string sheetName, string startCell, string endCell, bool hasHeader) where T : class, new()
        {
            return QueryImpl<T>(QueryRange(false, sheetName, startCell, endCell), startCell, hasHeader, _config);
        }

        public IEnumerable<IDictionary<string, object>> QueryRange(bool useHeaderRow, string sheetName, int startRowIndex, int startColumnIndex, int? endRowIndex, int? endColumnIndex)
        {
            if (startRowIndex <= 0)
                throw new ArgumentOutOfRangeException(nameof(startRowIndex), "Start row index is 1-based and must be greater than 0.");
            
            if (startColumnIndex <= 0)
                throw new ArgumentOutOfRangeException(nameof(startColumnIndex), "Start column index is 1-based and must be greater than 0.");
            
            // convert to 0-based
            startColumnIndex--;
            startRowIndex--;

            if (endRowIndex.HasValue)
            {
                if (endRowIndex.Value <= 0)
                    throw new ArgumentOutOfRangeException(nameof(endRowIndex), "End row index is 1-based and must be greater than 0.");
                
                // convert to 0-based
                endRowIndex--;
            }
            
            if (endColumnIndex.HasValue)
            {
                if (endColumnIndex.Value <= 0)
                    throw new ArgumentOutOfRangeException(nameof(endColumnIndex), "End column index is 1-based and must be greater than 0.");
                
                // convert to 0-based
                endColumnIndex--;
            }

            return QueryRangeInternal(useHeaderRow, sheetName, startRowIndex, startColumnIndex, endRowIndex, endColumnIndex);
        }

        public IEnumerable<T> QueryRange<T>(string sheetName, int startRowIndex, int startColumnIndex, int? endRowIndex, int? endColumnIndex, bool hasHeader) where T : class, new()
        {
            var cellValue = ReferenceHelper.ConvertXyToCell(startColumnIndex, startRowIndex);
            var rows = QueryRange(false, sheetName, startRowIndex, startColumnIndex, endRowIndex, endColumnIndex);
            
            return QueryImpl<T>(rows, cellValue, hasHeader, _config);
        }

        private IEnumerable<IDictionary<string, object>> QueryRangeInternal(bool useHeaderRow, string sheetName, int startRowIndex, int startColumnIndex, int? endRowIndex, int? endColumnIndex)
        {
            var sheetEntry = GetSheetEntry(sheetName);

            // TODO: need to optimize performance
            // Q. why need 3 times openstream merge one open read? A. no, zipstream can't use position = 0

            if (_config.FillMergedCells && !TryGetMergeCells(sheetEntry, _xmlSettings, out _mergeCells))
                yield break;

            if (!TryGetMaxRowColumnIndex(sheetEntry, _xmlSettings, out var withoutCr, out var maxRowIndex, out var maxColumnIndex))
                yield break;

            if (endColumnIndex.HasValue)
            {
                maxColumnIndex = endColumnIndex.Value;
            }

            using (var sheetStream = sheetEntry.Open())
            using (var reader = XmlReader.Create(sheetStream, _xmlSettings))
            {
                if (!XmlReaderHelper.IsStartElement(reader, "worksheet", Ns))
                    yield break;

                if (!XmlReaderHelper.ReadFirstContent(reader))
                    yield break;

                while (!reader.EOF)
                {
                    if (XmlReaderHelper.IsStartElement(reader, "sheetData", Ns))
                    {
                        if (!XmlReaderHelper.ReadFirstContent(reader))
                            continue;

                        var headRows = new Dictionary<int, string>();
                        int rowIndex = -1;
                        bool isFirstRow = true;
                        while (!reader.EOF)
                        {
                            if (XmlReaderHelper.IsStartElement(reader, "row", Ns))
                            {
                                var nextRowIndex = rowIndex + 1;
                                if (int.TryParse(reader.GetAttribute("r"), out int arValue))
                                    rowIndex = arValue - 1; // The row attribute is 1-based
                                else
                                    rowIndex++;

                                if (rowIndex < startRowIndex)
                                {
                                    XmlReaderHelper.ReadFirstContent(reader);
                                    XmlReaderHelper.SkipToNextSameLevelDom(reader);
                                    continue;
                                }
                                if (endRowIndex.HasValue && rowIndex > endRowIndex.Value)
                                {
                                    break;
                                }

                                // fill empty rows
                                if (!_config.IgnoreEmptyRows)
                                {
                                    var expectedRowIndex = isFirstRow ? startRowIndex : nextRowIndex;
                                    if (startRowIndex <= expectedRowIndex && expectedRowIndex < rowIndex)
                                    {
                                        for (int i = expectedRowIndex; i < rowIndex; i++)
                                        {
                                            yield return GetCell(useHeaderRow, maxColumnIndex, headRows, startColumnIndex);
                                        }
                                    }
                                }

                                // row -> c, must after `if (nextRowIndex < rowIndex)` condition code, eg. The first empty row has no xml element,and the second row xml element is <row r="2"/>
                                if (!XmlReaderHelper.ReadFirstContent(reader) && !_config.IgnoreEmptyRows)
                                {
                                    //Fill in case of self closed empty row tag eg. <row r="1"/>
                                    yield return GetCell(useHeaderRow, maxColumnIndex, headRows, startColumnIndex);
                                    continue;
                                }

                                #region Set Cells

                                var cell = GetCell(useHeaderRow, maxColumnIndex, headRows, startColumnIndex);
                                var columnIndex = withoutCr ? -1 : 0;
                                while (!reader.EOF)
                                {
                                    if (XmlReaderHelper.IsStartElement(reader, "c", Ns))
                                    {
                                        var aS = reader.GetAttribute("s");
                                        var aR = reader.GetAttribute("r");
                                        var aT = reader.GetAttribute("t");
                                        var cellValue = ReadCellAndSetColumnIndex(reader, ref columnIndex, withoutCr,
                                            startColumnIndex, aR, aT);

                                        if (_config.FillMergedCells)
                                        {
                                            if (_mergeCells.MergesValues.ContainsKey(aR))
                                            {
                                                _mergeCells.MergesValues[aR] = cellValue;
                                            }
                                            else if (_mergeCells.MergesMap.TryGetValue(aR, out var mergeKey))
                                            {
                                                _mergeCells.MergesValues.TryGetValue(mergeKey, out cellValue);
                                            }
                                        }

                                        if (columnIndex < startColumnIndex || (endColumnIndex.HasValue && columnIndex > endColumnIndex.Value))
                                            continue;

                                        if (!string.IsNullOrEmpty(aS)) // if c with s meaning is custom style need to check type by xl/style.xml
                                        {
                                            int xfIndex = -1;
                                            if (int.TryParse(aS, NumberStyles.Any, CultureInfo.InvariantCulture, out var styleIndex))
                                            {
                                                xfIndex = styleIndex;
                                            }

                                            // only when have s attribute then load styles xml data
                                            _style = _style ?? new ExcelOpenXmlStyles(Archive);

                                            cellValue = _style.ConvertValueByStyleFormat(xfIndex, cellValue);
                                        }

                                        SetCellsValueAndHeaders(cellValue, useHeaderRow, headRows, isFirstRow, cell, columnIndex);
                                    }
                                    else if (!XmlReaderHelper.SkipContent(reader))
                                    {
                                        break;
                                    }
                                }

                                #endregion

                                if (isFirstRow)
                                {
                                    isFirstRow = false; // for startcell logic
                                    if (useHeaderRow)
                                        continue;
                                }

                                yield return cell;
                            }
                            else if (!XmlReaderHelper.SkipContent(reader))
                            {
                                break;
                            }
                        }
                    }
                    else if (!XmlReaderHelper.SkipContent(reader))
                    {
                        break;
                    }
                }
            }
        }

        //Todo: Find a way if possible to remove the 'hasHeader' parameter to check whether or not to include
        // the first row in the result set in favor of modifying the already present 'useHeaderRow' to do the same job    
        public static IEnumerable<T> QueryImpl<T>(IEnumerable<IDictionary<string, object>> values, string startCell, bool hasHeader, Configuration configuration) where T : class, new()
        {
            var type = typeof(T);

            //TODO:need to optimize
            List<ExcelColumnInfo> props = null;
            Dictionary<string, int> headersDic = null;
            string[] keys = null;
            var first = true;
            var rowIndex = 0;

            foreach (var item in values)
            {
                if (first)
                {
                    keys = item.Keys.ToArray();
                    var trimColumnNames = (configuration as OpenXmlConfiguration)?.TrimColumnNames ?? false;
                    headersDic = CustomPropertyHelper.GetHeaders(item, trimColumnNames);

                    //TODO: alert don't duplicate column name
                    props = CustomPropertyHelper.GetExcelCustomPropertyInfos(type, keys, configuration);
                    first = false;

                    if (hasHeader)
                        continue;
                }

                var v = new T();
                foreach (var pInfo in props)
                {
                    if (pInfo.ExcelColumnAliases != null)
                    {
                        foreach (var alias in pInfo.ExcelColumnAliases)
                        {
                            if (headersDic.TryGetValue(alias, out var columnId))
                            {
                                var columnName = keys[columnId];
                                item.TryGetValue(columnName, out var aliasItemValue);

                                if (aliasItemValue != null)
                                {
                                    var newAliasValue = TypeHelper.TypeMapping(v, pInfo, aliasItemValue, rowIndex, startCell, configuration);
                                }
                            }
                        }
                    }

                    //Q: Why need to check every time? A: it needs to check everytime, because it's dictionary
                    object itemValue = null;
                    if (pInfo.ExcelIndexName != null && keys.Contains(pInfo.ExcelIndexName))
                    {
                        item.TryGetValue(pInfo.ExcelIndexName, out itemValue);
                    }
                    else if (headersDic.TryGetValue(pInfo.ExcelColumnName, out var columnId))
                    {
                        var columnName = keys[columnId];
                        item.TryGetValue(columnName, out itemValue);
                    }

                    if (itemValue != null)
                    {
                        var newValue = TypeHelper.TypeMapping(v, pInfo, itemValue, rowIndex, startCell, configuration);
                    }
                }
                rowIndex++;
                yield return v;
            }
        }
        
        private ZipArchiveEntry GetSheetEntry(string sheetName)
        {
            // if sheets count > 1 need to read xl/_rels/workbook.xml.rels
            var sheets = Archive.entries
                .Where(w => w.FullName.StartsWith("xl/worksheets/sheet", StringComparison.OrdinalIgnoreCase) ||
                            w.FullName.StartsWith("/xl/worksheets/sheet", StringComparison.OrdinalIgnoreCase))
                .ToArray();
            
            ZipArchiveEntry sheetEntry;
            if (sheetName != null)
            {
                SetWorkbookRels(Archive.entries);
                var sheetRecord = _sheetRecords.SingleOrDefault(s => s.Name == sheetName);
                if (sheetRecord == null)
                {
                    if (_config.DynamicSheets == null)
                        throw new InvalidOperationException("Please check that parameters sheetName/Index are correct");

                    var sheetConfig = _config.DynamicSheets.FirstOrDefault(ds => ds.Key == sheetName);
                    if (sheetConfig != null)
                    {
                        sheetRecord = _sheetRecords.SingleOrDefault(s => s.Name == sheetConfig.Name);
                    }
                }
                sheetEntry = sheets.Single(w => w.FullName == $"xl/{sheetRecord?.Path}" || 
                                                w.FullName == $"/xl/{sheetRecord?.Path}" || 
                                                w.FullName == sheetRecord?.Path || 
                                                sheetRecord?.Path == $"/{w.FullName}");
            }
            else if (sheets.Length > 1)
            {
                SetWorkbookRels(Archive.entries);
                var s = _sheetRecords[0];
                sheetEntry = sheets.Single(w => w.FullName == $"xl/{s.Path}" || 
                                                w.FullName == $"/xl/{s.Path}" || 
                                                w.FullName.TrimStart('/') == s.Path.TrimStart('/'));
            }
            else
            {
                sheetEntry = sheets.Single();
            }

            return sheetEntry;
        }

        private static IDictionary<string, object> GetCell(bool useHeaderRow, int maxColumnIndex, Dictionary<int, string> headRows, int startColumnIndex)
        {
            return useHeaderRow 
                ? CustomPropertyHelper.GetEmptyExpandoObject(headRows) 
                : CustomPropertyHelper.GetEmptyExpandoObject(maxColumnIndex, startColumnIndex);
        }

        private static void SetCellsValueAndHeaders(object cellValue, bool useHeaderRow, Dictionary<int, string> headRows, bool isFirstRow, IDictionary<string, object> cell, int columnIndex)
        {
            if (!useHeaderRow)
            {
                //if not using First Head then using A,B,C as index
                cell[ColumnHelper.GetAlphabetColumnName(columnIndex)] = cellValue;
                return;
            }

            if (isFirstRow) // for startcell logic
            {
                var cellValueString = cellValue?.ToString();
                if (!string.IsNullOrWhiteSpace(cellValueString))
                    headRows.Add(columnIndex, cellValueString);
            }
            else if (headRows.TryGetValue(columnIndex, out var key))
            {
                cell[key] = cellValue;
            }
        }

        private void SetSharedStrings()
        {
            if (SharedStrings != null)
                return;
            
            var sharedStringsEntry = Archive.GetEntry("xl/sharedStrings.xml");
            if (sharedStringsEntry == null)
                return;
            
            using (var stream = sharedStringsEntry.Open())
            {
                var idx = 0;
                if (_config.EnableSharedStringCache && sharedStringsEntry.Length >= _config.SharedStringCacheSize)
                {
                    SharedStrings = new SharedStringsDiskCache();
                    foreach (var sharedString in XmlReaderHelper.GetSharedStrings(stream, Ns))
                        SharedStrings[idx++] = sharedString;
                }
                else if (SharedStrings == null)
                {
                    SharedStrings = XmlReaderHelper.GetSharedStrings(stream, Ns).ToDictionary(x => idx++, x => x);
                }
            }
        }

        private void SetWorkbookRels(ReadOnlyCollection<ZipArchiveEntry> entries)
        {
            _sheetRecords = _sheetRecords ?? GetWorkbookRels(entries);
        }

        private static IEnumerable<SheetRecord> ReadWorkbook(ReadOnlyCollection<ZipArchiveEntry> entries, XmlReaderSettings settings)
        {
            using (var stream = entries.Single(w => w.FullName == "xl/workbook.xml").Open())
            using (var reader = XmlReader.Create(stream, settings))
            {
                if (!XmlReaderHelper.IsStartElement(reader, "workbook", Ns))
                    yield break;

                if (!XmlReaderHelper.ReadFirstContent(reader))
                    yield break;

                var activeSheetIndex = 0;
                while (!reader.EOF)
                {
                    if (XmlReaderHelper.IsStartElement(reader, "bookViews", Ns))
                    {
                        if (XmlReaderHelper.ReadFirstContent(reader))
                        {
                            while (!reader.EOF)
                            {
                                if (XmlReaderHelper.IsStartElement(reader, "workbookView", Ns))
                                {
                                    var activeSheet = reader.GetAttribute("activeTab");
                                    if (int.TryParse(activeSheet, out var index))
                                    {
                                        activeSheetIndex = index;
                                    }

                                    reader.Skip();
                                }
                                else if (!XmlReaderHelper.SkipContent(reader))
                                {
                                    break;
                                }
                            }
                        }
                    }
                    else if (XmlReaderHelper.IsStartElement(reader, "sheets", Ns))
                    {
                        if (!XmlReaderHelper.ReadFirstContent(reader))
                            continue;

                        var sheetCount = 0;
                        while (!reader.EOF)
                        {
                            if (XmlReaderHelper.IsStartElement(reader, "sheet", Ns))
                            {
                                yield return new SheetRecord(
                                    reader.GetAttribute("name"),
                                    reader.GetAttribute("state"),
                                    uint.Parse(reader.GetAttribute("sheetId")),
                                    XmlReaderHelper.GetAttribute(reader, "id", RelationshiopNs),
                                    sheetCount == activeSheetIndex
                                );
                                
                                sheetCount++;
                                reader.Skip();
                            }
                            else if (!XmlReaderHelper.SkipContent(reader))
                            {
                                break;
                            }
                        }
                    }
                    else if (!XmlReaderHelper.SkipContent(reader))
                    {
                        yield break;
                    }
                }
            }
        }

        internal List<SheetRecord> GetWorkbookRels(ReadOnlyCollection<ZipArchiveEntry> entries)
        {
            var sheetRecords = ReadWorkbook(entries, _xmlSettings).ToList();

            using (var stream = entries.Single(w => w.FullName == "xl/_rels/workbook.xml.rels").Open())
            using (var reader = XmlReader.Create(stream, _xmlSettings))
            {
                if (!XmlReaderHelper.IsStartElement(reader, "Relationships", "http://schemas.openxmlformats.org/package/2006/relationships"))
                    return null;

                if (!XmlReaderHelper.ReadFirstContent(reader))
                    return null;

                while (!reader.EOF)
                {
                    if (XmlReaderHelper.IsStartElement(reader, "Relationship", "http://schemas.openxmlformats.org/package/2006/relationships"))
                    {
                        var rid = reader.GetAttribute("Id");
                        foreach (var sheet in sheetRecords.Where(sheet => sheet.Rid == rid))
                        {
                            sheet.Path = reader.GetAttribute("Target");
                            break;
                        }

                        reader.Skip();
                    }
                    else if (!XmlReaderHelper.SkipContent(reader))
                    {
                        break;
                    }
                }
            }

            return sheetRecords;
        }

        private object ReadCellAndSetColumnIndex(XmlReader reader, ref int columnIndex, bool withoutCr, int startColumnIndex, string aR, string aT)
        {
            const int xfIndex = -1;
            int newColumnIndex;

            if (withoutCr)
                newColumnIndex = columnIndex + 1;

            //TODO:need to check only need nextColumnIndex or columnIndex
            else if (ReferenceHelper.ParseReference(aR, out int referenceColumn, out _))
                newColumnIndex = referenceColumn - 1; // ParseReference is 1-based
            else
                newColumnIndex = columnIndex;

            columnIndex = newColumnIndex;

            if (columnIndex < startColumnIndex)
            {
                if (XmlReaderHelper.ReadFirstContent(reader))
                {
                    while (!reader.EOF)
                    {
                        if (!XmlReaderHelper.SkipContent(reader))
                            break;
                    }
                }

                return null;
            }

            if (!XmlReaderHelper.ReadFirstContent(reader))
                return null;

            object value = null;
            while (!reader.EOF)
            {
                if (XmlReaderHelper.IsStartElement(reader, "v", Ns))
                {
                    var rawValue = reader.ReadElementContentAsString();
                    if (!string.IsNullOrEmpty(rawValue))
                        ConvertCellValue(rawValue, aT, xfIndex, out value);
                }
                else if (XmlReaderHelper.IsStartElement(reader, "is", Ns))
                {
                    var rawValue = StringHelper.ReadStringItem(reader);
                    if (!string.IsNullOrEmpty(rawValue))
                        ConvertCellValue(rawValue, aT, xfIndex, out value);
                }
                else if (!XmlReaderHelper.SkipContent(reader))
                {
                    break;
                }
            }

            return value;
        }

        private void ConvertCellValue(string rawValue, string aT, int xfIndex, out object value)
        {
            const NumberStyles style = NumberStyles.Any;
            var invariantCulture = CultureInfo.InvariantCulture;

            switch (aT)
            {
                case "s":
                    if (int.TryParse(rawValue, style, invariantCulture, out var sstIndex))
                    {
                        if (sstIndex >= 0 && sstIndex < SharedStrings.Count)
                        {
                            //value = Helpers.ConvertEscapeChars(_SharedStrings[sstIndex]);
                            value = XmlEncoder.DecodeString(SharedStrings[sstIndex]);
                            return;
                        }
                    }
                    value = null;
                    return;

                case "inlineStr":
                case "str":
                    //TODO: it will unbox,box
                    var v = XmlEncoder.DecodeString(rawValue);
                    if (_config.EnableConvertByteArray)
                    {
                        //if str start with "data:image/png;base64," then convert to byte[] https://github.com/mini-software/MiniExcel/issues/318
                        if (v != null && v.StartsWith("@@@fileid@@@,", StringComparison.Ordinal))
                        {
                            var path = v.Substring(13);
                            var entry = Archive.GetEntry(path);
                            var bytes = new byte[entry.Length];

                            using (var stream = entry.Open())
                            using (var ms = new MemoryStream(bytes))
                            {
                                stream.CopyTo(ms);
                            }
                            value = bytes;
                        }
                        else
                        {
                            value = v;
                        }
                    }
                    else
                    {
                        value = v;
                    }
                    return;

                case "b":
                    value = rawValue == "1";
                    return;

                case "d":
                    if (DateTime.TryParseExact(rawValue, "yyyy-MM-dd", invariantCulture, DateTimeStyles.AllowLeadingWhite | DateTimeStyles.AllowTrailingWhite, out var date))
                    {
                        value = date;
                        return;
                    }

                    value = rawValue;
                    return;

                case "e":
                    value = rawValue;
                    return;

                default:
                    if (double.TryParse(rawValue, style, invariantCulture, out var n))
                    {
                        value = n;
                        return;
                    }

                    value = rawValue;
                    return;
            }
        }

        internal IList<ExcelRange> GetDimensions()
        {
            var ranges = new List<ExcelRange>();

            var sheets = Archive.entries.Where(e =>
                e.FullName.StartsWith("xl/worksheets/sheet", StringComparison.OrdinalIgnoreCase) ||
                e.FullName.StartsWith("/xl/worksheets/sheet", StringComparison.OrdinalIgnoreCase));

            foreach (var sheet in sheets)
            {
                var maxRowIndex = -1;
                var maxColumnIndex = -1;

                string startCell = null;
                string endCell = null;

                var withoutCr = false;

                using (var sheetStream = sheet.Open())
                using (var reader = XmlReader.Create(sheetStream, _xmlSettings))
                {
                    while (reader.Read())
                    {
                        if (XmlReaderHelper.IsStartElement(reader, "c", Ns))
                        {
                            var r = reader.GetAttribute("r");
                            if (r != null)
                            {
                                if (ReferenceHelper.ParseReference(r, out var column, out var row))
                                {
                                    column--;
                                    row--;
                                    maxRowIndex = Math.Max(maxRowIndex, row);
                                    maxColumnIndex = Math.Max(maxColumnIndex, column);
                                }
                            }
                            else
                            {
                                withoutCr = true;
                                break;
                            }
                        }

                        else if (XmlReaderHelper.IsStartElement(reader, "dimension", Ns))
                        {
                            var refAttr = reader.GetAttribute("ref");
                            if (string.IsNullOrEmpty(refAttr))
                                throw new InvalidDataException("No dimension data found for the sheet");

                            var rs = refAttr.Split(':');
                            if (!ReferenceHelper.ParseReference(rs.Length == 2 ? rs[1] : rs[0], out var col, out var row))
                                throw new InvalidDataException("The dimensions of the sheet are invalid");

                            maxColumnIndex = col;
                            maxRowIndex = row;

                            startCell = rs[0];
                            endCell = rs[1];

                            break;
                        }
                    }
                }

                if (withoutCr)
                {
                    using (var sheetStream = sheet.Open())
                    using (var reader = XmlReader.Create(sheetStream, _xmlSettings))
                    {
                        if (!XmlReaderHelper.IsStartElement(reader, "worksheet", Ns))
                            throw new InvalidDataException("No worksheet data found for the sheet");

                        if (!XmlReaderHelper.ReadFirstContent(reader))
                            throw new InvalidOperationException("Excel sheet does not contain any data");

                        while (!reader.EOF)
                        {
                            if (XmlReaderHelper.IsStartElement(reader, "sheetData", Ns))
                            {
                                if (XmlReaderHelper.ReadFirstContent(reader))
                                {
                                    while (!reader.EOF)
                                    {
                                        if (XmlReaderHelper.IsStartElement(reader, "row", Ns))
                                        {
                                            maxRowIndex++;

                                            if (XmlReaderHelper.ReadFirstContent(reader))
                                            {
                                                var cellIndex = -1;
                                                while (!reader.EOF)
                                                {
                                                    if (XmlReaderHelper.IsStartElement(reader, "c", Ns))
                                                    {
                                                        cellIndex++;
                                                        maxColumnIndex = Math.Max(maxColumnIndex, cellIndex);
                                                    }

                                                    if (!XmlReaderHelper.SkipContent(reader))
                                                        break;
                                                }
                                            }
                                        }
                                        else if (!XmlReaderHelper.SkipContent(reader))
                                        {
                                            break;
                                        }
                                    }
                                }
                            }
                            else if (!XmlReaderHelper.SkipContent(reader))
                            {
                                break;
                            }
                        }
                    }
                }

                var range = new ExcelRange(maxRowIndex, maxColumnIndex)
                {
                    StartCell = startCell,
                    EndCell = endCell
                };
                ranges.Add(range);
            }

            return ranges;
        }

        private static bool TryGetMaxRowColumnIndex(ZipArchiveEntry sheetEntry, XmlReaderSettings settings, out bool withoutCr, out int maxRowIndex, out int maxColumnIndex)
        {
            withoutCr = false;
            maxRowIndex = -1;
            maxColumnIndex = -1;
            
            using (var sheetStream = sheetEntry.Open())
            using (var reader = XmlReader.Create(sheetStream, settings))
            {
                while (reader.Read())
                {
                    if (XmlReaderHelper.IsStartElement(reader, "c", Ns))
                    {
                        var r = reader.GetAttribute("r");
                        if (r != null)
                        {
                            if (ReferenceHelper.ParseReference(r, out var column, out var row))
                            {
                                column--;
                                row--;
                                maxRowIndex = Math.Max(maxRowIndex, row);
                                maxColumnIndex = Math.Max(maxColumnIndex, column);
                            }
                        }
                        else
                        {
                            withoutCr = true;
                            break;
                        }
                    }
                    //this method logic depends on dimension to get maxcolumnIndex, if without dimension then it need to foreach all rows first time to get maxColumn and maxRowColumn
                    else if (XmlReaderHelper.IsStartElement(reader, "dimension", Ns))
                    {
                        var refAttr = reader.GetAttribute("ref");
                        if (string.IsNullOrEmpty(refAttr))
                            throw new InvalidDataException("No dimension data found for the sheet");

                        var rs = refAttr.Split(':');

                        // issue : https://github.com/mini-software/MiniExcel/issues/102
                        if (!ReferenceHelper.ParseReference(rs.Length == 2 ? rs[1] : rs[0], out int cIndex, out int rIndex))
                            throw new InvalidDataException("The dimensions of the sheet are invalid");

                        maxRowIndex = rIndex - 1;
                        maxColumnIndex = cIndex - 1;
                        break;
                    }
                }
            }

            if (!withoutCr) 
                return true;
            
            using (var sheetStream = sheetEntry.Open())
            using (var reader = XmlReader.Create(sheetStream, settings))
            {
                if (!XmlReaderHelper.IsStartElement(reader, "worksheet", Ns))
                    return false;

                if (!XmlReaderHelper.ReadFirstContent(reader))
                    return false;

                while (!reader.EOF)
                {
                    if (XmlReaderHelper.IsStartElement(reader, "sheetData", Ns))
                    {
                        if (XmlReaderHelper.ReadFirstContent(reader))
                        {
                            while (!reader.EOF)
                            {
                                if (XmlReaderHelper.IsStartElement(reader, "row", Ns))
                                {
                                    maxRowIndex++;

                                    if (XmlReaderHelper.ReadFirstContent(reader))
                                    {
                                        var cellIndex = -1;
                                        while (!reader.EOF)
                                        {
                                            if (XmlReaderHelper.IsStartElement(reader, "c", Ns))
                                            {
                                                cellIndex++;
                                                maxColumnIndex = Math.Max(maxColumnIndex, cellIndex);
                                            }

                                            if (!XmlReaderHelper.SkipContent(reader))
                                                break;
                                        }
                                    }
                                }
                                else if (!XmlReaderHelper.SkipContent(reader))
                                {
                                    break;
                                }
                            }
                        }
                    }
                    else if (!XmlReaderHelper.SkipContent(reader))
                    {
                        break;
                    }
                }
            }

            return true;
        }

        private static bool TryGetMergeCells(ZipArchiveEntry sheetEntry, XmlReaderSettings settings, out MergeCells mergeCells)
        {
            mergeCells = new MergeCells();
            
            using (var sheetStream = sheetEntry.Open())
            using (var reader = XmlReader.Create(sheetStream, settings))
            {
                if (!XmlReaderHelper.IsStartElement(reader, "worksheet", Ns))
                    return false;
                
                while (reader.Read())
                {
                    if (!XmlReaderHelper.IsStartElement(reader, "mergeCells", Ns))
                        continue;

                    if (!XmlReaderHelper.ReadFirstContent(reader))
                        return false;

                    while (!reader.EOF)
                    {
                        if (XmlReaderHelper.IsStartElement(reader, "mergeCell", Ns))
                        {
                            var refAttr = reader.GetAttribute("ref");
                            var refs = refAttr.Split(':');
                            if (refs.Length == 1)
                                continue;

                            ReferenceHelper.ParseReference(refs[0], out var x1, out var y1);
                            ReferenceHelper.ParseReference(refs[1], out var x2, out var y2);

                            mergeCells.MergesValues.Add(refs[0], null);

                            // foreach range
                            var isFirst = true;
                            for (int x = x1; x <= x2; x++)
                            {
                                for (int y = y1; y <= y2; y++)
                                {
                                    if (!isFirst)
                                    {
                                        mergeCells.MergesMap.Add(ReferenceHelper.ConvertXyToCell(x, y), refs[0]);
                                    }

                                    isFirst = false;
                                }
                            }

                            XmlReaderHelper.SkipContent(reader);
                        }
                        else if (!XmlReaderHelper.SkipContent(reader))
                        {
                            break;
                        }
                    }
                }
                return true;
            }
        }

        ~ExcelOpenXmlSheetReader()
        {
            Dispose(false);
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    if (SharedStrings is SharedStringsDiskCache cache)
                    {
                        cache.Dispose();
                    }
                }

                _disposed = true;
            }
        }
    }
}