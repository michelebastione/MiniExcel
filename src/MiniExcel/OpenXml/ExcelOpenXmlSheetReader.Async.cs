using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using MiniExcelLibs.Utils;

namespace MiniExcelLibs.OpenXml
{
    internal partial class ExcelOpenXmlSheetReader
    {
        public async Task<IEnumerable<IDictionary<string, object>>> QueryAsync(bool useHeaderRow, string sheetName, string startCell, CancellationToken cancellationToken = default)
        {
            return await Task.Run(() => Query(useHeaderRow, sheetName, startCell), cancellationToken).ConfigureAwait(false);
        }

        public async Task<IEnumerable<T>> QueryAsync<T>(string sheetName, string startCell, bool hasHeader = true, CancellationToken cancellationToken = default) where T : class, new()
        {
            return await Task.Run(() => Query<T>(sheetName, startCell, hasHeader), cancellationToken).ConfigureAwait(false);
        }

        public async Task<IEnumerable<IDictionary<string, object>>> QueryRangeAsync(bool useHeaderRow, string sheetName, string startCell, string endCell, CancellationToken cancellationToken = default)
        {
            return await Task.Run(() => Query(useHeaderRow, sheetName, startCell), cancellationToken).ConfigureAwait(false);
        }

        public async Task<IEnumerable<T>> QueryRangeAsync<T>(string sheetName, string startCell, string endCell, bool hasHeader = true, CancellationToken cancellationToken = default) where T : class, new()
        {
            return await Task.Run(() => QueryRange<T>(sheetName, startCell, endCell, hasHeader), cancellationToken).ConfigureAwait(false);
        }
               
        public async Task<IEnumerable<IDictionary<string, object>>> QueryRangeAsync(bool useHeaderRow, string sheetName, int startRowIndex, int startColumnIndex, int? endRowIndex, int? endColumnIndex, CancellationToken cancellationToken = default)
        {
            return await Task.Run(() => QueryRange(useHeaderRow, sheetName, startRowIndex, startColumnIndex, endRowIndex, endColumnIndex), cancellationToken).ConfigureAwait(false);
        }
        
        public async Task<IEnumerable<T>> QueryRangeAsync<T>(string sheetName, int startRowIndex, int startColumnIndex, int? endRowIndex, int? endColumnIndex, bool hasHeader = true, CancellationToken cancellationToken = default) where T : class, new()
        {
            return await Task.Run(() => QueryRange<T>(sheetName, startRowIndex, startColumnIndex, endRowIndex, endColumnIndex, hasHeader), cancellationToken).ConfigureAwait(false);
        }
        
        
#if NETCOREAPP3_0_OR_GREATER
        public IAsyncEnumerable<T> EnumerateAsync<T>(string sheetName, string startCell, bool hasHeader, CancellationToken cancellationToken = default) where T : class, new()
        {
            sheetName ??= CustomPropertyHelper.GetExcellSheetInfo(typeof(T), _config)?.ExcelSheetName;
            var query = EnumerateAsync(false, sheetName, startCell, cancellationToken).ConfigureAwait(false);
            return EnumerateImplAsync<T>(query, startCell, hasHeader, _config, cancellationToken);
        }

        public IAsyncEnumerable<IDictionary<string, object>> EnumerateAsync(bool useHeaderRow, string sheetName, string startCell, CancellationToken cancellationToken = default)
        {
            return EnumerateRangeAsync(useHeaderRow, sheetName, startCell, "", cancellationToken);
        }

        public IAsyncEnumerable<IDictionary<string, object>> EnumerateRangeAsync(bool useHeaderRow, string sheetName, string startCell, string endCell, CancellationToken cancellationToken = default)
        {
            // convert to 0-based
            if (!ReferenceHelper.ParseReference(startCell, out var startColumnIndex, out var startRowIndex))
                throw new InvalidDataException($"Value {startCell} is not a valid cell reference.");
            
            startRowIndex--;
            startColumnIndex--;

            // endCell is allowed to be empty to query for all rows and columns
            int? endColumnIndex = null;
            int? endRowIndex = null;
            if (!string.IsNullOrWhiteSpace(endCell))
            {
                if (!ReferenceHelper.ParseReference(endCell, out int cIndex, out int rIndex))
                    throw new InvalidDataException($"Value {endCell} is not a valid cell reference.");

                // convert to 0-based
                endRowIndex = rIndex - 1;
                endColumnIndex = cIndex - 1;
            }

            return EnumerateRangeAsyncInternal(useHeaderRow, sheetName, startRowIndex, startColumnIndex, endRowIndex, endColumnIndex, cancellationToken);
        }

        public IAsyncEnumerable<T> EnumerateRangeAsync<T>(string sheetName, string startCell, string endCell, bool hasHeader, CancellationToken cancellationToken = default) where T : class, new()
        {
            sheetName ??= CustomPropertyHelper.GetExcellSheetInfo(typeof(T), _config)?.ExcelSheetName;
            var query = EnumerateRangeAsync(false, sheetName, startCell, endCell, cancellationToken).ConfigureAwait(false);
            
            return EnumerateImplAsync<T>(query, startCell, hasHeader, _config, cancellationToken);
        }
        
        public IAsyncEnumerable<IDictionary<string, object>> EnumerateRangeAsync(bool useHeaderRow, string sheetName, int startRowIndex, int startColumnIndex, int? endRowIndex, int? endColumnIndex, CancellationToken cancellationToken = default)
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

            return EnumerateRangeAsyncInternal(useHeaderRow, sheetName, startRowIndex, startColumnIndex, endRowIndex, endColumnIndex, cancellationToken);
        }
        
        public IAsyncEnumerable<T> EnumerateRangeAsync<T>(string sheetName, int startRowIndex, int startColumnIndex, int? endRowIndex, int? endColumnIndex, bool hasHeader, CancellationToken cancellationToken = default) where T : class, new()
        {
            var cellValue = ReferenceHelper.ConvertXyToCell(startColumnIndex, startRowIndex);
            var query = EnumerateRangeAsync(false, sheetName, startRowIndex, startColumnIndex, endRowIndex, endColumnIndex, cancellationToken).ConfigureAwait(false);
            
            return EnumerateImplAsync<T>(query, cellValue, hasHeader, _config, cancellationToken);
        }
        
        private async IAsyncEnumerable<IDictionary<string, object>> EnumerateRangeAsyncInternal(bool useHeaderRow, string sheetName, int startRowIndex, int startColumnIndex, int? endRowIndex, int? endColumnIndex, [EnumeratorCancellation] CancellationToken cancellationToken = default)
        {
            var sheetEntry = await GetSheetEntryAsync(sheetName);

            // TODO: need to optimize performance
            // Q. why need 3 times openstream merge one open read? A. no, zipstream can't use position = 0

            if (_config.FillMergedCells)
            {
                var mergeCells = await TryGetMergeCellsAsync(sheetEntry, _xmlSettings);
                if (mergeCells.Item1)
                {
                    _mergeCells = mergeCells.Item2;
                }
                else
                {
                    yield break;
                }
            }

            var maxRowColIndexResult = await TryGetMaxRowColumnIndexAsync(sheetEntry, _xmlSettings); 
            if (!maxRowColIndexResult.Item1)
                yield break;

            var (maxRowIndex, maxColumnIndex, withoutCr) = maxRowColIndexResult.Item2!.Value;
            if (endColumnIndex.HasValue)
            {
                maxColumnIndex = endColumnIndex.Value;
            }

            await using var sheetStream = sheetEntry.Open();
            using var reader = XmlReader.Create(sheetStream, _xmlSettings);
            
            if (!XmlReaderHelper.IsStartElement(reader, "worksheet", Ns))
                yield break;

            if (!await XmlReaderHelper.ReadFirstContentAsync(reader))
                yield break;

            while (!reader.EOF)
            {
                cancellationToken.ThrowIfCancellationRequested();
                
                if (XmlReaderHelper.IsStartElement(reader, "sheetData", Ns))
                {
                    if (!await XmlReaderHelper.ReadFirstContentAsync(reader))
                        continue;

                    var headRows = new Dictionary<int, string>();
                    int rowIndex = -1;
                    bool isFirstRow = true;
                    while (!reader.EOF)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
        
                        if (XmlReaderHelper.IsStartElement(reader, "row", Ns))
                        {
                            var nextRowIndex = rowIndex + 1;
                            if (int.TryParse(reader.GetAttribute("r"), out int arValue))
                                rowIndex = arValue - 1; // The row attribute is 1-based
                            else
                                rowIndex++;

                            if (rowIndex < startRowIndex)
                            {
                                await XmlReaderHelper.ReadFirstContentAsync(reader);
                                await XmlReaderHelper.SkipToNextSameLevelDomAsync(reader);
                                continue;
                            }
                            if (rowIndex > endRowIndex)
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
                            if (!await XmlReaderHelper.ReadFirstContentAsync(reader) && !_config.IgnoreEmptyRows)
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
                                cancellationToken.ThrowIfCancellationRequested();
                                
                                if (XmlReaderHelper.IsStartElement(reader, "c", Ns))
                                {
                                    var aS = reader.GetAttribute("s");
                                    var aR = reader.GetAttribute("r");
                                    var aT = reader.GetAttribute("t");
                                    (var cellValue, columnIndex) = await ReadCellAndSetColumnIndexAsync(reader, columnIndex, withoutCr, startColumnIndex, aR, aT);

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

                                    if (columnIndex < startColumnIndex || columnIndex > endColumnIndex)
                                        continue;

                                    if (!string.IsNullOrEmpty(aS)) // if c with s meaning is custom style need to check type by xl/style.xml
                                    {
                                        int xfIndex = -1;
                                        if (int.TryParse(aS, NumberStyles.Any, CultureInfo.InvariantCulture, out var styleIndex))
                                        {
                                            xfIndex = styleIndex;
                                        }

                                        // only when have s attribute then load styles xml data
                                        _style ??= new ExcelOpenXmlStyles(Archive);

                                        cellValue = _style.ConvertValueByStyleFormat(xfIndex, cellValue);
                                    }

                                    SetCellsValueAndHeaders(cellValue, useHeaderRow, headRows, isFirstRow, cell, columnIndex);
                                }
                                else if (!await XmlReaderHelper.SkipContentAsync(reader))
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
                        else if (!await XmlReaderHelper.SkipContentAsync(reader))
                        {
                            break;
                        }
                    }
                }
                else if (!await XmlReaderHelper.SkipContentAsync(reader))
                {
                    break;
                }
            }
        }

        internal static async IAsyncEnumerable<T> EnumerateImplAsync<T>(ConfiguredCancelableAsyncEnumerable<IDictionary<string, object>> values, string startCell, bool hasHeader, Configuration configuration, [EnumeratorCancellation] CancellationToken cancellationToken = default) where T : class, new()
        {
            var type = typeof(T);

            List<ExcelColumnInfo> props = null;
            Dictionary<string, int> headersDic = null;
            string[] keys = null;
            var first = true;
            var rowIndex = 0;
            
            await foreach (var item in values)
            {
                if (first)
                {
                    keys = item.Keys.ToArray();
                    var trimColumnNames = (configuration as OpenXmlConfiguration)?.TrimColumnNames ?? false;
                    headersDic = CustomPropertyHelper.GetHeaders(item, trimColumnNames);
                    
                    //TODO: alert don't duplicate column name
                    props = CustomPropertyHelper.GetExcelCustomPropertyInfos(type, keys, configuration);
                    first = false;
                    continue;
                }
                var v = new T();
                foreach (var pInfo in props)
                {
                    if (pInfo.ExcelColumnAliases != null)
                    {
                        foreach (var alias in pInfo.ExcelColumnAliases)
                        {
                            if (!headersDic.TryGetValue(alias, out var columnId)) 
                                continue;
                            
                            object newV = null;
                            var columnName = keys[columnId];
                            item.TryGetValue(columnName, out var itemValue);

                            if (itemValue == null)
                                continue;

                            newV = TypeHelper.TypeMapping(v, pInfo, itemValue, rowIndex, startCell, configuration);
                        }
                    }

                    //Q: Why need to check every time? A: it needs to check everytime, because it's dictionary
                    {
                        object newV = null;
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

                        if (itemValue == null)
                            continue;
                        
                        newV = TypeHelper.TypeMapping(v, pInfo, itemValue, rowIndex, startCell, configuration);
                    }
                }
                rowIndex++;
                yield return v;
            }
        }
        
        private async Task<ZipArchiveEntry> GetSheetEntryAsync(string sheetName)
        {
            // if sheets count > 1 need to read xl/_rels/workbook.xml.rels
            var sheets = Archive.entries
                .Where(w => w.FullName.StartsWith("xl/worksheets/sheet", StringComparison.OrdinalIgnoreCase) ||
                            w.FullName.StartsWith("/xl/worksheets/sheet", StringComparison.OrdinalIgnoreCase))
                .ToArray();
            
            ZipArchiveEntry sheetEntry;
            if (sheetName != null)
            {
                await SetWorkbookRelsAsync(Archive.entries);
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
                await SetWorkbookRelsAsync(Archive.entries);
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
        
        internal async Task<List<SheetRecord>> GetWorkbookRelsAsync(ReadOnlyCollection<ZipArchiveEntry> entries, CancellationToken cancellationToken = default)
        {
            List<SheetRecord> sheetRecords = [];
            await foreach (var sheet in ReadWorkbookAsync(entries, _xmlSettings, cancellationToken).ConfigureAwait(false))
            {
                sheetRecords.Add(sheet);
            }

            await using var stream = entries.Single(w => w.FullName == "xl/_rels/workbook.xml.rels").Open();
            using var reader = XmlReader.Create(stream, _xmlSettings);
            
            if (!XmlReaderHelper.IsStartElement(reader, "Relationships", "http://schemas.openxmlformats.org/package/2006/relationships"))
                return null;

            if (!await XmlReaderHelper.ReadFirstContentAsync(reader))
                return null;

            while (!reader.EOF)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (XmlReaderHelper.IsStartElement(reader, "Relationship", "http://schemas.openxmlformats.org/package/2006/relationships"))
                {
                    var rid = reader.GetAttribute("Id");
                    foreach (var sheet in sheetRecords.Where(sheet => sheet.Rid == rid))
                    {
                        sheet.Path = reader.GetAttribute("Target");
                        break;
                    }

                    await reader.SkipAsync();
                }
                else if (!await XmlReaderHelper.SkipContentAsync(reader))
                {
                    break;
                }
            }

            return sheetRecords;
        }
        
        private static async IAsyncEnumerable<SheetRecord> ReadWorkbookAsync(ReadOnlyCollection<ZipArchiveEntry> entries, XmlReaderSettings settings, [EnumeratorCancellation] CancellationToken cancellationToken = default)
        {
            await using var stream = entries.Single(w => w.FullName == "xl/workbook.xml").Open();
            using var reader = XmlReader.Create(stream, settings);
            
            if (!XmlReaderHelper.IsStartElement(reader, "workbook", Ns))
                yield break;
            
            if (!await XmlReaderHelper.ReadFirstContentAsync(reader))
                yield break;
                
            var activeSheetIndex = 0;
            while (!reader.EOF)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (XmlReaderHelper.IsStartElement(reader, "bookViews", Ns))
                {
                    if (await XmlReaderHelper.ReadFirstContentAsync(reader))
                    {
                        while (!reader.EOF)
                        {
                            cancellationToken.ThrowIfCancellationRequested();
                            if (XmlReaderHelper.IsStartElement(reader, "workbookView", Ns))
                            {
                                var activeSheet = reader.GetAttribute("activeTab");
                                if (int.TryParse(activeSheet, out var index))
                                {
                                    activeSheetIndex = index;
                                }

                                await reader.SkipAsync();
                            }
                            else if (!await XmlReaderHelper.SkipContentAsync(reader))
                            {
                                break;
                            }
                        }
                    }
                }
                else if (XmlReaderHelper.IsStartElement(reader, "sheets", Ns))
                {
                    if (!await XmlReaderHelper.ReadFirstContentAsync(reader))
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
                            await reader.SkipAsync();
                        }
                        else if (!await XmlReaderHelper.SkipContentAsync(reader))
                        {
                            break;
                        }
                    }
                }
                else if (!await XmlReaderHelper.SkipContentAsync(reader))
                {
                    yield break;
                }
            }
        }
        
        private async Task SetSharedStringsAsync()
        {
            if (SharedStrings != null)
                return;
            
            var sharedStringsEntry = Archive.GetEntry("xl/sharedStrings.xml");
            if (sharedStringsEntry == null)
                return;
            
            await using (var stream = sharedStringsEntry.Open())
            {
                var idx = 0;
                if (_config.EnableSharedStringCache && sharedStringsEntry.Length >= _config.SharedStringCacheSize)
                {
                    SharedStrings = new SharedStringsDiskCache();
                    await foreach (var sharedString in XmlReaderHelper.GetSharedStringsAsync(stream, Ns))
                    {
                        SharedStrings[idx++] = sharedString;
                    }
                }
                else if (SharedStrings == null)
                {
                    SharedStrings = XmlReaderHelper.GetSharedStrings(stream, Ns).ToDictionary(x => idx++, x => x);
                }
            }
        }
        
        private async Task SetWorkbookRelsAsync(ReadOnlyCollection<ZipArchiveEntry> entries)
        {
            _sheetRecords = _sheetRecords ?? await GetWorkbookRelsAsync(entries);
        }
        
        public static async Task<ExcelOpenXmlSheetReader> CreateAsync(Stream stream, IConfiguration configuration, bool isUpdateMode = true)
        {
            var reader = new ExcelOpenXmlSheetReader(stream, configuration, isUpdateMode, true);
            await reader.SetSharedStringsAsync();

            return reader;
        }
#endif
        
        private static async Task<Tuple<bool, MaxRowColumnIndexes?>> TryGetMaxRowColumnIndexAsync(ZipArchiveEntry sheetEntry, XmlReaderSettings settings)
        {
            var result = new MaxRowColumnIndexes(-1, -1, false);

            using (var sheetStream = sheetEntry.Open())
            using (var reader = XmlReader.Create(sheetStream, settings))
            {
                while (await reader.ReadAsync())
                {
                    if (XmlReaderHelper.IsStartElement(reader, "c", Ns))
                    {
                        var r = reader.GetAttribute("r");
                        if (r != null)
                        {
                            if (ReferenceHelper.ParseReference(r, out var column, out var row))
                            {
                                result.MaxRowIndex = Math.Max(result.MaxRowIndex, --row);
                                result.MaxColumnIndex = Math.Max(result.MaxColumnIndex, --column);
                            }
                        }
                        else
                        {
                            result.WithoutCr = true;
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

                        result.MaxRowIndex = rIndex - 1;
                        result.MaxColumnIndex = cIndex - 1;
                        break;
                    }
                }
            }

            if (!result.WithoutCr) 
                return new Tuple<bool, MaxRowColumnIndexes?>(true, result);
            
            using (var sheetStream = sheetEntry.Open())
            using (var reader = XmlReader.Create(sheetStream, settings))
            {
                if (!XmlReaderHelper.IsStartElement(reader, "worksheet", Ns) || !await XmlReaderHelper.ReadFirstContentAsync(reader))
                    return new Tuple<bool, MaxRowColumnIndexes?>(false, null);

                while (!reader.EOF)
                {
                    if (XmlReaderHelper.IsStartElement(reader, "sheetData", Ns))
                    {
                        if (!await XmlReaderHelper.ReadFirstContentAsync(reader))
                        {
                            while (!reader.EOF)
                            {
                                if (XmlReaderHelper.IsStartElement(reader, "row", Ns))
                                {
                                    result.MaxRowIndex++;

                                    if (await XmlReaderHelper.ReadFirstContentAsync(reader))
                                    {
                                        var cellIndex = -1;
                                        while (!reader.EOF)
                                        {
                                            if (XmlReaderHelper.IsStartElement(reader, "c", Ns))
                                            {
                                                cellIndex++;
                                                result.MaxColumnIndex = Math.Max(result.MaxColumnIndex, cellIndex);
                                            }

                                            if (!await XmlReaderHelper.SkipContentAsync(reader))
                                                break;
                                        }
                                    }
                                }
                                else if (!await XmlReaderHelper.SkipContentAsync(reader))
                                {
                                    break;
                                }
                            }
                        }
                    }
                    else if (!await XmlReaderHelper.SkipContentAsync(reader))
                    {
                        break;
                    }
                }
            }

            return new Tuple<bool, MaxRowColumnIndexes?>(true, result);
        }

        private static async Task<Tuple<bool, MergeCells>> TryGetMergeCellsAsync(ZipArchiveEntry sheetEntry, XmlReaderSettings settings)
        {
            var mergeCells = new MergeCells();
            
            using (var sheetStream = sheetEntry.Open())
            using (var reader = XmlReader.Create(sheetStream, settings))
            {
                if (!XmlReaderHelper.IsStartElement(reader, "worksheet", Ns))
                    return new Tuple<bool, MergeCells>(false, null);
                
                while (await reader.ReadAsync())
                {
                    if (!XmlReaderHelper.IsStartElement(reader, "mergeCells", Ns))
                        continue;

                    if (!await XmlReaderHelper.ReadFirstContentAsync(reader))
                        return new  Tuple<bool, MergeCells>(false, null);

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

                            await XmlReaderHelper.SkipContentAsync(reader);
                        }
                        else if (!await XmlReaderHelper.SkipContentAsync(reader))
                        {
                            break;
                        }
                    }
                }
                return new Tuple<bool, MergeCells>(true, mergeCells);
            }
        }

        private async Task<Tuple<object, int>> ReadCellAndSetColumnIndexAsync(XmlReader reader, int columnIndex, bool withoutCr, int startColumnIndex, string aR, string aT)
        {
            const int xfIndex = -1;

            if (withoutCr)
                columnIndex++;
            else if (ReferenceHelper.ParseReference(aR, out var referenceColumn, out _))
                columnIndex = referenceColumn - 1; // ParseReference is 1-based

            if (columnIndex < startColumnIndex)
            {
                if (await XmlReaderHelper.ReadFirstContentAsync(reader))
                {
                    while (!reader.EOF)
                    {
                        if (!await XmlReaderHelper.SkipContentAsync(reader))
                            break;
                    }
                }

                return new Tuple<object, int>(null, columnIndex);
            }

            if (!await XmlReaderHelper.ReadFirstContentAsync(reader))
                return new Tuple<object, int>(null, columnIndex);

            object value = null;
            while (!reader.EOF)
            {
                if (XmlReaderHelper.IsStartElement(reader, "v", Ns))
                {
                    var rawValue = await reader.ReadElementContentAsStringAsync();
                    if (!string.IsNullOrEmpty(rawValue))
                        ConvertCellValue(rawValue, aT, xfIndex, out value);
                }
                else if (XmlReaderHelper.IsStartElement(reader, "is", Ns))
                {
                    var rawValue = await StringHelper.ReadStringItemAsync(reader);
                    if (!string.IsNullOrEmpty(rawValue))
                        ConvertCellValue(rawValue, aT, xfIndex, out value);
                }
                else if (!await XmlReaderHelper.SkipContentAsync(reader))
                {
                    break;
                }
            }

            return new Tuple<object, int>(value, columnIndex);
        }

        
        private struct MaxRowColumnIndexes
        {
            public MaxRowColumnIndexes(int maxRowIndex, int maxColumnIndex, bool withoutCr)
            {
                MaxRowIndex =  maxRowIndex;
                MaxColumnIndex = maxColumnIndex;
                WithoutCr = withoutCr;
            }

            public void Deconstruct(out int maxRowIndex, out int maxColumnIndex, out bool withoutCr)
            {
                maxRowIndex = MaxRowIndex;
                maxColumnIndex = MaxColumnIndex;
                withoutCr = WithoutCr;
            }

            public int MaxRowIndex { get; set; }
            public int MaxColumnIndex { get; set; }
            public bool WithoutCr { get; set; }
        } 
    }
}
