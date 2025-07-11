﻿using System.IO.Compression;

namespace MiniExcelLibs.Zip;

internal class ZipPackageInfo(ZipArchiveEntry zipArchiveEntry, string contentType)
{
    public ZipArchiveEntry ZipArchiveEntry { get; set; } = zipArchiveEntry;
    public string ContentType { get; set; } = contentType;
}