using MiniExcelLibs.OpenXml;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
    
namespace MiniExcelLibs.Utils
{
    internal static class StringHelper
    {
        private static readonly string[] Ns = { Config.SpreadsheetmlXmlns, Config.SpreadsheetmlXmlStrictns };

        public static string GetLetters(string content) => new string(content.Where(char.IsLetter).ToArray());
        public static int GetNumber(string content) => int.Parse(new string(content.Where(char.IsNumber).ToArray()));

        /// <summary>
        /// Copied and modified from ExcelDataReader - @MIT License
        /// </summary>
        public static string ReadStringItem(XmlReader reader)
        {
            var result = new StringBuilder();
            if (!XmlReaderHelper.ReadFirstContent(reader))
                return string.Empty;

            while (!reader.EOF)
            {
                if (XmlReaderHelper.IsStartElement(reader, "t", Ns))
                {
                    // There are multiple <t> in a <si>. Concatenate <t> within an <si>.
                    result.Append(reader.ReadElementContentAsString());
                }
                else if (XmlReaderHelper.IsStartElement(reader, "r", Ns))
                {
                    result.Append(ReadRichTextRun(reader));
                }
                else if (!XmlReaderHelper.SkipContent(reader))
                {
                    break;
                }
            }

            return result.ToString();
        }
        
        public static async Task<string> ReadStringItemAsync(XmlReader reader)
        {
            var result = new StringBuilder();
            if (!await XmlReaderHelper.ReadFirstContentAsync(reader))
                return string.Empty;

            while (!reader.EOF)
            {
                if (XmlReaderHelper.IsStartElement(reader, "t", Ns))
                {
                    // There are multiple <t> in a <si>. Concatenate <t> within an <si>.
                    result.Append(await reader.ReadElementContentAsStringAsync());
                }
                else if (XmlReaderHelper.IsStartElement(reader, "r", Ns))
                {
                    result.Append(await ReadRichTextRunAsync(reader));
                }
                else if (!await XmlReaderHelper.SkipContentAsync(reader))
                {
                    break;
                }
            }

            return result.ToString();
        }

        /// <summary>
        /// Copied and modified from ExcelDataReader - @MIT License
        /// </summary>
        private static string ReadRichTextRun(XmlReader reader)
        {
            var result = new StringBuilder();
            if (!XmlReaderHelper.ReadFirstContent(reader))
                return string.Empty;

            while (!reader.EOF)
            {
                if (XmlReaderHelper.IsStartElement(reader, "t", Ns))
                {
                    result.Append(reader.ReadElementContentAsString());
                }
                else if (!XmlReaderHelper.SkipContent(reader))
                {
                    break;
                }
            }

            return result.ToString();
        }
        
        private static async Task<string> ReadRichTextRunAsync(XmlReader reader)
        {
            var result = new StringBuilder();
            if (! await XmlReaderHelper.ReadFirstContentAsync(reader))
                return string.Empty;

            while (!reader.EOF)
            {
                if (XmlReaderHelper.IsStartElement(reader, "t", Ns))
                {
                    result.Append(await reader.ReadElementContentAsStringAsync());
                }
                else if (!await XmlReaderHelper.SkipContentAsync(reader))
                {
                    break;
                }
            }

            return result.ToString();
        }
    }
}
