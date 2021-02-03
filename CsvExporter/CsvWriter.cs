using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CsvExporter
{
    class CsvWriter
    {
        private CsvConfig m_config;
        private StringBuilder m_csvContents;

        public CsvWriter(CsvConfig config = null)
        {
            if (config == null)
                m_config = CsvConfig.Default;
            else
                m_config = config;

            m_csvContents = new StringBuilder();
        }

        public void AddRow(IEnumerable<string> cells)
        {
            int i = 0;
            foreach (string cell in cells)
            {
                m_csvContents.Append(ParseCell(cell));
                m_csvContents.Append(m_config.Delimiter);

                i++;
            }

            m_csvContents.Length--; // remove last delimiter
            m_csvContents.Append("\r\n");
        }

        private string ParseCell(string cell)
        {
            // cells cannot be multi-line
            cell = cell.Replace("\r", "");
            cell = cell.Replace("\n", "");

            if (!NeedsToBeEscaped(cell))
                return cell;

            // double every quotation mark
            cell = cell.Replace(m_config.QuotationMark.ToString(), string.Format("{0}{0}", m_config.QuotationMark));

            // add quotation marks at the beginning and at the end
            cell = m_config.QuotationMark + cell + m_config.QuotationMark;

            return cell;
        }

        private bool NeedsToBeEscaped(string cell)
        {
            if (cell.Contains(m_config.QuotationMark.ToString()))
                return true;

            if (cell.Contains(m_config.Delimiter.ToString()))
                return true;

            return false;
        }

        public string Write()
        {
            return m_csvContents.ToString();
        }
    }
}
