using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CsvExporter
{
    static class Util
    {
        public static void ConvertToCsv(this ExcelPackage package, string targetFile)
        {
            var worksheet = package.Workbook.Worksheets[1];

            var maxColumnNumber = worksheet.Dimension.End.Column;
            var currentRow = new List<string>(maxColumnNumber);
            var totalRowCount = worksheet.Dimension.End.Row;
            var currentRowNum = 1;

            //No need for a memory buffer, writing directly to a file
            //var memory = new MemoryStream();

            using (var writer = new StreamWriter(targetFile, false, Encoding.UTF8))
            {
                //the rest of the code remains the same
            }

            // No buffer returned
            //return memory.ToArray();
        }
    }

    public class ExportToCsv
    {
        private const string srcFolder = "SrcFolder";
        private const string serverFolder = "ServerCsv";
        private const string clientFolder = "ClientCsv";

        public enum TargetType
        {
            Server,
            Client,
        }
        public ExportToCsv(TargetType targetType)
        {
            string serverFolderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, serverFolder);
            if (!Directory.Exists(serverFolderPath))
                Directory.CreateDirectory(serverFolderPath);

            string clientFolderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, clientFolder);
            if (!Directory.Exists(clientFolderPath))
                Directory.CreateDirectory(clientFolderPath);

            string srcFolderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, srcFolder);
            if (!Directory.Exists(clientFolderPath))
                Directory.CreateDirectory(clientFolderPath);

            //FileInfo srcFile = new FileInfo(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, srcfileName));

            DirectoryInfo di = new DirectoryInfo(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, srcFolder));
            foreach (FileInfo srcFile in di.GetFiles())
            {
                if (srcFile.Extension.ToLower().CompareTo(".xlsx") == 0)
                {
                    using (ExcelPackage excelPackage = new ExcelPackage(srcFile))
                    {
                        //excelPackage.ConvertToCsv(Path.Combine(csvFolderPath, "test"));
                        var format = new ExcelOutputTextFormat();
                        format.Encoding = Encoding.UTF8;

                        for (int cnt = 0; cnt < excelPackage.Workbook.Worksheets.Count; cnt++)
                        {
                            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets[cnt];

                            int totalColumn = worksheet.Dimension.End.Column;
                            for (int i = totalColumn; i > 0; i--)
                            {
                                string targetStr = worksheet.Cells[3, i].Text.ToLower();

                                if (targetStr.CompareTo("nodata") == 0)
                                    worksheet.DeleteColumn(i);

                                if (targetType == TargetType.Client && targetStr.CompareTo("server") == 0)
                                    worksheet.DeleteColumn(i);

                                if (targetType == TargetType.Server && targetStr.CompareTo("client") == 0)
                                    worksheet.DeleteColumn(i);

                            }
                            worksheet.DeleteRow(3);

                            //for (int i = 0; i < worksheet.Dimension.End.Row; i++)
                            //{
                            //    for (int j = 0; j < worksheet.Dimension.End.Column; j++)
                            //    {
                            //        worksheet.Cells[i + 1, j + 1].Value = string.Format($"\"{worksheet.Cells[i + 1, j + 1].Text}\"");
                            //    }
                            //}
                            
                            //Path.Combine(csvFolderPath, worksheet.Name)

                            FileInfo dstFile = new FileInfo(Path.Combine(targetType == TargetType.Server ? serverFolderPath : clientFolderPath, $"{worksheet.Name}.csv"));
                            worksheet.Cells[1, 1, worksheet.Dimension.End.Row, worksheet.Dimension.End.Column].SaveToText(dstFile, format);
                        }
                    }
                }
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            ExportToCsv exportClient = new ExportToCsv( ExportToCsv.TargetType.Client);
            ExportToCsv exportServer = new ExportToCsv(ExportToCsv.TargetType.Server);
        }
    }
}
