using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace EasySearchHelper.Converter
{
    class ConvertExcel : IConverter
    {
        private string srcFilePath;
        private string dstPath;
        
        public ConvertExcel()
        {
        }

        public void Initialize(string srcFilePath, string dstPath)
        {
            this.srcFilePath = srcFilePath;
            this.dstPath = dstPath;
        }

        public void Convert()
        {
            string dstFilePath = Path.Combine(dstPath, Path.GetFileNameWithoutExtension(srcFilePath) + ".txt");
            
            if(File.Exists(dstFilePath))
            {
                // Skip if already exists
                return;
            }
            
            StreamWriter outputFile = new StreamWriter(dstFilePath);

            foreach (string line in ReadLines())
            {
                if (!String.IsNullOrEmpty(line))
                {
                    outputFile.WriteLine(line);
                }
            }

            outputFile.Close();
        }
        
        private IEnumerable<string> ReadLines()
        {
            Application application = new Application();
            Workbook workbook = application.Workbooks.Open(srcFilePath);

            try
            {
                foreach (Worksheet worksheet in workbook.Worksheets)
                {
                    for (int i = 1; i <= worksheet.UsedRange.Rows.Count; i++)
                    {
                        StringBuilder rowValue = new StringBuilder();

                        for (int j = 1; j <= worksheet.UsedRange.Columns.Count; j++)
                        {
                            var value = worksheet.Cells[i, j].Value;

                            if (value != null)
                            {
                                rowValue.Append(value.ToString());
                                rowValue.Append(", ");
                            }
                        }

                        yield return rowValue.ToString();
                    }
                }
            }
            finally
            {
                workbook.Close();
                application.Quit();
            }
        }
    }
}
