using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace EasySearchHelper.Converter
{
    public class ConverterFactory
    {
        private static volatile ConverterFactory instance;
        private static object syncRoot = new Object();

        private ConverterFactory()
        {
            
        }

        public static ConverterFactory Instance
        {
            get
            {
                if (instance == null)
                {
                    lock (syncRoot)
                    {
                        if (instance == null)
                            instance = new ConverterFactory();
                    }
                }
                return instance;
            }
        }

        public IConverter getConverter(string filePath)
        {
            string extension = Path.GetExtension(filePath);

            switch (extension)
            {
                case ".xls":
                case ".xlsx":
                    return new ConvertExcel();
                    
                default:
                    return new ConvertDefault();
            }
        }
    }
}
