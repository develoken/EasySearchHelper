using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using EasySearchHelper.Converter;

namespace EasySearchHelper
{
    class Program
    {
        static void PrintUsage()
        {
            Console.WriteLine("Wrong arguments");
        }

        static void Main(string[] args)
        {
            if (args.Length != 2)
            {
                PrintUsage();
            }

            string srcPath = args[0];
            string dstPath = args[1];
            Console.WriteLine("Source path: " + srcPath);
            Console.WriteLine("Destination path: " + dstPath);

            if (!Directory.Exists(dstPath))
            {
                Directory.CreateDirectory(dstPath);
            }

            foreach (string filePath in DirectoryHelper.Instance.FindXlsFiles(srcPath))
            {
                if (!File.Exists(filePath))
                {
                    Console.WriteLine("Something went wrong. File does not exists: {0}", filePath);
                    continue;
                }

                Console.WriteLine("Converting {0}...", filePath);
                
                IConverter converter = ConverterFactory.Instance.getConverter(filePath);
                converter.Initialize(filePath, dstPath);
                converter.Convert();

                Console.WriteLine("Convert completed.");
            }

            Console.WriteLine("Enter any keys to exit");
            Console.ReadLine();
        }
    }
}
