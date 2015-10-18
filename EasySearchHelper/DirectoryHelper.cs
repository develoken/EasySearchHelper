using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace EasySearchHelper
{
    class DirectoryHelper
    {
        private const string TYPE_XLS = "*.xls";
        private const string TYPE_TXT = "*.txt";

        private static volatile DirectoryHelper instance;
        private static object syncRoot = new Object();

        private DirectoryHelper()
        {
        }

        public static DirectoryHelper Instance
        {
            get
            {
                if (instance == null)
                {
                    lock (syncRoot)
                    {
                        if (instance == null)
                            instance = new DirectoryHelper();
                    }
                }
                return instance;
            }
        }

        public List<string> FindXlsFiles(string rootPath)
        {
            return FindFiles(rootPath, TYPE_XLS);
        }

        private List<string> FindFiles(string rootPath, string fileType)
        {
            List<string> pathList = new List<string>();

            foreach (string path in Directory.GetFiles(rootPath, fileType, SearchOption.AllDirectories))
            {
                pathList.Add(path);
            }
            return pathList;
        }
    }
}