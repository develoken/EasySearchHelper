using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EasySearchHelper.Converter
{
    public interface IConverter
    {
        void Initialize(string sourceFilePath, string destinationPath);
        void Convert();
    }
}
