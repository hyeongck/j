using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace Dir
{
    public class Dir_Directory
    {
        public Dir_Directory(string Address)
        {
            DirectoryInfo Di = new DirectoryInfo(Address);
            if (Di.Exists == false)
            {
                Di.Create();
            }
        }

        public Dir_Directory()
        {

        }

        public bool File_Exits(string FileName)
        {
            bool Flag = true;
            if (File.Exists(FileName))
            {
                Flag = true;
            }
            else
            {
                Flag = false;
            }
            return Flag;
        }
    }
}
