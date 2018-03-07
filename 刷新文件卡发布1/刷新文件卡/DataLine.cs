using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace 刷新文件卡
{
    class DataLine
    {        
        public int number
        {
            get { return number; }
            set { number = value; }
        }

        public int level
        {
            get { return level; }
            set { level = value; }
        }
        public string partNumber
        {
            get { return partNumber; }
            set { partNumber = value; }
        }
        public string partName
        {
            get { return partName; }
            set { partName = value; }
        }
        public string originalPath
        {
            get { return originalPath; }
            set { originalPath = value; }
        }
        public string targetPath
        {
            get { return targetPath; }
            set { targetPath = value; }
        }

    }
}
