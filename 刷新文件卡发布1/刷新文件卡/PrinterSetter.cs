using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;

namespace 刷新文件卡
{
    class PrinterSetter
    {
        [DllImport("winspool.drv")]
        public static extern bool SetDefaultPrinter(String Name);
    }
}
