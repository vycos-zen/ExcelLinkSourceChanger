using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LinkSourceChanger
{
    class XlWorkBook
    {
        public XlWorkBook()
        {
            XlLinkPaths = new List<string[]>();
        }
        public string XlFileName { get; set; }
        public string XlFilePath { get; set; }
        public string XlFullPath { get; set; }
        public bool XlSetToUpdateLinks { get; set; }
        // 0 is the old link, 1 is the new link, and 2 is the replace target, 3 is the destination path in XlLinkPaths[]
        public List<string[]> XlLinkPaths { get; set; }
    }
}
