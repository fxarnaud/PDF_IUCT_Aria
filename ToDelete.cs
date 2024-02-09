using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PDF_IUCT
{
    public class ToDelete
    {
        public Dictionary<string, string> files_pathes { get; set; }
        public ToDelete()
        {
            files_pathes = new Dictionary<string, string>();
        }
    }
}
