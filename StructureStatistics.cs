using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using VMS.TPS.Common.Model.API;
//using System.Drawing;
using System.Windows.Media;
//using System.Windows.Media;
//using System.Windows.Media;

namespace PDF_IUCT
{
    public class StructureStatistics
    {
        public Structure structure { get; set; }
        public string structure_id { get; set; }
        public double volume { get; set; }
        public double meandose { get; set; }
        public double maxdose { get; set; }
        public double d1cc { get; set; }

        public double d0035cc { get; set; }
        public bool isChecked { get; set; }
        public SolidColorBrush BackgroundColor
        {
            get { return new SolidColorBrush(structure.Color); }
        }

    }
}
