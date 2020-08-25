using DocumentFormat.OpenXml.Office.CustomUI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace docs
{
    class DocInfo
    {
        public string CampoDW { get; set; }
        public string TabelaDW { get; set; }
        public string CampoOrigem { get; set; }
        public string TabelaOrigem { get; set; }
        public string BaseOrigem { get; set; }

        public DocInfo(string campoDW, string tabelaDW)
        {
            this.CampoDW = campoDW;
            this.TabelaDW = tabelaDW;
        }
    }
}
