using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WriteWord
{
    public class ReportPartInfo
    {
        public ReportPartInfo(string PartStr, string ChTitle, string Entitle)
        {
            _PartStr = PartStr;
            _ChTitle = ChTitle;
            _EnTitle = Entitle;
        }
        string _PartStr;

        public string PartStr
        {
            get { return _PartStr; }
            set { _PartStr = value; }
        }
        string _ChTitle;

        public string ChTitle
        {
            get { return _ChTitle; }
            set { _ChTitle = value; }
        }
        string _EnTitle;

        public string EnTitle
        {
            get { return _EnTitle; }
            set { _EnTitle = value; }
        }
    }
}
