using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WriteWord
{
    public class ReportTest
    {
        string _disease_name;

        public string disease_name
        {
            get { return _disease_name; }
            set { _disease_name = value; }
        }
        string _early_detection;

        public string early_detection
        {
            get { return _early_detection; }
            set { _early_detection = value; }
        }
        string _introduction;

        public string introduction
        {
            get { return _introduction; }
            set { _introduction = value; }
        }
        string _prevention;

        public string prevention
        {
            get { return _prevention; }
            set { _prevention = value; }
        }
        string _reference;

        public string reference
        {
            get {
                if (_reference == "")
                    return "暂无建议";
                return _reference; 
            }
            set { _reference = value; }
        }
        string _risk_level;

        public string risk_level
        {
            get { return _risk_level; }
            set { _risk_level = value; }
        }
    }
}
