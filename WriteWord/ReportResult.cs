using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace WriteWord
{
    public class ReportResult
    {
        int _status;
        public int status
        {
            get { return _status; }
            set { _status = value; }
        }

        string _message;
        public string message
        {
            get { return _message; }
            set { _message = value; }
        }

        ReportInfo _data;
        public ReportInfo data
        {
            get { return _data; }
            set { _data = value; }
        }
        public static ReportResult LoadFromFile(string FilePath)
        {
            try
            {
                StreamReader sr = new StreamReader(FilePath);
                string content = sr.ReadToEnd();
                sr.Close();
                return LoadFromString(content);
            }
            catch (Exception ex)
            {
                return null;
            }
            
        }

        public static ReportResult LoadFromString(string JsonStr)
        {
            ReportResult rr = new ReportResult();
            try
            {
                rr = (ReportResult)JsonTools.JsonToObject(JsonStr, typeof(ReportResult));
            }
            catch (Exception ex)
            {
                return null;
            }
            return rr;
        }
    }
}
