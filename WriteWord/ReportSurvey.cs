using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace WriteWord
{
    public class ReportSurvey
    {
        string _birth;

        public string birth
        {
            get { return _birth; }
            set { _birth = value; }
        }
        public string Birth
        {
            get
            {
                if (_birth == null || _birth == "")
                    return "";
                string b = "";
                Regex reg = new Regex(@"[^\d]");
                string [] bd = reg.Split(_birth);
                if (bd.Length > 0)
                    b += bd[0] + "年";
                if (bd.Length > 1)
                    b += bd[1] + "月";
                if (bd.Length > 2)
                    b += bd[2] + "日";
                return b;
            }
        }
        string _name;

        public string name
        {
            get { return _name; }
            set { _name = value; }
        }
        string _concern;

        public string concern
        {
            get { return _concern; }
            set { _concern = value; }
        }
        string _height;

        public string height
        {
            get 
            {
                if (_height.Length > 0)
                    return _height + "cm";
                return _height; 
            }
            set { _height = value; }
        }
        string _id;

        public string id
        {
            get { return _id; }
            set { _id = value; }
        }
        public string ID
        {
            get
            {
                if (_id == null)
                    return "******************";
                if (_id == "")
                    return "";
                string s = _id.Substring(0, 5) + "****" + _id.Substring(9, 5) + "****";
                return s;
            }
        }
        int _sex;

        public int sex
        {
            set { _sex = value; }
        }

        string _Sex;
        public string Sex
        {
            get {
                if (_Sex != null)
                    return _Sex;
                if (_sex == 1)
                    _Sex = "男";
                else if (_sex == 2)
                    _Sex = "女";
                else
                    _Sex = "未知";
                return _Sex;
            }
            set { _Sex = value; }
        }
        string _weight;

        public string weight
        {
            get 
            {
                if (_weight.Length > 0)
                    return _weight + "kg";
                return _weight; 
            }
            set { _weight = value; }
        }
        public string PersonalInfo
        {
            get
            {
                string info = "";
                info += "姓名：" + name;
                info += "       性别：" + Sex;
                info += "       出生日期：" + Birth;
                info += "       身高：" + height;
                info += "       体重：" + weight;
                info += "\n身份证号：" + ID;
                return info;
            }
        }

        // 家族史
        List<string> _family_history = new List<string>();

        public List<string> family_history
        {
            get { return _family_history; }
            set { _family_history = value; }
        }

        // 生活习惯
        List<string> _habits = new List<string>();

        public List<string> habits
        {
            get { return _habits; }
            set { _habits = value; }
        }
        // 既往史
        List<string> _past_medical_history = new List<string>();

        public List<string> past_medical_history
        {
            get 
            {
                if (_past_medical_history.Count <= 0)
                    _past_medical_history.Add("");
 
                return _past_medical_history; 
            }
            set { _past_medical_history = value; }
        }

        public static ReportSurvey LoadFromFile(string FilePath)
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

        public static ReportSurvey LoadFromString(string JsonStr)
        {
            ReportSurvey rs = new ReportSurvey();
            try
            {
                rs = (ReportSurvey)JsonTools.JsonToObject(JsonStr, typeof(ReportSurvey));
            }
            catch (Exception ex)
            {
                return null;
            }
            return rs;
        }
        string _report_no;

        public string report_no
        {
            get {
                if (_report_no == null)
                    _report_no = "";
                return _report_no; 
            }
            set { _report_no = value; }
        }
    }
}
