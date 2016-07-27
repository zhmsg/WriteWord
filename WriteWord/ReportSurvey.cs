using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
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
                string b = _birth;
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
            get { return _height; }
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
                string s = _id.Substring(0, 5) + "****" + _id.Substring(9, 5) + "****";
                return s;
            }
        }
        int _sex;

        public int sex
        {
            set { _sex = value; }
        }
        public string Sex
        {
            get {
                if (_sex == 1)
                    return "男";
                else if (_sex == 2)
                    return "女";
                return "未知";
            }
        }
        string _weight;

        public string weight
        {
            get { return _weight; }
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
                info += "       身高：" + height + "cm";
                info += "       体重：" + weight + "kg";
                info += "\n身份证号：" + ID;
                return info;
            }
        }
    }
}
