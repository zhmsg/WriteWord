using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace WriteWord
{
    class ReportResourcesRequired
    {
        public string CoverPicPath;
        public string SpeechPicPath;
        public string BackCoverPicPath;

        public string GreenPicPath;
        public string RedPicPath;
        public string YellowPicPath;

        public string HeaderPicPath;
        public string LogoPicPath;

        public string ZhiCiPath;
        public string FuhaoYiyiPath;
        public string HowToPath;
        public string MianZePath;
        public string GongSiPath;


        public ReportResourcesRequired(string ResourcesDirPath = "")
        {
            if (ResourcesDirPath == "")
                ResourcesDirPath = Environment.CurrentDirectory;
            YellowPicPath = ResourcesDirPath + "\\yellow.png";
            RedPicPath = ResourcesDirPath + "\\red.png";
            GreenPicPath = ResourcesDirPath + "\\green.png";
            HeaderPicPath = ResourcesDirPath + "\\header.png";
            LogoPicPath = ResourcesDirPath + "\\logo.png";
            CoverPicPath = ResourcesDirPath + "\\cover.png";
            SpeechPicPath = ResourcesDirPath + "\\speech.jpg";
            BackCoverPicPath = ResourcesDirPath + "\\backcover.png";

            ZhiCiPath = ResourcesDirPath + "\\zhici.txt";
            FuhaoYiyiPath = ResourcesDirPath + "\\fuhaoyiyi.txt";
            HowToPath = ResourcesDirPath + "\\HowTo.txt";
            MianZePath = ResourcesDirPath + "\\mianze.txt";
            GongSiPath = ResourcesDirPath + "\\gongsi.txt";
        }

        public void TestAndLockResources()
        {
            Type t = this.GetType();
            MemberInfo [] MIs =  t.GetMembers();
            foreach (MemberInfo mi in MIs)
            {
                Console.WriteLine(mi.Name);
            }
        }
    }
}
