using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace WriteWord
{
    class Program
    {
        
        static void Main(string[] args)
        {
            string TaskID = "d4814f9652e311e6879900163e000a31";
            string SaveDir = @"C:\Users\meisanggou\Desktop";
            if (args.Length >= 1)
            {
                TaskID = args[0];
                if (args.Length >= 2)
                    SaveDir = args[1];
            }
            Console.WriteLine("TaskID：" + TaskID);
            Console.WriteLine("SaveDir：" + SaveDir);
            if (Directory.Exists(SaveDir) == false)
            {
                Console.WriteLine(SaveDir + "目录不存在");
            }
            string url = String.Format("http://inhouse.gene.ac/api/v2/health/task/v3/204/{0}/", TaskID);
            string res;
            HttpHelper.getResponseText(url, out res, "GET", "", "dinghui", "gene.ac");
            ReportResult r = ReportResult.LoadFromString(res);
            if (r.status % 10000 == 1)
            {
                MSGReport mr = new MSGReport();
                object WordPath = String.Format(@"{0}\{1}_{2}_报告.docx", SaveDir, r.data.survey_result.name, TaskID);
                mr.WriteReport(r.data);
                mr.SaveASDocx(WordPath);
                mr.Close();

                string PDFPath = String.Format(@"{0}\{1}_{2}_报告.pdf", SaveDir, r.data.survey_result.name, TaskID);
                WordToPDF WToP = new WordToPDF();
                WToP.ChangeToPDF(WordPath, PDFPath);
                WToP.Close();

            }
            else
            {
                Console.WriteLine(r.message);
            }
            //Console.WriteLine(res);

            Console.Out.WriteLine("Hello World!");
        }
    }
}
