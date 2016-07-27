using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace WriteWord
{
    public class ReportInfo
    {
        public ReportInfo()
        {
            _PartInfo.Add(new ReportPartInfo("PART 01", "阅读指南", "Reading Guide"));
            _PartInfo.Add(new ReportPartInfo("PART 02", "如何理解基因检测和疾病风险", "How to understand genetic testing and disease risk"));
            _PartInfo.Add(new ReportPartInfo("PART 03", "您的基因检测概况", "Overview of your genetic testing"));
            _PartInfo.Add(new ReportPartInfo("PART 04", "肿瘤风险详解", "Cancer risk interpretation"));
            _PartInfo.Add(new ReportPartInfo("PART 05", "慢病风险管理", "Chronic disease risk interpretation"));
            _PartInfo.Add(new ReportPartInfo("PART 06", "药物治疗与副作用详解", "The curative effect and side effects of drugs"));
            _PartInfo.Add(new ReportPartInfo("PART 07", "营养物质吸收代谢详解", "The absorption and metabolism of nutrients"));
            _PartInfo.Add(new ReportPartInfo("PART 08", "附录", "The appendix"));
            _PartInfo.Add(new ReportPartInfo("PART 09", "免责说明", "Disclaimer of liability"));
            _PartInfo.Add(new ReportPartInfo("PART 10", "公司简介", "Company profile"));
        }

        List<ReportTestItem> _chronic_disease_list = new List<ReportTestItem>();

        public List<ReportTestItem> chronic_disease_list
        {
            get { return _chronic_disease_list; }
            set { _chronic_disease_list = value; }
        }
        List<ReportTestMedicine> _medicine_list = new List<ReportTestMedicine>();

        public List<ReportTestMedicine> medicine_list
        {
            get { return _medicine_list; }
            set { _medicine_list = value; }
        }
        List<ReportTestItem> _nutrition_list = new List<ReportTestItem>();

        public List<ReportTestItem> nutrition_list
        {
            get { return _nutrition_list; }
            set { _nutrition_list = value; }
        }
        ReportSurvey _survey_result;

        public ReportSurvey survey_result
        {
            get { return _survey_result; }
            set { _survey_result = value; }
        }
        List<ReportTestItem> _tomour_list = new List<ReportTestItem>();

        public List<ReportTestItem> tomour_list
        {
            get { return _tomour_list; }
            set { _tomour_list = value; }
        }

        List<ReportPartInfo> _PartInfo = new List<ReportPartInfo>();

        public List<ReportPartInfo> PartInfo
        {
            get { return _PartInfo; }
            set { _PartInfo = value; }
        }
    }
}
