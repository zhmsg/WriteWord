using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WriteWord
{
    public class ReportTestMedicine: ReportTest
    {
        
        List<List<string>> _medication_tips = new List<List<string>>();

        public List<List<string>> medication_tips
        {
            get { return _medication_tips; }
            set { _medication_tips = value; }
        }
        List<List<List<string>>> _patient_sample_variant = new List<List<List<string>>>();

        public List<List<List<string>>> patient_sample_variant
        {
            get { return _patient_sample_variant; }
            set { _patient_sample_variant = value; }
        }
    }
}
