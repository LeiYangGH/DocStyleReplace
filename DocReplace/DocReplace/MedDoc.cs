using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocReplace
{
    public class MedDoc
    {
        public string FullName { get; set; }
        public string MedName { get; set; }
        public string MedSN { get; set; }
    }

    public class MedStandard
    {
        public MedStandard(string chname, string oldno, string newno)
        {
            this.ChineseName = chname;
            this.OldNO = oldno;
            this.NewNO = newno;
        }
        public string ChineseName { get; set; }
        public string OldNO { get; set; }
        public string NewNO { get; set; }
        public override string ToString()
        {
            //return $"Ch={this.ChineseName}\tOldNO={this.OldNO}\tNewNO={this.NewNO}";
            //return $"Ch={this.ChineseName}*\tOldNO={this.OldNO}*\tNewNO={this.NewNO}*";
            return $"{this.ChineseName}\t{this.OldNO}\t{this.NewNO}";
        }

        private static string ReplaceSpecialChars(string No)
        {
            return No.Replace("/", "").Replace("•", "").Replace("·", "").Replace("-", "").ToUpper();
        }

        public static bool EqualAfterReplaceSpecialChars(string no1, string no2)
        {
            return ReplaceSpecialChars(no1) == ReplaceSpecialChars(no2);
        }
    }
}
