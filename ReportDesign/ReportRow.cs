using System;
using System.Collections.Generic;
using System.Text;

namespace ReportDesign
{
    class Company
    {
        public string name = "";
        public string year = "";
        public override int GetHashCode()
        {
            if (name == null || year == null) return 0;
            return name.GetHashCode() + year.GetHashCode();
        }

        public override bool Equals(object obj)
        {
            Company other = (Company)obj;
            return other != null && other.name == this.name && other.year == this.year;
        }
    }
    class CompanyWithoutYear
    {
        public string name = "";
        public override int GetHashCode()
        {
            if (name == null) return 0;
            return base.GetHashCode();
        }
        public override bool Equals(object obj)
        {
            CompanyWithoutYear other = (CompanyWithoutYear)obj;
            return other != null && other.name == this.name;
        }
    }
}
