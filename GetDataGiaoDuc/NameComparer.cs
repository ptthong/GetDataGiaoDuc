using GetDataGiaoDuc.APISMAS;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace GetDataGiaoDuc
{
    class NameComparer : IComparer<PupilProfile>
    {
        public int Compare(PupilProfile x, PupilProfile y)
        {
            if(x.ClassID!=y.ClassID) return x.ClassID.CompareTo(y.ClassID);
            return getName(x.FullName).CompareTo(getName(y.FullName));
            
        }
        public String getName(String fullname)
        {
            String[] arrname = fullname.Split(' ');
            return arrname[arrname.Length - 1];
        }
    }
}
