using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace macros.Models
{
    public class StudentRow
    {
        public int RowNumber { get; set; }

        public string LastName { get; set; }
        public string FirstName { get; set; }
        public string MiddleName { get; set; }
        public string Login { get; set; }
        public string Password { get; set; }

        public string FullName
        {
            get
            {
                return string.Format("{0} {1} {2}", LastName, FirstName, MiddleName).Trim();
            }
        }

        public string ShortFolderName
        {
            get
            {
                string firstInitial = string.IsNullOrWhiteSpace(FirstName) ? "" : FirstName.Substring(0, 1);
                string middleInitial = string.IsNullOrWhiteSpace(MiddleName) ? "" : MiddleName.Substring(0, 1);

                return string.Format("{0} {1} {2}", LastName, firstInitial, middleInitial).Trim();
            }
        }
    }
}
