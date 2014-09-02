using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace RightFaxIt
{
    public class Fax
    {
        
        //Class Properties
        public string CustomerName { get; set; }
        public string FaxNumber { get; set; }
        public string Document { get; set; }
        public string Account {get; set; }
        public string FileName { get; set; }
        public Boolean IsValid { get; set; }
        public Boolean Sent { get; set; }
        public Boolean Rejected { get; set; }

        //~Fax() { }
        public Fax(){ }

        /// <summary>
        /// Fax with predetermined info
        /// </summary>
        /// <param name="document"></param>
        /// <param name="customerName"></param>
        /// <param name="number"></param>
        public Fax(String document, String customerName, String number)
        {
            this.CustomerName = customerName;
            this.Document = document;
            this.FaxNumber = number;
            this.FileName = System.IO.Path.GetFileNameWithoutExtension(document);
        }

        /// <summary>
        /// Fax constructor where data is parsed from filename.
        /// </summary>
        /// <param name="document"></param>
        public Fax(String document)
        {
            this.IsValid = true;
            this.Sent = false;
            this.Rejected = false;
            this.Document = document;
            this.FileName = System.IO.Path.GetFileNameWithoutExtension(document);
            ParseFileName(this.FileName);
            this.Account = RegexFileName(@"\d{5}-\d{5}");
            this.FaxNumber = RegexFileName(@"\d{1}-\d{3}-\d{3}-\d{4}");
            if ((this.FaxNumber == "1-999-999-9999") || (this.FaxNumber == "1-888-888-8888"))
            {
                this.IsValid = false;
            }
        }
        
        /// <summary>
        /// Parses the filename to retrieve fax recipient info.
        /// </summary>
        /// <remarks>
        /// Strictly follows the formart of FAXTO-...
        /// </remarks>
        private void ParseFileName(string fileName)
        {
            try
            {
                string[] strSplit = fileName.Split('-');
                this.CustomerName = strSplit[4].Replace('_', ' ');
            }
            catch (IndexOutOfRangeException)
            {
                this.CustomerName = "NOT_FOUND";
                this.IsValid = false;
            }
        }

        private string RegexFileName(string pattern) 
        {
            Regex rgx = new Regex(pattern, RegexOptions.IgnoreCase);
            MatchCollection matches = rgx.Matches(this.FileName);
            if (matches.Count > 0)
            {
                    return matches[0].Value;
            }
            else
            {

                this.IsValid = false;
                return "NOT_FOUND";
            }
        }

    
    }

}
