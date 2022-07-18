﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace IMcorrectionTool
{
    class Warming
    {
        public string ID
        {
            get
            {
                string resultText = WarningText;
                // Выборка строк, в которых есть id. Необходимо для ограничения количества
                // строк, прогоняемых через регулярное выражение, ибо оно работает очень медленно.
                if (WarningText.Contains("id"))
                    resultText = Regex.Replace(WarningText, @"\w*id\W*[0-9]*", "Id=", RegexOptions.IgnoreCase).Trim();
                return ObjectUID + resultText;
            }
        }
        public string ODU { get; set; }
        public string ModelingAuthoritySet { get; set; }
        public string RuleID { get; set; }

        public string ObjectUID { get; set; }

        public string ObjectName { get; set; }
        public string ObjectClass { get; set; }
        public string WarningText { get; set; }
        public string Comment { get; set; }
        public string PreviousComment { get; set; }

        public bool IsNewInMonth { get; set; }
        public bool IsNewInKGID { get; set; }
        //public string thisMonthComment { get; set; }
        public bool IsCorrectedInKGID { get; set; }


        public Warming(string odu, string modelingAuthoritySet, string ruleId, string objectUID, string objectName, string objectClass, string warningText, string commentText = "")
        {
            ODU = odu;
            ModelingAuthoritySet = modelingAuthoritySet;
            RuleID = ruleId;
            ObjectUID = objectUID;
            ObjectName = objectName;
            ObjectClass = objectClass;
            WarningText = warningText;
            Comment = commentText;
            IsNewInMonth = true;
        }
    }
}
