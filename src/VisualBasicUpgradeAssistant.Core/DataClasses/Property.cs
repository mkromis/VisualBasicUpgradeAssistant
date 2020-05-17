using System;
using System.Collections;

namespace VisualBasicUpgradeAssistant.Core.DataClasses
{
    /// <summary>
    /// Summary description for Property.
    /// </summary>
    public class Property
    {
        public Property()
        {
            LineList = new ArrayList();
            ParameterList = new ArrayList();
        }

        public String Name { get; set; }

        public String Type { get; set; }

        public String Direction { get; set; }

        public String Comment { get; set; }

        public String Scope { get; set; }

        public ArrayList ParameterList { get; set; }

        public ArrayList LineList { get; set; }
    }
}
