using System;
using System.Collections;

namespace VisualBasicUpgradeAssistant.Core.DataClasses
{
    /// <summary>
    /// Summary description for Property.
    /// </summary>
    public class ControlProperty
    {
        public String Name { get; set; }

        public String Value { get; set; }

        public String Comment { get; set; }

        public Boolean Valid { get; set; }

        public ControlProperty()
        {
            PropertyList = new ArrayList();
            Valid = false;
        }

        public ArrayList PropertyList { get; private set; }

    }
}
