using System;
using System.Collections;

namespace VisualBasicUpgradeAssistant.Core.DataClasses
{
    /// <summary>
    /// Summary description for Enum.
    /// </summary>
    public class EnumType
    {
        public EnumType()
        {
            ItemList = new ArrayList();
        }

        public String Name { get; set; }

        public String Scope { get; set; }

        public ArrayList ItemList { get; private set; }
    }
}
