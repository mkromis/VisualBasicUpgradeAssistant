using System;

namespace VisualBasicUpgradeAssistant.Core.DataClasses
{
    /// <summary>
    /// Summary description for ControlListItem.
    /// </summary>
    public class ControlListItem
    {
        public String VB6Name { get; set; }
        public String CsharpName { get; set; }
        public Boolean Unsupported { get; set; }
        public Boolean InvisibleAtRuntime { get; set; }
    }
}
