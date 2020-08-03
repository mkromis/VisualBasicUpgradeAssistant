using System;
using System.Collections.Generic;

namespace VisualBasicUpgradeAssistant.Core.DataClasses
{
    /// <summary>
    /// Summary description for ControlListItem.
    /// </summary>
    public class Controltem
    {
        public String VB6 { get; set; }
        public String CSharp { get; set; }
        public Boolean Unsupported { get; set; }
        public Boolean InvisibleAtRuntime { get; set; }
        private List<Controltem> Properties { get; set; }

        public Controltem()
        {
            VB6 = String.Empty;
            CSharp = String.Empty;
            Properties = new List<Controltem>();
        }
    }
}
