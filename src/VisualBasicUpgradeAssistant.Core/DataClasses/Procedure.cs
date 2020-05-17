using System;
using System.Collections;

namespace VisualBasicUpgradeAssistant.Core.DataClasses
{
    /// <summary>
    /// Summary description for Procedure.
    /// </summary>
    /// 

    public enum ProcedureType
    {
        PROCEDURE_EVENT = 1,
        PROCEDURE_SUB = 2,
        PROCEDURE_FUNCTION = 3
    }

    public class Procedure
    {
        public Procedure()
        {
            LineList = new ArrayList();
            ParameterList = new ArrayList();
        }

        public String Name { get; set; }

        public String Comment { get; set; }

        public String Scope { get; set; }

        public ProcedureType Type { get; set; }

        public ArrayList ParameterList { get; set; }

        public String ReturnType { get; set; }

        public ArrayList LineList { get; set; }
    }
}
