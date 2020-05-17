using System;
using System.Collections;

namespace VisualBasicUpgradeAssistant.Core.DataClasses
{
    /// <summary>
    /// Summary description for Module.
    /// </summary>
    public class Module
    {
        public String Name { get; set; }
        public String Type { get; set; }
        public String Version { get; set; }
        public String FileName { get; set; }
        public Boolean ImagesUsed { get; set; } = false;
        public Boolean MenuUsed { get; set; } = false;

        public Module()
        {
            FormPropertyList = new ArrayList();
            ControlList = new ArrayList();
            ImageList = new ArrayList();
            VariableList = new ArrayList();
            PropertyList = new ArrayList();
            ProcedureList = new ArrayList();
            EnumList = new ArrayList();
        }

        public ArrayList ImageList { get; set; }
        public ArrayList ControlList { get; set; }
        public ArrayList FormPropertyList { get; set; }
        public ArrayList VariableList { get; set; }
        public ArrayList PropertyList { get; set; }
        public ArrayList ProcedureList { get; set; }
        public ArrayList EnumList { get; set; }

        public void FormPropertyAdd(ControlProperty oProperty)
        {
            FormPropertyList.Add(oProperty);
        }

        public void ControlAdd(ControlType oControl)
        {
            ControlList.Add(oControl);
        }

        public void VariableAdd(Variable oVariable)
        {
            VariableList.Add(oVariable);
        }

        public void PropertyAdd(Property oProperty)
        {
            PropertyList.Add(oProperty);
        }

        public void ProcedureAdd(Procedure oProcedure)
        {
            ProcedureList.Add(oProcedure);
        }
    }
}
