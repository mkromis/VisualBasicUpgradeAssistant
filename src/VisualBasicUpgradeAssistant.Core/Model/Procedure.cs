using System;
using System.Collections;

namespace VisualBasicUpgradeAssistant.Core.Model
{
    /// <summary>
    /// Summary description for Procedure.
    /// </summary>
    /// 

    public enum PROCEDURE_TYPE
    {
        PROCEDURE_EVENT = 1,
        PROCEDURE_SUB = 2,
        PROCEDURE_FUNCTION = 3
    }

    public class Procedure
    {
        private string mName;
        private string mScope;
        private PROCEDURE_TYPE mType;
        private ArrayList mParameterList;
        private string mReturnType;

        private string mComment;
        private ArrayList mLineList;

        public Procedure()
        {
            mLineList = new ArrayList();
            mParameterList = new ArrayList();
        }

        public string Name
        {
            get { return mName; }
            set { mName = value; }
        }

        public string Comment
        {
            get { return mComment; }
            set { mComment = value; }
        }

        public string Scope
        {
            get { return mScope; }
            set { mScope = value; }
        }

        public PROCEDURE_TYPE Type
        {
            get { return mType; }
            set { mType = value; }
        }

        public ArrayList ParameterList
        {
            get { return mParameterList; }
            set { mParameterList = value; }
        }

        public string ReturnType
        {
            get { return mReturnType; }
            set { mReturnType = value; }
        }

        public ArrayList LineList
        {
            get { return mLineList; }
            set { mLineList = value; }
        }
    }
}
