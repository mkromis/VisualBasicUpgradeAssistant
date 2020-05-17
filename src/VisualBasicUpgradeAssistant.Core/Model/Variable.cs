using System;

namespace VisualBasicUpgradeAssistant.Core.Model
{
    /// <summary>
    /// TODO - Add class summary
    /// </summary>
    public class Variable
    {
        private string mName;
        private string mType;
        private string mScope;
        private string mComment;

        public string Name
        {
            get { return mName; }
            set { mName = value; }
        }

        public string Type
        {
            get { return mType; }
            set { mType = value; }
        }

        public string Scope
        {
            get { return mScope; }
            set { mScope = value; }
        }

        public string Comment
        {
            get { return mComment; }
            set { mComment = value; }
        }
    }
}
