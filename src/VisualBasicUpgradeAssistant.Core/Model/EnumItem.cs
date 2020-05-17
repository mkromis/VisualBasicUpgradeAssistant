using System;

namespace VisualBasicUpgradeAssistant.Core.Model
{
    /// <summary>
    /// Summary description for EnumItem.
    /// </summary>
    public class EnumItem
    {
        private string mName;
        private string mValue;
        private string mComment;

        public string Name
        {
            get { return mName; }
            set { mName = value; }
        }

        public string Value
        {
            get { return mValue; }
            set { mValue = value; }
        }

        public string Comment
        {
            get { return mComment; }
            set { mComment = value; }
        }
    }
}
