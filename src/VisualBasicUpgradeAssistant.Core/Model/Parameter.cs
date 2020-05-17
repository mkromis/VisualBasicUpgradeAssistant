using System;

namespace VisualBasicUpgradeAssistant.Core.Model
{
    /// <summary>
    /// TODO - Add class summary
    /// </summary>
    public class Parameter
    {
        private string mName;
        private string mType;
        private string mPass;

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

        public string Pass
        {
            get { return mPass; }
            set { mPass = value; }
        }
    }
}
