using System;

namespace VB2C
{
    /// <summary>
    /// Summary description for ControlListItem.
    /// </summary>
    public class ControlListItem
    {
        private string mVB6Name;
        private string mCsharpName;
        private bool mUnsupported;
        private bool mInvisibleAtRuntime;

        public string VB6Name
        {
            get { return mVB6Name; }
            set { mVB6Name = value; }
        }
        public string CsharpName
        {
            get { return mCsharpName; }
            set { mCsharpName = value; }
        }
        public bool Unsupported
        {
            get { return mUnsupported; }
            set { mUnsupported = value; }
        }
        public bool InvisibleAtRuntime
        {
            get { return mInvisibleAtRuntime; }
            set { mInvisibleAtRuntime = value; }
        }
    }
}
