using System;
using System.Collections;

namespace VisualBasicUpgradeAssistant.Core.Model
{
    /// <summary>
    /// Summary description for Control.
    /// </summary>
    public class Control
    {
        private string msName;
        private string msType;
        private string msOwner;
        private ArrayList mPropertyList;
        private bool mContainer;
        private bool mValid;
        private bool mInvisibleAtRuntime;

        public string Name
        {
            get { return msName; }
            set { msName = value; }
        }

        public string Type
        {
            get { return msType; }
            set { msType = value; }
        }

        public string Owner
        {
            get { return msOwner; }
            set { msOwner = value; }
        }

        public bool Container
        {
            get { return mContainer; }
            set { mContainer = value; }
        }

        public bool Valid
        {
            get { return mValid; }
            set { mValid = value; }
        }

        public bool InvisibleAtRuntime
        {
            get { return mInvisibleAtRuntime; }
            set { mInvisibleAtRuntime = value; }
        }

        public Control()
        {
            mPropertyList = new ArrayList();
        }

        public ArrayList PropertyList
        {
            get { return mPropertyList; }
            set { mPropertyList = value; }
        }

        public void PropertyAdd(ControlProperty oProperty)
        {
            mPropertyList.Add(oProperty);
        }

    }
}
