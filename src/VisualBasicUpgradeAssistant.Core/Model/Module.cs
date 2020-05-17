using System;
using System.Collections;

namespace VB2C
{
    /// <summary>
    /// Summary description for Module.
    /// </summary>
    public class Module
    {
        private string mName;
        private string mFileName;
        private string mType;
        private string mVersion;
        private bool mImagesUsed = false;
        private bool mMenuUsed = false;
        private ArrayList mFormPropertyList;
        private ArrayList mControlList;
        private ArrayList mImageList;
        private ArrayList mVariableList;
        private ArrayList mPropertyList;
        private ArrayList mProcedureList;
        private ArrayList mEnumList;

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

        public string Version
        {
            get { return mVersion; }
            set { mVersion = value; }
        }

        public string FileName
        {
            get { return mFileName; }
            set { mFileName = value; }
        }

        public bool ImagesUsed
        {
            get { return mImagesUsed; }
            set { mImagesUsed = value; }
        }

        public bool MenuUsed
        {
            get { return mMenuUsed; }
            set { mMenuUsed = value; }
        }

        public Module()
        {
            mFormPropertyList = new ArrayList();
            mControlList = new ArrayList();
            mImageList = new ArrayList();
            mVariableList = new ArrayList();
            mPropertyList = new ArrayList();
            mProcedureList = new ArrayList();
            mEnumList = new ArrayList();
        }

        public ArrayList ImageList
        {
            get { return mImageList; }
            set { mImageList = value; }
        }

        public ArrayList ControlList
        {
            get { return mControlList; }
            set { mControlList = value; }
        }

        public ArrayList FormPropertyList
        {
            get { return mFormPropertyList; }
            set { mFormPropertyList = value; }
        }

        public ArrayList VariableList
        {
            get { return mVariableList; }
            set { mVariableList = value; }
        }

        public ArrayList PropertyList
        {
            get { return mPropertyList; }
            set { mPropertyList = value; }
        }

        public ArrayList ProcedureList
        {
            get { return mProcedureList; }
            set { mProcedureList = value; }
        }

        public ArrayList EnumList
        {
            get { return mEnumList; }
            set { mEnumList = value; }
        }

        public void FormPropertyAdd(ControlProperty oProperty)
        {
            mFormPropertyList.Add(oProperty);
        }

        public void ControlAdd(Control oControl)
        {
            mControlList.Add(oControl);
        }

        public void VariableAdd(Variable oVariable)
        {
            mVariableList.Add(oVariable);
        }

        public void PropertyAdd(Property oProperty)
        {
            mPropertyList.Add(oProperty);
        }

        public void ProcedureAdd(Procedure oProcedure)
        {
            mProcedureList.Add(oProcedure);
        }
    }
}
