using System;
using System.Xml;

namespace VisualBasicUpgradeAssistant.Core.Model
{
    public class XMLConfig
    {
        XmlDocument Doc = new XmlDocument();
        private string msFileName;
        private bool doesExist;

        public string FileName
        {
            get { return msFileName; }
        }

        public XMLConfig(string sFileName)
        {
            msFileName = sFileName;
            try
            {
                Doc.Load(msFileName);
                doesExist = true;
            }
            catch
            {
                Doc.LoadXml("<configuration>" + "</configuration>");
                Doc.Save(msFileName);
            }
        }

        // **********************************************************************************
        //  
        // **********************************************************************************

        public bool ReadBool(string aSection, string aKey, bool bDefaultValue)
        {
            string sResult;

            // return immediately if the file didn't exist
            if (doesExist == false)
                return bDefaultValue;
            if (aSection == "")
                return bDefaultValue;
            if (aKey == "")
                return bDefaultValue;
            sResult = getKeyValue(aSection, aKey, bDefaultValue.ToString());

            if (sResult.ToLower() == "true")
                return true;
            else
                return false;
        }

        public int ReadInt(string aSection, string aKey, int iDefaultValue)
        {
            string sResult;

            // return immediately if the file didn't exist
            if (doesExist == false)
                return iDefaultValue;
            if (aSection == "")
                return iDefaultValue;
            if (aKey == "")
                return iDefaultValue;
            sResult = getKeyValue(aSection, aKey, iDefaultValue.ToString());
            return int.Parse(sResult);
        }

        public string ReadString(string aSection, string aKey, string aDefaultValue)
        {
            // return immediately if the file didn't exist
            if (doesExist == false)
                return aDefaultValue;
            if (aSection == "")
                return aDefaultValue;
            if (aKey == "")
                return aDefaultValue;
            return getKeyValue(aSection, aKey, aDefaultValue);
        }

        // **********************************************************************************
        //  
        // **********************************************************************************

        public bool WriteString(string aSection, string aKey, string sValue)
        {
            return SetKeyValue(aSection, aKey, sValue);
        }

        public bool WriteBool(string aSection, string aKey, bool bValue)
        {
            string sValue;

            if (bValue)
                sValue = "true";
            else
                sValue = "false";

            return SetKeyValue(aSection, aKey, sValue);
        }

        public bool WriteInt(string aSection, string aKey, int iValue)
        {
            return SetKeyValue(aSection, aKey, iValue.ToString());
        }

        // **********************************************************************************
        //  
        // **********************************************************************************

        private bool SetKeyValue(string aSection, string aKey, string aValue)
        {
            XmlNode node1;
            XmlNode node2;
            bool bReturn = false;

            if (aKey == "")
            // find the section, remove all its keys and remove the section
            {
                node1 = Doc.DocumentElement.SelectSingleNode("/configuration/" + aSection);
                // if no such section, return true
                if (node1 == null)
                    return true;                 // remove all its children
                node1.RemoveAll();
                // select its parent ("configuration")
                node2 = Doc.DocumentElement.SelectSingleNode("configuration");
                // remove the section
                node2.RemoveChild(node1);
            }
            else
            {
                if (aValue == "")
                {
                    // find the section of this key
                    node1 = Doc.DocumentElement.SelectSingleNode("/configuration/" + aSection);
                    // return if the section doesn't exist
                    if (node1 == null)
                        return true;                     // find the key
                    node2 = Doc.DocumentElement.SelectSingleNode("/configuration/" + aSection + "/" + aKey);
                    // return true if the key doesn't exist
                    if (node2 == null)
                        return true;                     // remove the key
                    if (node1.RemoveChild(node2) == null)
                        return false;
                }
                else
                {
                    // Both the Key and the Value are filled 
                    // Find the key
                    node1 = Doc.DocumentElement.SelectSingleNode("/configuration/" + aSection + "/" + aKey);
                    if (node1 == null)
                    {
                        // The key doesn't exist: find the section
                        node2 = Doc.DocumentElement.SelectSingleNode("/configuration/" + aSection);
                        if (node2 == null)
                        {
                            // Create the section first
                            XmlElement e = Doc.CreateElement(aSection);
                            // Add the new node at the end of the children of ("configuration")
                            node2 = Doc.DocumentElement.AppendChild(e);
                            // return false if failure
                            if (node2 == null)
                                return false;                             // now create key and value
                            e = Doc.CreateElement(aKey);
                            e.InnerText = aValue;
                            // return false if failure
                            if (node2.AppendChild(e) == null)
                                return false;
                        }
                        else
                        {
                            // Create the key and put the value
                            XmlElement e = Doc.CreateElement(aKey);
                            e.InnerText = aValue;
                            node2.AppendChild(e);
                        }
                    }
                    else
                        // Key exists: set its Value
                        node1.InnerText = aValue;
                }
                // Save the document
                Doc.Save(FileName);
            }
            return bReturn;
        }

        private string getKeyValue(string aSection, string aKey, string aDefaultValue)
        {
            XmlNode node;
            node = Doc.DocumentElement.SelectSingleNode("/configuration/" + aSection + "/" + aKey);
            if (node == null)
                return aDefaultValue;
            return node.InnerText;
        }

        public string[] getchildren(string aNodeName)
        {
            XmlNode node;
            string[] sReturn = new string[0];

            // Select the root if the Node is empty
            if (aNodeName == "")
                node = Doc.DocumentElement;
            else
                // Select the node given
                node = Doc.DocumentElement.SelectSingleNode(aNodeName);

            // exit with an empty collection if nothing here
            if (node == null)
                return sReturn;             // exit with an empty colection if the node has no children
            if (node.HasChildNodes == false)
                return sReturn;             // get the nodelist of all children
            XmlNodeList nodeList = node.ChildNodes;
            int i;
            // transform the Nodelist into an ordinary collection
            sReturn = new string[nodeList.Count];
            for (i = 0; i < nodeList.Count; i++)
                sReturn[i] = nodeList.Item(i).Name;
            return sReturn;
        }
    }
}

