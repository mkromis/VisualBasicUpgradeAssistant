using System;
using System.Xml;

namespace VisualBasicUpgradeAssistant.Core.Model
{
    public class XMLConfig
    {
        private XmlDocument _doc = new XmlDocument();
        private Boolean _doesExist;

        public String FileName { get; }

        public XMLConfig(String fileName)
        {
            FileName = fileName;
            try
            {
                _doc.Load(FileName);
                _doesExist = true;
            }
            catch
            {
                _doc.LoadXml("<configuration>" + "</configuration>");
                _doc.Save(FileName);
            }
        }

        // **********************************************************************************
        //  
        // **********************************************************************************

        public Boolean ReadBool(String section, String key, Boolean defaultValue)
        {
            String result;

            // return immediately if the file didn't exist
            if (_doesExist == false)
                return defaultValue;
            if (section == "")
                return defaultValue;
            if (key == "")
                return defaultValue;
            result = GetKeyValue(section, key, defaultValue.ToString());

            return result.ToLower() == "true";
        }

        public Int32 ReadInt(String section, String key, Int32 defaultValue)
        {
            String result;

            // return immediately if the file didn't exist
            if (_doesExist == false)
                return defaultValue;
            if (section == "")
                return defaultValue;
            if (key == "")
                return defaultValue;
            result = GetKeyValue(section, key, defaultValue.ToString());
            return Int32.Parse(result);
        }

        public String ReadString(String section, String key, String defaultValue)
        {
            // return immediately if the file didn't exist
            if (_doesExist == false)
                return defaultValue;
            if (section == "")
                return defaultValue;
            if (key == "")
                return defaultValue;
            return GetKeyValue(section, key, defaultValue);
        }

        // **********************************************************************************
        //  
        // **********************************************************************************

        public Boolean WriteString(String section, String key, String value)
        {
            return SetKeyValue(section, key, value);
        }

        public Boolean WriteBool(String section, String key, Boolean value)
        {
            String sValue;

            if (value)
                sValue = "true";
            else
                sValue = "false";

            return SetKeyValue(section, key, sValue);
        }

        public Boolean WriteInt(String section, String key, Int32 iValue)
        {
            return SetKeyValue(section, key, iValue.ToString());
        }

        // **********************************************************************************
        //  
        // **********************************************************************************

        private Boolean SetKeyValue(String section, String key, String value)
        {
            XmlNode node1;
            XmlNode node2;
            Boolean result = false;

            if (key == "")
            // find the section, remove all its keys and remove the section
            {
                node1 = _doc.DocumentElement.SelectSingleNode("/configuration/" + section);
                // if no such section, return true
                if (node1 == null)
                    return true;                 // remove all its children
                node1.RemoveAll();
                // select its parent ("configuration")
                node2 = _doc.DocumentElement.SelectSingleNode("configuration");
                // remove the section
                node2.RemoveChild(node1);
            }
            else
            {
                if (value == "")
                {
                    // find the section of this key
                    node1 = _doc.DocumentElement.SelectSingleNode("/configuration/" + section);
                    // return if the section doesn't exist
                    if (node1 == null)
                        return true;                     // find the key
                    node2 = _doc.DocumentElement.SelectSingleNode("/configuration/" + section + "/" + key);
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
                    node1 = _doc.DocumentElement.SelectSingleNode("/configuration/" + section + "/" + key);
                    if (node1 == null)
                    {
                        // The key doesn't exist: find the section
                        node2 = _doc.DocumentElement.SelectSingleNode("/configuration/" + section);
                        if (node2 == null)
                        {
                            // Create the section first
                            XmlElement e = _doc.CreateElement(section);
                            // Add the new node at the end of the children of ("configuration")
                            node2 = _doc.DocumentElement.AppendChild(e);
                            // return false if failure
                            if (node2 == null)
                                return false;                             // now create key and value
                            e = _doc.CreateElement(key);
                            e.InnerText = value;
                            // return false if failure
                            if (node2.AppendChild(e) == null)
                                return false;
                        }
                        else
                        {
                            // Create the key and put the value
                            XmlElement e = _doc.CreateElement(key);
                            e.InnerText = value;
                            node2.AppendChild(e);
                        }
                    }
                    else
                        // Key exists: set its Value
                        node1.InnerText = value;
                }
                // Save the document
                _doc.Save(FileName);
            }
            return result;
        }

        private String GetKeyValue(String section, String key, String defaultValue)
        {
            XmlNode node;
            node = _doc.DocumentElement.SelectSingleNode("/configuration/" + section + "/" + key);
            if (node == null)
                return defaultValue;
            return node.InnerText;
        }

        public String[] GetChildren(String nodeName)
        {
            XmlNode node;
            String[] result = new String[0];

            // Select the root if the Node is empty
            if (nodeName == "")
                node = _doc.DocumentElement;
            else
                // Select the node given
                node = _doc.DocumentElement.SelectSingleNode(nodeName);

            // exit with an empty collection if nothing here
            if (node == null)
                return result;             // exit with an empty colection if the node has no children
            if (node.HasChildNodes == false)
                return result;             // get the nodelist of all children
            XmlNodeList nodeList = node.ChildNodes;
            Int32 i;
            // transform the Nodelist into an ordinary collection
            result = new String[nodeList.Count];
            for (i = 0; i < nodeList.Count; i++)
                result[i] = nodeList.Item(i).Name;
            return result;
        }
    }
}

