using System;
using System.Collections;
using System.Drawing;
using System.IO;
using System.Text; 
using System.Xml;

namespace VB2C
{
	/// <summary>
	/// Summary description for Tools.
	/// </summary>
	sealed class Number
	{
    public static bool IsOdd(int iNumber)
    {
      if (iNumber != (iNumber/2)*2)
        {return true;}
      else
        {return false;}
    }
  }
  
  sealed class Debug
  {
    public static void WriteLine(string Message)
    {
      MessageBox.Show(Message ,"", MessageBoxButtons.OK ,MessageBoxIcon.Error);      
    }
	}
	
	// parse VB6 properties and values to C#
  sealed class Tools
  {
    private static Hashtable mControlList;

    private static void ControlListLoad()
    {
      mControlList = new Hashtable();
      XmlDocument Doc = new XmlDocument();
      XmlNode node;
      ControlListItem oItem;

      // get current directory
      string[] CommandLineArgs;
      CommandLineArgs = Environment.GetCommandLineArgs();
      // index 0 contain path and name of exe file
      string BinPath = Path.GetDirectoryName(CommandLineArgs[0].ToLower());
      string FileName = BinPath + @"\vb2c.xml"; 

      Doc.Load(FileName);
      // Select the node given
      node = Doc.DocumentElement.SelectSingleNode("/configuration/ControlList");
      // exit with an empty collection if nothing here
      if (node == null) {return;}
      // exit with an empty colection if the node has no children
      if (node.HasChildNodes == false ) {return;}
      // get the nodelist of all children
      XmlNodeList nodeList = node.ChildNodes;
        
      foreach( XmlElement element in nodeList )
      {
        oItem = new ControlListItem();
        oItem.VB6Name = String.Empty;
        oItem.CsharpName = String.Empty;
        oItem.Unsupported = false;
        oItem.InvisibleAtRuntime = false;
        foreach ( XmlElement childElement in element)
        {
          switch(childElement.Name)
          {
            case "VB6":
              // compare in uppercase
              oItem.VB6Name = childElement.InnerText.ToUpper();
              break;
            case "Csharp":
              oItem.CsharpName = childElement.InnerText;
              break;
            case "Unsupported":
              oItem.Unsupported = bool.Parse(childElement.InnerText);
              break;
            case "InvisibleAtRuntime":
              oItem.InvisibleAtRuntime = bool.Parse(childElement.InnerText);
              break;
          }          
        }
        mControlList.Add(oItem.VB6Name, oItem);
      }


//      private string getKeyValue(string aSection, string aKey, string aDefaultValue)
//      {
//        XmlNode node;
//        node = (Doc.DocumentElement).SelectSingleNode("/configuration/" + aSection + "/" + aKey);
//        if (node == null) {return aDefaultValue;}
//        return node.InnerText;
//      }
      
    }

    public static bool ParseModule( Module SourceModule, Module TargetModule )
    {

      ControlListLoad();

      // module name
      TargetModule.Name = SourceModule.Name; 
      // file name
      TargetModule.FileName = Path.GetFileNameWithoutExtension(SourceModule.FileName) + ".cs";
      // type
      TargetModule.Type = SourceModule.Type; 
      // version
      TargetModule.Version = SourceModule.Version; 
      // process own properties - forms
      Tools.ParseModuleProperties(TargetModule, SourceModule.FormPropertyList, TargetModule.FormPropertyList );
      // process controls - form
      Tools.ParseControls(TargetModule, SourceModule.ControlList, TargetModule.ControlList ); 
      
      // special exception for menu
      if (TargetModule.MenuUsed)
      {
        // add main menu control
        Control oControl = new Control(); 
        oControl.Name = "MainMenu";
        oControl.Owner = TargetModule.Name;
        oControl.Type = "MainMenu";
        oControl.Valid = true;
        oControl.InvisibleAtRuntime = true;
        TargetModule.ControlList.Insert(0, oControl); 
        foreach ( Control oMenuControl in TargetModule.ControlList )
        {
          if ((oMenuControl.Type == "MenuItem") && (oMenuControl.Owner == TargetModule.Name))
          {
            // rewrite previous owner
            oMenuControl.Owner = oControl.Name; 
          }
        }
      }

      ArrayList TempControlList = new ArrayList();
      int TabControlIndex = 0;

      // check for TabDlg.SSTab
      foreach ( Control oTargetControl in TargetModule.ControlList  )
      {
        if ((oTargetControl.Type == "TabControl") && (oTargetControl.Valid))
        {
          // for each source table is necessary
//          this.tabControl1 = new System.Windows.Forms.TabControl();
//          this.tabPage1 = new System.Windows.Forms.TabPage();

          int Index = 0;
          Control oTabPage = null;
          // each property  
          foreach ( ControlProperty oTargetProperty in oTargetControl.PropertyList )
          {
            // TabCaption = create new tab
            //      this.SSTab1.(TabCaption(0)) = "Tab 0";

            Console.WriteLine(oTargetProperty.Name); 

            if (oTargetProperty.Name.IndexOf( "TabCaption(" + Index.ToString() + ")"  ,0) > -1)
            {
              // new tab
              oTabPage = new Control(); 
              oTabPage.Type = "TabPage";
              oTabPage.Name = "tabPage" + Index.ToString();
              oTabPage.Owner = oTargetControl.Name; 
              oTabPage.Container = true;
              oTabPage.Valid = true;
              oTabPage.InvisibleAtRuntime = false;

              // add some necessary properties
              ControlProperty TargetProperty = new ControlProperty();
              TargetProperty.Name = "Location";
              TargetProperty.Value = "new System.Drawing.Point(4, 22)";
              TargetProperty.Valid = true;
              oTabPage.PropertyList.Add(TargetProperty); 

              TargetProperty = new ControlProperty();
              TargetProperty.Name = "Size";
              TargetProperty.Value = "new System.Drawing.Size(477, 374)";
              TargetProperty.Valid = true;
              oTabPage.PropertyList.Add(TargetProperty); 

              TargetProperty = new ControlProperty();
              TargetProperty.Name = "Text";
              TargetProperty.Value = oTargetProperty.Value;
              TargetProperty.Valid = true;
              oTabPage.PropertyList.Add(TargetProperty); 

              TargetProperty = new ControlProperty();
              TargetProperty.Name = "TabIndex";
              TargetProperty.Value = Index.ToString();
              TargetProperty.Valid = true;
              oTabPage.PropertyList.Add(TargetProperty); 

              TempControlList.Add(oTabPage); 
              Index ++;
            }

            // Control = change owner of control to current tab
            //      this.SSTab1.(Tab(0).Control(0) = "ImageControl";
            if (oTargetProperty.Name.IndexOf( ".Control(", 0) > -1)
            {
              if ( oTargetProperty.Name.IndexOf("Enable",0) == -1 )
              {
                string TabName = oTargetProperty.Value.Substring(1,oTargetProperty.Value.Length -2);
                TabName = GetControlIndexName(TabName);
                // search for "oTargetProperty.Value" control
                // and replace owner of this control to current tab
                foreach ( Control oNewOwner in TargetModule.ControlList  )
                {
                  if ((oNewOwner.Name == TabName) && ( !oNewOwner.InvisibleAtRuntime ))
                  {
                    oNewOwner.Owner = oTabPage.Name;
                  }
                }
              }
            }
          }
        }
        TabControlIndex ++;
      }

      if (TempControlList.Count > 0)
      {
        // right order of tabs
        int Position = 0;
        foreach ( Control oControl in TempControlList  )
        {
          TargetModule.ControlList.Insert(TabControlIndex + Position ,oControl); 
          Position ++;
        }
      }


      // process enums
      Tools.ParseEnums( SourceModule, TargetModule ); 
      // process variables
      Tools.ParseVariables( SourceModule.VariableList, TargetModule.VariableList );       
      // process properties
      Tools.ParseClassProperties( SourceModule, TargetModule );  
      // process procedures
      Tools.ParseProcedures( SourceModule, TargetModule );    


      return true;
    }

    // return control name
    private static string GetControlIndexName(string TabName)
    {
      //  this.SSTab1.(Tab(1).Control(4) = "Option1(0)";
      int Start = 0;
      int End = 0;

      Start = TabName.IndexOf("(",0);
      if ( Start > -1)
      {
        End = TabName.IndexOf(")",0);
        return TabName.Substring(0, Start ) + TabName.Substring( Start + 1, End - Start - 1); 
      }
      else
      {
        return TabName;
      }

      
    }

    public static bool ParseControls( Module oModule, ArrayList SourceControlList, ArrayList TargetControlList )
    {
      string Type = String.Empty;

      foreach ( Control oSourceControl in SourceControlList )
      {
        Control oTargetControl = new Control();

        oTargetControl.Name = oSourceControl.Name;
        oTargetControl.Owner = oSourceControl.Owner;
        oTargetControl.Container = oSourceControl.Container;
        oTargetControl.Valid = true;

        // compare upper case type
        if ( mControlList.ContainsKey(oSourceControl.Type.ToUpper()) )
        {
          ControlListItem oItem = (ControlListItem) mControlList[oSourceControl.Type.ToUpper()];

          if (oItem.Unsupported)
          {
            Type = "Unsuported";
            oTargetControl.Valid = false;
          }
          else
          {
            Type = oItem.CsharpName;
            if (Type == "MenuItem")
            {
              oModule.MenuUsed = true;
            }
          }
          oTargetControl.InvisibleAtRuntime = oItem.InvisibleAtRuntime;
        }
        else
        {
          Type = oSourceControl.Type;
        }

        oTargetControl.Type = Type;
        ParseControlProperties( oModule, oTargetControl, oSourceControl.PropertyList, oTargetControl.PropertyList );  
        
        TargetControlList.Add(oTargetControl);
      }
      return true;
    }

    public static bool ParseModuleProperties( Module oModule, 
                                              ArrayList SourcePropertyList, 
                                              ArrayList TargetPropertyList )
    {
      ControlProperty TargetProperty = null;

      // each property  
      foreach ( ControlProperty SourceProperty in SourcePropertyList )
      {
        TargetProperty = new ControlProperty();
        if (ParseProperties(oModule.Type, SourceProperty, TargetProperty, SourcePropertyList))
        {
          if (TargetProperty.Name == "BackgroundImage" || TargetProperty.Name == "Icon")
          {
            oModule.ImagesUsed = true;          
          }
          TargetPropertyList.Add(TargetProperty);
        }
      }
      return true;
    }

    public static bool ParseEnums( Module SourceModule, Module TargetModule )
    {
      foreach ( Enum SourceEnum in SourceModule.EnumList )
      {
        TargetModule.EnumList.Add(SourceEnum);
      }
      return true;
    }

    public static bool ParseVariables( ArrayList SourceVariableList, ArrayList TargetVariableList )
    {
      Variable TargetVariable = null;

      // each property  
      foreach ( Variable SourceVariable in SourceVariableList )
      {
        TargetVariable = new Variable();
        if (ParseVariable(SourceVariable, TargetVariable))
        {
          TargetVariableList.Add(TargetVariable);
        }
      }
      return true;
    }

    private static string VariableTypeConvert(string SourceType)
    {
      string TargetType;

      switch(SourceType)
      {
        case "Long":
          TargetType = "int";
          break;
        case "Integer":
          TargetType = "short";
          break;
        case "Byte":
          TargetType = "byte";
          break;
        case "String":
          TargetType = "string";
          break;
        case "Boolean":
          TargetType = "bool";
          break;
        case "Currency":
          TargetType = "decimal";
          break;
        case "Single":
          TargetType = "float";
          break;
        case "Double":
          TargetType = "double";
          break;
        case "ADODB.Recordset":
        case "DAO.Recordset":
        case "Recordset":
          TargetType = "DataReader";
          break;
        default:
          TargetType = SourceType;
          break;
      }
      return TargetType;
    }

    public static bool ParseVariable(Variable SourceVariable, Variable TargetVariable)
    {

      TargetVariable.Scope = SourceVariable.Scope; 
      TargetVariable.Name = SourceVariable.Name;
      TargetVariable.Type = VariableTypeConvert(SourceVariable.Type);

      return true;
    }

    public static bool ParseClassProperties( Module SourceModule, Module TargetModule )
    {
      Property TargetProperty;

      foreach ( Property SourceProperty in SourceModule.PropertyList )
      {
        TargetProperty = new Property();

        TargetProperty.Name = SourceProperty.Name ;
        TargetProperty.Comment = SourceProperty.Comment ;
        switch (SourceProperty.Direction)
        {
          case "Get":
            TargetProperty.Direction = "get";
            break;
          case "Set":
          case "Let":
            TargetProperty.Direction = "set";
            break;
        }
        TargetProperty.Scope = SourceProperty.Scope; 
        TargetProperty.Type = VariableTypeConvert(SourceProperty.Type);
        // lines
        foreach ( string Line in SourceProperty.LineList )
        {
          if (Line.Trim() != String.Empty )
          {
            TargetProperty.LineList.Add(Line);
          }
        }




        TargetModule.PropertyList.Add(TargetProperty);
      }
      return true;
    }

    public static bool ParseProcedures( Module SourceModule, Module TargetModule )
    {
      const string Indent6 = "      ";
      Procedure TargetProcedure;
      string Temp;

      foreach ( Procedure SourceProcedure in SourceModule.ProcedureList )
      {
        TargetProcedure = new Procedure();

        TargetProcedure.Name = SourceProcedure.Name;
        TargetProcedure.Scope = SourceProcedure.Scope;
        TargetProcedure.Comment = SourceProcedure.Comment;
        TargetProcedure.Type = SourceProcedure.Type; 
        TargetProcedure.ReturnType = VariableTypeConvert(SourceProcedure.ReturnType); 

        TargetProcedure.ParameterList = SourceProcedure.ParameterList; 

        // lines
        foreach ( string Line in SourceProcedure.LineList )
        {
          Temp = Line.Trim(); 
          if (Temp.Length > 0)
          {
            string TempLine = String.Empty;
            // vbNullString = String.Empty
            if ( Temp.IndexOf("vbNullString", 0) > -1)
            {
              TempLine = Temp.Replace( "vbNullString", "String.Empty");
              Temp = TempLine;
            }
            // Nothing = null
            if ( Temp.IndexOf("Nothing", 0) > -1)
            {
              TempLine = Temp.Replace( "Nothing", "null");
              Temp = TempLine;
            }
            // Set
            if ( Temp.IndexOf("Set ", 0) > -1)
            {
              TempLine = Temp.Replace( "Set ", " ");
              Temp = TempLine;
            }
            // remark
            if ( Temp[0] == (char) 39 ) // '
            {
              TempLine = Temp.Replace( "'", "//");
              Temp = TempLine;
            }
            // & to +
            if ( Temp.IndexOf("&", 0) > -1)
            {
              TempLine = Temp.Replace( "&", "+");
              Temp = TempLine;
            }
            // Select Case
            if ( Temp.IndexOf("Select Case", 0) > -1)
            {
              TempLine = Temp.Replace( "Select Case", "switch");
              Temp = TempLine;
            }
            // End Select
            if ( Temp.IndexOf("End Select", 0) > -1)
            {
              TempLine = Temp.Replace( "End Select", "}");
              Temp = TempLine;
            }
            // _
            if ( Temp.IndexOf(" _", 0) > -1)
            {
              TempLine = Temp.Replace( " _", "\r\n");
              Temp = TempLine;
            }
            // If
            if ( Temp.IndexOf("If ", 0) > -1)
            {
              TempLine = Temp.Replace( "If ", "if ( ");
              Temp = TempLine;
            }
            // Not
            if ( Temp.IndexOf("Not ", 0) > -1)
            {
              TempLine = Temp.Replace( "Not ", "! ");
              Temp = TempLine;
            }
            // then
            if ( Temp.IndexOf(" Then", 0) > -1)
            {
              TempLine = Temp.Replace(" Then", " )\r\n" + Indent6 + "{\r\n");
              Temp = TempLine;
            }
            // else
            if ( Temp.IndexOf("Else", 0) > -1)
            {
              TempLine = Temp.Replace( "Else", "}\r\n" + Indent6 + "else\r\n" + Indent6 + "{");
              Temp = TempLine;
            }
            // End if
            if ( Temp.IndexOf("End If", 0) > -1)
            {
              TempLine = Temp.Replace( "End If", "}");
              Temp = TempLine;
            }
            // Unload Me
            if ( Temp.IndexOf("Unload Me", 0) > -1)
            {
              TempLine = Temp.Replace( "Unload Me", "Close()");
              Temp = TempLine;
            }
            // .Caption
            if ( Temp.IndexOf(".Caption", 0) > -1)
            {
              TempLine = Temp.Replace( ".Caption", ".Text");
              Temp = TempLine;
            }
            // True
            if ( Temp.IndexOf("True", 0) > -1)
            {
              TempLine = Temp.Replace( "True", "true");
              Temp = TempLine;
            }
            // False
            if ( Temp.IndexOf("False", 0) > -1)
            {
              TempLine = Temp.Replace( "False", "false");
              Temp = TempLine;
            }

            // New
            if ( Temp.IndexOf("New", 0) > -1)
            {
              TempLine = Temp.Replace( "New", "new");
              Temp = TempLine;
            }

            if (TempLine == String.Empty)
            {
              TargetProcedure.LineList.Add(Temp);
            }
            else
            {
              TargetProcedure.LineList.Add(TempLine);
            }
            
          }
          else
          {
            TargetProcedure.LineList.Add(String.Empty);
          }
        }

        TargetModule.ProcedureList.Add(TargetProcedure);
      }
      return true;
    }


    public static bool ParseControlProperties( Module oModule, Control oControl, 
                                                ArrayList SourcePropertyList, 
                                                ArrayList TargetPropertyList )
    {
      ControlProperty TargetProperty = null;

      // each property  
      foreach ( ControlProperty SourceProperty in SourcePropertyList )
      {
        if (SourceProperty.Name == "Index")
        {
          // Index           =   3
          oControl.Name = oControl.Name + SourceProperty.Value;
        }
        else
        {
          TargetProperty = new ControlProperty();
          if (ParseProperties(oControl.Type, SourceProperty, TargetProperty, SourcePropertyList))
          {
            if (TargetProperty.Name == "Image")
            {
              oModule.ImagesUsed = true;          
            }
            TargetPropertyList.Add(TargetProperty);
          }
        }
      }
      return true;
    }

    public static bool ParseProperties(string Type, 
                                        ControlProperty SourceProperty, 
                                        ControlProperty TargetProperty,
                                        ArrayList SourcePropertyList)
    {
      
      bool ValidProperty = false;
      
      ValidProperty = true;
        
      TargetProperty.Valid = true;
        
      switch(SourceProperty.Name)
      {
          // not used
        case "Appearance":  
        case "ScaleHeight":
        case "ScaleWidth":   
        case "Style":             // button  
        case "BackStyle":         //label
        case "IMEMode":
        case "WhatsThisHelpID":
        case "Mask":              // maskedit
        case "PromptChar":        // maskedit

          ValidProperty = false;
          break;
                            
          // begin common properties

        case "Alignment":
          //              0 - left
          //              1 - right
          //              2 - center
          TargetProperty.Name = "TextAlign";
          TargetProperty.Value = "System.Drawing.ContentAlignment.";
        switch (SourceProperty.Value)
        {
          case "0":
            TargetProperty.Value = TargetProperty.Value + "TopLeft";
            break;
          case "1":
            TargetProperty.Value = TargetProperty.Value + "TopRight";
            break;
          case "2":
          default:
            TargetProperty.Value = TargetProperty.Value + "TopCenter";
            break;
        }
          break;

        case "BackColor":
        case "ForeColor":
          if (Type != "ImageList")
          {
            TargetProperty.Name = SourceProperty.Name;
            TargetProperty.Value = GetColor(SourceProperty.Value);                  
          }
          else
          {
            ValidProperty = false;  
          }
          break;

        case "BorderStyle":
          if (Type == "form")
          {
            TargetProperty.Name = "FormBorderStyle";
            // 0 - none
            // 1 - fixed single
            // 2 - sizable
            // 3 - fixed dialog
            // 4 - fixed toolwindow
            // 5 - sizable toolwindow

            // FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;

            TargetProperty.Value = "System.Windows.Forms.FormBorderStyle.";
            switch (SourceProperty.Value)
            {
              case "0":
                TargetProperty.Value = TargetProperty.Value + "None";
                break;
              default:
              case "1":
                TargetProperty.Value = TargetProperty.Value + "FixedSingle";
                break;
              case "2":
                TargetProperty.Value = TargetProperty.Value + "Sizable";
                break;
              case "3":
                TargetProperty.Value = TargetProperty.Value + "FixedDialog";
                break;
              case "4":
                TargetProperty.Value = TargetProperty.Value + "FixedToolWindow";
                break;
              case "5":
                TargetProperty.Value = TargetProperty.Value + "SizableToolWindow";
                break;
            }
          }
          else
          {
            TargetProperty.Name = SourceProperty.Name;
            TargetProperty.Value = "System.Windows.Forms.BorderStyle.";
            switch (SourceProperty.Value)
            {
              case "0":
                TargetProperty.Value = TargetProperty.Value + "None";
                break;
              case "1":
                TargetProperty.Value = TargetProperty.Value + "FixedSingle";
                break;
              case "2":
              default:
                TargetProperty.Value = TargetProperty.Value + "Fixed3D";
                break;
            }
          }
          break;
              
        case "Caption":
        case "Text":
          TargetProperty.Name = "Text";
          TargetProperty.Value = SourceProperty.Value;
          break;
              
          // this.cmdExit.Size = new System.Drawing.Size(80, 40);              
        case "Height":
          TargetProperty.Name = "Size";
          TargetProperty.Value = "new System.Drawing.Size(" + GetSize("Height","Width", SourcePropertyList) + ")";
          break;
              
          // this.cmdExit.Location = new System.Drawing.Point(616, 520);
        case "Left":
          if ((Type != "ImageList") && (Type != "Timer"))
          {
            TargetProperty.Name = "Location";
            TargetProperty.Value = "new System.Drawing.Point(" + GetLocation(SourcePropertyList) + ")";
          }
          else
          {
            ValidProperty = false;  
          }
          break;
        case "Top":
        case "Width":
          // nothing, already processed by Height, Left
          ValidProperty = false;
          break;

        case "Enabled":
        case "Locked":
        case "TabStop":
        case "Visible":
        case "UseMnemonic":
          TargetProperty.Name = SourceProperty.Name;
          TargetProperty.Value = GetBool(SourceProperty.Value);
          break;

        case "WordWrap":
          if ( Type == "Text")
          {
            TargetProperty.Name = SourceProperty.Name;
            TargetProperty.Value = GetBool(SourceProperty.Value);
          }
          else
          {
            ValidProperty = false;
          }
          break;

        case "Font":
          ConvertFont(SourceProperty, TargetProperty);
          break;
          // end common properties

        case "MaxLength":
          TargetProperty.Name = SourceProperty.Name;
          TargetProperty.Value = SourceProperty.Value;
          break;

          // PasswordChar
        case "PasswordChar":
          TargetProperty.Name = SourceProperty.Name;
          // PasswordChar = '*';
          TargetProperty.Value = "'" + SourceProperty.Value.Substring(1,1)  + "'";
          break;

          // Value
        case "Value":
        switch (Type)
        {
          case "RadioButton":
            // .Checked = true;
            TargetProperty.Name = "Checked";
            TargetProperty.Value = GetBool(SourceProperty.Value);
            break;
          case "CheckBox":
            //.CheckState = System.Windows.Forms.CheckState.Checked;
            TargetProperty.Name = "CheckState";
            TargetProperty.Value = "System.Windows.Forms.CheckState.";
            // 0 - Unchecked
            // 1 - checked 
            // 2 - grayed
          switch (SourceProperty.Value)
          {
            default:
            case "0":
              TargetProperty.Value = TargetProperty.Value + "Unchecked";
              break;
            case "1":
              TargetProperty.Value = TargetProperty.Value + "Checked";
              break;
            case "2":
              TargetProperty.Value = TargetProperty.Value + "Indeterminate";
              break;
          }
            break;
          default:
            TargetProperty.Value = TargetProperty.Value + "Both";
            break;
        }
          break;

          // timer
        case "Interval":
          TargetProperty.Name = "Interval";
          TargetProperty.Value = SourceProperty.Value;
          break;  

          // this.cmdExit.Anchor = (System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right);
        case "Cancel": 
          if (int.Parse(SourceProperty.Value) != 0)
          {
            TargetProperty.Name = "DialogResult";
            TargetProperty.Value = "System.Windows.Forms.DialogResult.Cancel";
          }
          break;
        case "Default":  
          if (int.Parse(SourceProperty.Value) != 0)
          {
            TargetProperty.Name = "DialogResult";
            TargetProperty.Value = "System.Windows.Forms.DialogResult.OK";
          }
          break;              

          //                this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
          //                this.ClientSize = new System.Drawing.Size(704, 565);
          //                this.MinimumSize = new System.Drawing.Size(712, 592);

              
          // direct value    
        case "TabIndex":
        case "Tag":
          // except MenuItem
          if (Type != "MenuItem")
          {
            TargetProperty.Name = SourceProperty.Name;
            TargetProperty.Value = SourceProperty.Value;            
          }
          else
          {
            ValidProperty = false;  
          }                
          break;

          // -1 converted to true
          // 0 to false
        case "AutoSize":
          // only for Label
          if (Type == "Label")
          {
            TargetProperty.Name = SourceProperty.Name;
            TargetProperty.Value = GetBool(SourceProperty.Value);                
          }
          else
          {
            ValidProperty = false;  
          }                
          break;

        case "Icon":     
          // "Form1.frx":0000;
          // exist file ?

//          System.Drawing.Bitmap pic = null;
//          GetFRXImage(@"C:\temp\test\form1.frx", 0x13960, pic );

          if (Type == "form")
          {           
            //.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            TargetProperty.Name = "Icon";
            TargetProperty.Value = SourceProperty.Value;
          }
          else
          {
            // ((System.Drawing.Bitmap)(resources.GetObject("Command1.Image")));
            TargetProperty.Name = "Image";
            TargetProperty.Value = SourceProperty.Value;
          }
          break;

        case "Picture":
          // = "Form1.frx":13960;
          if (Type == "form")
          {
            // ((System.Drawing.Bitmap)(resources.GetObject("$this.BackgroundImage")));
            TargetProperty.Name = "BackgroundImage";
            TargetProperty.Value = SourceProperty.Value;             
          }
          else
          {
            // ((System.Drawing.Bitmap)(resources.GetObject("Command1.Image")));
            TargetProperty.Name = "Image";
            TargetProperty.Value = SourceProperty.Value;  
          }
          break;

        case "ScrollBars":
          // ScrollBars = System.Windows.Forms.ScrollBars.Both;
          TargetProperty.Name = SourceProperty.Name;
          
          if (Type == "RichTextBox")
          {
            TargetProperty.Value = "System.Windows.Forms.RichTextBoxScrollBars.";
          }
          else
          {
            TargetProperty.Value = "System.Windows.Forms.ScrollBars.";
          }
          switch (SourceProperty.Value)
          {
            default:
            case "0":
              TargetProperty.Value = TargetProperty.Value + "None";
              break;
            case "1":
              TargetProperty.Value = TargetProperty.Value + "Horizontal";
              break;
            case "2":
              TargetProperty.Value = TargetProperty.Value + "Vertical";
              break;                  
            case "3":
              TargetProperty.Value = TargetProperty.Value + "Both";
              break;
          }
          break;

          // SS tab
        case "TabOrientation":
        TargetProperty.Name = "Alignment"; 
        TargetProperty.Value = "System.Windows.Forms.TabAlignment.";
        switch (SourceProperty.Value)
        {
          default:
          case "0":
            TargetProperty.Value = TargetProperty.Value + "Top";
            break;
          case "1":
            TargetProperty.Value = TargetProperty.Value + "Bottom";
            break;
          case "2":
            TargetProperty.Value = TargetProperty.Value + "Left";
            break;                  
          case "3":
            TargetProperty.Value = TargetProperty.Value + "Right";
            break;
          }
          break;

          // begin Listview

          // unsupported properties
        case "_ExtentX":   
        case "_ExtentY":   
        case "_Version":
        case "OLEDropMode":  
          ValidProperty = false;
          break;   

          // this.listView.View = System.Windows.Forms.View.List;
        case "View":
          TargetProperty.Name = SourceProperty.Name;
          TargetProperty.Value = "System.Windows.Forms.View.";
          switch (SourceProperty.Value)
          {
            case "0":
              TargetProperty.Value = TargetProperty.Value + "Details";
              break;
            case "1":
              TargetProperty.Value = TargetProperty.Value + "LargeIcon";
              break;
            case "2":
              TargetProperty.Value = TargetProperty.Value + "SmallIcon";
              break;                  
            case "3":
            default:
              TargetProperty.Value = TargetProperty.Value + "List";
              break;
          }
          break;
                
        case "LabelEdit":  
        case "LabelWrap":
        case "MultiSelect":
        case "HideSelection":  
          TargetProperty.Name = SourceProperty.Name;
          TargetProperty.Value = GetBool(SourceProperty.Value);
          break;              

          // end List view
  

          // VB6 form unsupported properties
        case "MDIChild":  
        case "WhatsThisButton":
        case "NegotiateMenus":  
        case "HelpContextID":  
        case "LinkTopic":   
        case "PaletteMode":  
        case "ClipControls":
        case "LockControls":
        case "FillStyle":
          ValidProperty = false;
          break;   

          // supported properties  

        case "ControlBox":
        case "KeyPreview":
          TargetProperty.Name = SourceProperty.Name;
          TargetProperty.Value = GetBool(SourceProperty.Value);
          break;

        case "ClientHeight":                
          TargetProperty.Name = "ClientSize";
          TargetProperty.Value = "new System.Drawing.Size(" + GetSize("ClientHeight","ClientWidth", SourcePropertyList) + ")";
          break;
              
        case "ClientWidth":
          // nothing, already processed by Height, Left
          ValidProperty = false;
          break;
              
        case "ClientLeft":
        case "ClientTop":              
          ValidProperty = false;
          break;
                
        case "MaxButton":
          TargetProperty.Name = "MaximizeBox";
          TargetProperty.Value = GetBool(SourceProperty.Value);
          break;
        case "MinButton":
          TargetProperty.Name = "MinimizeBox";
          TargetProperty.Value = GetBool(SourceProperty.Value);
          break;
        case "WhatsThisHelp":  
          TargetProperty.Name = "HelpButton";
          TargetProperty.Value = GetBool(SourceProperty.Value);
          break;
        case "ShowInTaskbar":
          TargetProperty.Name = SourceProperty.Name;
          TargetProperty.Value = GetBool(SourceProperty.Value);
          break;
        case"WindowList":
          TargetProperty.Name = "MdiList";
          TargetProperty.Value = GetBool(SourceProperty.Value);
          break;

// this.WindowState = System.Windows.Forms.FormWindowState.Minimized;
          // 0 - normal
          // 1 - minimized
          // 2 - maximized
        case "WindowState":
          TargetProperty.Name = SourceProperty.Name;
          TargetProperty.Value = "System.Windows.Forms.FormWindowState.";
          switch (SourceProperty.Value)
          {
            case "0":
            default:
              TargetProperty.Value = TargetProperty.Value + "Normal";
              break;
            case "1":
              TargetProperty.Value = TargetProperty.Value + "Minimized";
              break;
            case "2":
              TargetProperty.Value = TargetProperty.Value + "Maximized";
              break;                  
          }
          break;

        case "StartUpPosition":
          // 0 - manual
          // 1 - center owner
          // 2 - center screen
          // 3 - windows default
          TargetProperty.Name = "StartPosition";
          TargetProperty.Value = "System.Windows.Forms.FormStartPosition.";
          switch (SourceProperty.Value)
          {
            case "0":
              TargetProperty.Value = TargetProperty.Value + "Manual";
              break;
            case "1":
              TargetProperty.Value = TargetProperty.Value + "CenterParent";
              break;
            case "2":
              TargetProperty.Value = TargetProperty.Value + "CenterScreen";
              break;                  
            case "3":
            default:
              TargetProperty.Value = TargetProperty.Value + "WindowsDefaultLocation";
              break;
          }
          break;                
                
        default:
          TargetProperty.Name = SourceProperty.Name;
          TargetProperty.Value = SourceProperty.Value;
          TargetProperty.Valid = false;
          break;
      }
      return ValidProperty;

    }

    private static void ConvertFont(ControlProperty SourceProperty, ControlProperty TargetProperty)
    {
      string FontName = String.Empty;
      int FontSize = 0;
      int FontCharSet = 0;
      bool FontBold = false;
      bool FontUnderline = false;
      bool FontItalic = false;
      bool FontStrikethrough = false;
      string Temp = String.Empty;
      //      BeginProperty Font 
      //         Name            =   "Arial"
      //         Size            =   8.25
      //         Charset         =   238
      //         Weight          =   400
      //         Underline       =   0   'False
      //         Italic          =   0   'False
      //         Strikethrough   =   0   'False
      //      EndProperty

      foreach (ControlProperty oProperty in SourceProperty.PropertyList) 
      {
        switch(oProperty.Name)
        {
          case "Name":
            FontName = oProperty.Value;
            break;
          case "Size":
            FontSize = GetFontSizeInt(oProperty.Value);
            break;            
          case "Weight":
            //        If tLogFont.lfWeight >= FW_BOLD Then
            //          bFontBold = True
            //        Else
            //          bFontBold = False
            //        End If
            // FW_BOLD = 700
            FontBold = (int.Parse(oProperty.Value) >= 700); 
            break;
          case "Charset":
            FontCharSet = int.Parse(oProperty.Value);
            break;
          case "Underline":
            FontUnderline = (int.Parse(oProperty.Value) != 0); 
            break;
          case "Italic":
            FontItalic = (int.Parse(oProperty.Value) != 0); 
            break;
          case "Strikethrough":
            FontStrikethrough = (int.Parse(oProperty.Value) != 0); 
            break;
        }
      }

      //      this.cmdExit.Font = new System.Drawing.Font("Tahoma", 12F, 
      //        (System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline
      //        | System.Drawing.FontStyle.Strikeout), System.Drawing.GraphicsUnit.Point, 
      //        ((System.Byte)(0)));

      // this.cmdExit.Font = new System.Drawing.Font("Tahoma", 12F, 
      // System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 
      // ((System.Byte)(238)));

      TargetProperty.Name = "Font";
      TargetProperty.Value = "new System.Drawing.Font(" + FontName + ",";
      TargetProperty.Value = TargetProperty.Value + FontSize.ToString() + "F,";

      Temp = String.Empty;
      if (FontBold)
      {
        Temp = "System.Drawing.FontStyle.Bold";
      }
      if (FontItalic)
      {
        if (Temp != String.Empty) {Temp = Temp + " | ";}
        Temp = Temp + "System.Drawing.FontStyle.Italic";
      }
      if (FontUnderline)
      {
        if (Temp != String.Empty) {Temp = Temp + " | ";}
        Temp = Temp + "System.Drawing.FontStyle.Underline";
      }
      if (FontStrikethrough)
      {
        if (Temp != String.Empty) {Temp = Temp + " | ";}
        Temp = Temp + "System.Drawing.FontStyle.Strikeout";
      }
      if (Temp == String.Empty)
      {
        TargetProperty.Value = TargetProperty.Value + " System.Drawing.FontStyle.Regular,";
      }
      else
      {
        TargetProperty.Value = TargetProperty.Value + " ( " + Temp + " ),";
      }
      TargetProperty.Value = TargetProperty.Value + " System.Drawing.GraphicsUnit.Point, ";
      TargetProperty.Value = TargetProperty.Value + "((System.Byte)(" + FontCharSet.ToString() + ")));";
    }

    private static int GetFontSizeInt(string Value)
    {
      int Position = 0;

      Position = Value.IndexOf(",",0);
      if (Position > -1)
      {
        return int.Parse(Value.Substring(0, Position));
      }

      Position = Value.IndexOf(".",0);
      if (Position > 0)
      {
        return int.Parse(Value.Substring(0, Position));
      }
      return int.Parse(Value);
    }

    private static string GetColor(string Value)
    {
      System.Drawing.Color Color = System.Drawing.SystemColors.Control;
      string ColorValue;
      
      ColorValue = "0x" + Value.Substring(2, Value.Length - 3); 
      Color = System.Drawing.ColorTranslator.FromWin32(System.Convert.ToInt32(ColorValue, 16));

      if (!Color.IsSystemColor) 
      {
        if (Color.IsNamedColor) 
        {
          // System.Drawing.Color.Yellow;
          return "System.Drawing.Color." + Color.Name;
        }
        else
        {
          return "System.Drawing.Color.FromArgb(" + Color.ToArgb() + ")";
        }
      }
      else
      {
        return "System.Drawing.SystemColors." + Color.Name;
      }
    }

    private static string GetBool(string Value)
    {
      if (int.Parse(Value) == 0)
      {
        return "false";
      }
      else
      {
        return "true";
      }
    }

    private static string GetSize(string Height, string Width, ArrayList PropertyList)
    {
      int HeightValue = 0;
      int WidthValue = 0;
    
      // each property  
      foreach ( ControlProperty oProperty in PropertyList )
      {
        if (oProperty.Name == Height)
        {
          HeightValue = int.Parse(oProperty.Value) / 15;
        }
        if (oProperty.Name == Width)
        {
          WidthValue = int.Parse(oProperty.Value) / 15;
        }        
      }
      // 0, 120
      return WidthValue.ToString() + ", " + HeightValue.ToString();        
    }
    
    private static string GetLocation( ArrayList PropertyList)
    {
      int Left = 0;
      int Top = 0;
    
      // each property  
      foreach ( ControlProperty oProperty in PropertyList )
      {
        if (oProperty.Name == "Left")
        {
          Left = int.Parse(oProperty.Value);
          if (Left < 0) 
          {
            Left = 75000 + Left;
          }
          Left = Left / 15;
        }
        if (oProperty.Name == "Top")
        {
          Top = int.Parse(oProperty.Value) / 15;
        }        
      }
      // 616, 520
      return Left.ToString() + ", " + Top.ToString();        
    }    

    public static void GetFRXImage(string ImageFile, int ImageOffset, out byte[] ImageString)
    {
      byte[] Header;
      int BytesToRead = 0;

      // open file
      FileStream Stream = new FileStream( ImageFile, FileMode.Open, FileAccess.Read );
      BinaryReader Reader = new BinaryReader( Stream );
      // Start from offset
      Reader.BaseStream.Seek( ImageOffset, SeekOrigin.Begin);
      // Get the four byte header
      Header = new byte[4];
      Header = Reader.ReadBytes(4);
      // Convert This Header Into The Number Of Bytes
      // To Read For This Image
      BytesToRead = Header[0];
      BytesToRead = BytesToRead + ( Header[1] * 0x100 );
      BytesToRead = BytesToRead + ( Header[2] * 0x10000 );        
      BytesToRead = BytesToRead + ( Header[3] * 0x1000000 );
      // Get image information
      ImageString = new byte[BytesToRead];
      ImageString = Reader.ReadBytes(BytesToRead);


//      Stream = new FileStream( @"C:\temp\test\Ba.bmp", FileMode.CreateNew, FileAccess.Write );
//      BinaryWriter Writer = new BinaryWriter( Stream );
//      Writer.Write(ImageString, 8, ImageString.GetLength(0) - 8);           
//      Stream.Close();
//      Writer.Close(); 
     
    //  FileStream inFile = new FileStream(@"C:\WINdows\Blue Lace 16.bmp", FileMode.Open, FileAccess.Read);
    //  ReturnImage = Image.FromStream(inFile, false);

      Stream.Close();
      Reader.Close();

    }
	}
}
