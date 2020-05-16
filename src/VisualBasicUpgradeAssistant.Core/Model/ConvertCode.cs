using System;
using System.Collections;
using System.Drawing;
using System.IO;
using System.Resources;
using System.Text;
 
namespace VB2C
{
  public enum VB_FILE_TYPE
  {
    VB_FILE_UNKNOWN = 0,
    VB_FILE_FORM = 1,
    VB_FILE_MODULE = 2,
    VB_FILE_CLASS = 3
  };

	/// <summary>
	/// Summary description for Convert.
	/// </summary>
  public class ConvertCode
  {
    private VB_FILE_TYPE mFileType;
    private string msActionResult;
    private Module mSourceModule;
    private Module mTargetModule;
    private ArrayList moOwnerStock;
    private string mOutSourceCode;

    private const string FORM_FIRST_LINE = "VERSION 5.00";
    private const string MODULE_FIRST_LINE = "ATTRIBUTE";
    private const string CLASS_FIRST_LINE = "1.0 CLASS";

    private const string Indent2 = "  ";
    private const string Indent4 = "    ";
    private const string Indent6 = "      ";
    
    public ConvertCode()
    {
      moOwnerStock = new ArrayList();
    }

    public string ActionResult 
    {
      get { return msActionResult; }
    }
    
    public string OutSourceCode
    {
      get { return mOutSourceCode; }
    }
    
    public bool ParseFile(string FileName, string OutPath)
    {
      string Line = String.Empty;
      string Temp = String.Empty;
      string Extension = String.Empty;
      string Version = String.Empty;
      bool Result = false;
      int Position = 0;

      // try recognize source code type depend by file extension
      Extension = FileName.Substring(FileName.Length - 3, 3);
      switch (Extension.ToUpper())     
      {         
        case "FRM":   
          mFileType = VB_FILE_TYPE.VB_FILE_FORM;
          break;
        case "BAS":   
          mFileType = VB_FILE_TYPE.VB_FILE_MODULE;
          break;                  
        case "CLS":            
          mFileType = VB_FILE_TYPE.VB_FILE_CLASS;
          break;           
        default:            
          mFileType = VB_FILE_TYPE.VB_FILE_UNKNOWN;            
          break;      
      }

      // open file
      FileStream Stream = new FileStream( FileName, FileMode.Open, FileAccess.Read );
      StreamReader Reader = new StreamReader( Stream );
      // Start from begin.
      Reader.BaseStream.Seek(0, SeekOrigin.Begin);
      Line = Reader.ReadLine();
      // verify type of file based on first line - form, module, class

        // get first word from first line
        Position = 0;
        Temp = GetWord(Line, ref Position);
        switch (Temp.ToUpper())     
        {   
          // module first line
          // 'Attribute VB_Name = "ModuleName"'
          case MODULE_FIRST_LINE:   
            mFileType = VB_FILE_TYPE.VB_FILE_MODULE;
            break; 
            // form or class first line
            // 'VERSION 5.00' or 'VERSION 1.0 CLASS' 
          case "VERSION":  
            Position ++;
            Version = GetWord(Line, ref Position);
            
//            Debug.WriteLine (Line + " " + Version);
          
            if (Line.IndexOf(CLASS_FIRST_LINE, 0) > -1 )
            {
              mFileType = VB_FILE_TYPE.VB_FILE_CLASS;
            }
            else
            {
              mFileType = VB_FILE_TYPE.VB_FILE_FORM;
            }
            break;
          default:            
            mFileType = VB_FILE_TYPE.VB_FILE_UNKNOWN;            
            break;      
        }
        // if file is still unknown
        if (mFileType == VB_FILE_TYPE.VB_FILE_UNKNOWN)
        {
          msActionResult = "Unknown file type";
          return false;
        }

      mSourceModule = new Module();
      mSourceModule.Version = Version;
      mSourceModule.FileName = FileName;

      // now parse specifics of each type
      switch (Extension.ToUpper())     
      {         
        case "FRM":   
          mSourceModule.Type = "form";
          Result = ParseForm(Reader);
          break;
        case "BAS":   
          mSourceModule.Type = "module";
          Result = ParseModule(Reader);
          break;                  
        case "CLS":          
          mSourceModule.Type = "class";
          Result = ParseClass(Reader);
          break;           
      }
      // parse remain - variables, functions, procedures
      Result = ParseProcedures(Reader);

      Stream.Close();
      Reader.Close();

      // generate output file
      mOutSourceCode = GetOutSourceCode(OutPath);

      // save result
      string OutFileName = OutPath + mTargetModule.FileName;
      Stream = new FileStream(OutFileName, FileMode.OpenOrCreate);
      StreamWriter Writer = new StreamWriter(Stream);
      Writer.Write(mOutSourceCode);
      Writer.Close();
      // generate resx file if source form contain any images

      if ((mTargetModule.ImagesUsed))
      {
        WriteResX(mTargetModule.ImageList, OutPath, mTargetModule.Name); 
      }

      return Result;
    }

    private bool ParseForm(StreamReader oReader)
    {
      bool bProcess = false;
      bool bFinish = false;
      bool bEnd = true;
      string sLine = null;
      string sName = null;
      string sOwner = null;
      string sTemp = null;
      int iTemp = 0;
      int iComment = 0;
      string sType = null;
      int iPosition = 0;
      int iLevel = 0;
      Control oControl = null;
      ControlProperty oNestedProperty = null;
      bool bNestedProperty = false;

      // parse only visual part of form
      while (! bFinish) // ( ( bFinish || (oReader.Peek() > -1)) )
      {        
        sLine = oReader.ReadLine();
        sLine = sLine.Trim();         
        iPosition = 0;
        // get first word in line
        sTemp = GetWord(sLine, ref iPosition);
        switch(sTemp)
        {
          case "Begin":   
            bProcess = true;
            // new level
            iLevel ++;   
            // next word - control type
            iPosition ++;
            sType = GetWord(sLine, ref iPosition);
            // next word - control name
            iPosition ++;
            sName = GetWord(sLine, ref iPosition);          
            // detected missing end -> it indicate that next control is container
            if (! bEnd)
            {
              // add container control to colection
              if (!(oControl == null))
              {                
                oControl.Container = true;
                mSourceModule.ControlAdd(oControl);
              }
              // save name of previous control as owner for current and next controls
              moOwnerStock.Add(sOwner);
            }
            bEnd = false;
            
            switch(sType)
            {
              case "Form":            // VERSION 2.00 - VB3
              case "VB.Form":
              case "VB.MDIForm":
                mSourceModule.Name = sName;
                // first owner
                // save control name for possible next controls as owner
                sOwner = sName;
                break;
              default:
                // new control
                oControl = new Control();
                oControl.Name = sName;
                oControl.Type = sType;
                // save control name for possible next controls as owner
                sOwner = sName;
                // set current container name
                oControl.Owner = (string) moOwnerStock[moOwnerStock.Count - 1];
                break;
            }
            break;
          
          case "End":  
            // double end - we leaving some container
            if (bEnd)
              {
                // remove last item from stock
                moOwnerStock.Remove((string) moOwnerStock[moOwnerStock.Count - 1]);              
              }     
            else
              {
                // level 1 is form and all higher levels are controls
                if (iLevel > 1)
                {
                  // add control to colection
                  mSourceModule.ControlAdd(oControl);
                }                
              }
            // form or control end detected
            bEnd = true;
            // back to previous level
            iLevel --;

            break;   

          case "Object":  
            // used controls in form
            break; 

          case "BeginProperty":
            bNestedProperty = true;
            
            oNestedProperty = new ControlProperty();
            // next word - nested property name
            iPosition ++;
            sName = GetWord(sLine, ref iPosition); 
            oNestedProperty.Name = sName;
//            Debug.WriteLine(sName); 
            break;

          case "EndProperty":
            bNestedProperty = false;
            // add property to control or form
            if (iLevel == 1)
            {
              // add property to form
              mSourceModule.FormPropertyAdd(oNestedProperty);
            }
            else
            {
              // to controls
              oControl.PropertyAdd(oNestedProperty);          
            }
            break;
          
          default:         
            // parse property
            ControlProperty oProperty = new ControlProperty();

            iTemp = sLine.IndexOf("=");
            if (iTemp > -1)
            {
              oProperty.Name = sLine.Substring(0, iTemp - 1).Trim();
              iComment = sLine.IndexOf("'", iTemp);
              if ( iComment > -1 )
              {
                oProperty.Value = sLine.Substring( iTemp + 1, iComment - iTemp - 1).Trim();
                oProperty.Comment = sLine.Substring( iComment + 1, sLine.Length - iComment - 1).Trim();
              }
              else
              {
                oProperty.Value = sLine.Substring( iTemp + 1, sLine.Length - iTemp - 1).Trim();
              }

              if (bNestedProperty)
              {
                oNestedProperty.PropertyList.Add(oProperty);
              }
              else
                {
                  // depend by level insert property to form or control
                  if (iLevel > 1)
                  {
                    // add property to control
                    oControl.PropertyAdd(oProperty);
                  }
                  else
                  {
                    // add property to form
                    mSourceModule.FormPropertyAdd(oProperty);
                  }
                }  
            }
            break;  
        }

        if ((iLevel == 0) && bProcess)
        {
          // visual part of form is finish
          bFinish = true;
        }

      }
      return true;
    }

    private bool ParseModule(StreamReader Reader)
    {
      string Line = String.Empty;
      int Position = 0;

      // name of module
// Attribute VB_Name = "ModuleName"

      // Start from begin again
      Reader.DiscardBufferedData();
      Reader.BaseStream.Seek(0, SeekOrigin.Begin);
      // search for module name
      while ( Reader.Peek() > -1 )
      {
        Line = Reader.ReadLine();
        Position = Line.IndexOf('"');
        mSourceModule.Name = Line.Substring(Position + 1, Line.Length - Position - 2); 
        return true;
      }
      return false;
    }

    private bool ParseClass(StreamReader Reader)
    {
      int Position = 0;
      string Line = null;
      string TempString = String.Empty; 
      
//VERSION 1.0 CLASS
//BEGIN
//  MultiUse = -1  'True
//END
//Attribute VB_Name = "CList"
//Attribute VB_GlobalNameSpace = False
//Attribute VB_Creatable = True
//Attribute VB_PredeclaredId = False
//Attribute VB_Exposed = True

      // Start from begin again
      Reader.DiscardBufferedData();
      Reader.BaseStream.Seek(0, SeekOrigin.Begin);
      while ( Reader.Peek() > -1 )
      {
        Position = 0;
        // verify type of file based on first line
        // form, module, class
        Line = Reader.ReadLine();
        // next word - control type
        TempString = GetWord(Line, ref Position);
        if (TempString == "Attribute")
        {
          Position ++;
          TempString = GetWord(Line, ref Position);
          switch(TempString)
          {
            case "VB_Name":
              Position ++;
              TempString = GetWord(Line, ref Position);
              Position ++;
              mSourceModule.Name = GetWord(Line, ref Position);
              break;

            case "VB_Exposed":
              return true;
              //break;
          }
        }
      }
      return false;
    }

    private bool ParseProcedures(StreamReader Reader)
    {
      string Line = null;
      string TempString = null;
      string sComments = null;
      string sScope = null;
      
      int iPosition = 0;
      //bool bProcess = false;
      
      bool bEnum = false;
      bool bVariable = false;
      bool bProperty = false;
      bool bProcedure = false;
      bool bEnd = false;
      
      Variable oVariable = null;
      Property oProperty = null;
      Procedure oProcedure = null;
      Enum oEnum = null;
      EnumItem oEnumItem = null;
      
      while ( Reader.Peek() > -1 )
      {
        Line = Reader.ReadLine();
        //Line = Line.Trim();  
        
        iPosition = 0;

        if (Line != null && Line != String.Empty)
        {
          // check if next line is same command, join it together ?
          while ( Line.Substring(Line.Length - 1, 1) == "_" )
          {
            Line = Line + Reader.ReadLine();
          }
        }
        // : is command delimiter
        
        
      //  Debug.WriteLine(Line); 
        
        // get first word in line
        TempString = GetWord(Line, ref iPosition);
        switch(TempString)
        {
// ignore this section

//Attribute VB_Name = "frmAttachement"
// ...
//Option Explicit          
          case "Attribute":   
          case "Option":  
            break;

// comments          
          case "'":
            sComments = sComments + Line + "\r\n";
            break;
          
// next can be declaration of variables   
     
//Private mlParentID As Long
//Private mlOwnerType As ENUM_FORM_TYPE
//Private moAttachement As Attachement
          
          case "Public":
          case "Private":
            // save it for later use
            sScope = TempString.ToLower();
            // read next word
            // next word - control type
            iPosition ++;
            TempString = GetWord(Line, ref iPosition);

            switch(TempString)
            {
              // functions or procedures
              case "Sub":
              case "Function":

                oProcedure = new Procedure();
                oProcedure.Comment = sComments;
                sComments = String.Empty; 
                ParseProcedureName(oProcedure, Line);

                bProcedure = true;
                break;

              case "Enum":  
                oEnum = new Enum();
                oEnum.Scope = sScope;
                // next word is enum name
                iPosition ++;
                oEnum.Name = GetWord(Line, ref iPosition);
                bEnum = true;
                break;

              case "Property":
                oProperty = new Property();
                oProperty.Comment = sComments;
                sComments = String.Empty; 
                ParsePropertyName(oProperty, Line);
                bProperty = true;

                break;
              default:
                // variable declaration
                oVariable = new Variable();
                ParseVariableDeclaration(oVariable, Line);
                bVariable = true;          
                break;              
            }
            
            break;
          
          case "Dim":  
            // variable declaration
            oVariable = new Variable();
            oVariable.Comment = sComments;
            sComments = String.Empty; 
            ParseVariableDeclaration(oVariable, Line);
            bVariable = true;          
            break;         
          
          // functions or procedures
          case "Sub":
          case "Function":
            
            break;
          
          case "End":
            bEnd = true;
            break;


          default:  
            if (bEnum)
            {
              // first word is name, second =, thirt value if is preset
              oEnumItem = new EnumItem();
              oEnumItem.Comment = sComments;
              sComments = String.Empty; 
              ParseEnumItem(oEnumItem, Line);
              // add item
              oEnum.ItemList.Add(oEnumItem);
            }
            if (bProperty) 
            {
              // add line of property
              oProperty.LineList.Add(Line); 
            }
            if (bProcedure)
            {
              oProcedure.LineList.Add(Line); 
            }
          break;
          

// events
//Private Sub cmdCancel_Click()
//  mbEdit = False
//  If mbNew Then
//    Unload Me
//  Else
//    ShowCurRec
//    SetControls False
//  End If
//End Sub
//
//Private Sub cmdClose_Click()
//  Unload Me
//End Sub


        }

        // if something end
        if (bEnd)
        {
          // 
          if (bEnum)
          {
            mSourceModule.EnumList.Add(oEnum); 
            bEnum = false;
          }
          if (bProperty) 
          {
            mSourceModule.PropertyAdd(oProperty);
            bProperty = false;
          }
          if (bProcedure)
          {
            mSourceModule.ProcedureAdd(oProcedure);
            bProcedure = false;
          }
          bEnd = false;
        }
        else
        {
          if (bVariable)
          {
            mSourceModule.VariableAdd(oVariable);
          }          
        }

        bVariable = false;
      }
      
      return true;
    }

    //Public Enum ENUM_BUG_LEVEL
    //  BUG_LEVEL_PROJECT = 1
    //  BUG_LEVEL_VERSION = 2
    //End Enum
    private void ParseEnumItem(EnumItem oEnumItem, string Line)
    {
      string TempString = String.Empty; 
      int iPosition = 0;

      Line = Line.Trim(); 
      // first word is ame
      oEnumItem.Name = GetWord(Line, ref iPosition);
      iPosition ++;
      // next word =
      TempString = GetWord(Line, ref iPosition);
      iPosition ++;
      // optional
      oEnumItem.Value = GetWord(Line, ref iPosition);
    }


//Private mlID As Long

    private void ParseVariableDeclaration(Variable oVariable, string Line)
    {
      string TempString = String.Empty; 
      int iPosition = 0;
      bool Status = false;

      // next word - control type
      TempString = GetWord(Line, ref iPosition);
      switch(TempString)
      {
        case "Dim":        
        case "Private":
          oVariable.Scope = "private";
          break;
        case "Public":
          oVariable.Scope = "public";
          break;
        default:
          oVariable.Scope = "private";
          // variable name
          oVariable.Name = TempString;  
          Status = true;
          break;
      }

      // variable name
      if (!Status) 
      {
        iPosition ++;
        TempString = GetWord(Line, ref iPosition);
        oVariable.Name = TempString;  
      }
      // As 
      iPosition ++;
      TempString = GetWord(Line, ref iPosition);
      // variable type
      iPosition ++;
      TempString = GetWord(Line, ref iPosition);
      oVariable.Type = TempString;

    }

// properties Let, Get, Set
//Public Property Let ParentID(ByVal lValue As Long)
//  mlParentID = lValue
//End Property
//
//Public Property Get FormType() As ENUM_FORM_TYPE
//  FormType = FORM_ATTACHEMENT
//End Property

    private void ParsePropertyName( Property oProperty, string Line)
    {
      string TempString = String.Empty; 
      int iPosition = 0;
      int Start = 0;
      bool Status = false;

      // next word - control type
      TempString = GetWord(Line, ref iPosition);
      switch(TempString)
      {   
        case "Private":
          oProperty.Scope = "private";
          break;
        case "Public":
          oProperty.Scope = "public";
          break;
        default:
          oProperty.Scope = "private";
          Status = true;
          break;
      }

      if (!Status) 
      {
        // property
        iPosition ++;
        TempString = GetWord(Line, ref iPosition);
      }
      
      // direction Let,Get, Set
      iPosition ++;
      TempString = GetWord(Line, ref iPosition);
      oProperty.Direction = TempString;
      
//Public Property Let ParentID(ByVal lValue As Long)

      // name       
      Start = iPosition;
      iPosition = Line.IndexOf("(", Start + 1);
      oProperty.Name = Line.Substring(Start ,iPosition - Start);
      
      // + possible parameters
      iPosition ++;
      Start = iPosition;      
      iPosition = Line.IndexOf(")", Start);
            
      if ((iPosition - Start) > 0)
      {        
        TempString = Line.Substring( Start , iPosition - Start);       
        ArrayList ParameterList = new ArrayList();
        // process parametres
        ParseParametries(ParameterList, TempString );
        oProperty.ParameterList = ParameterList;        
      }
      
      // As 
      iPosition ++;
      iPosition ++;
      TempString = GetWord(Line, ref iPosition);
      
      // type
      iPosition ++;
      TempString = GetWord(Line, ref iPosition);
      oProperty.Type = TempString;
      
    }

// ByVal lValue As Long, ByVal sValue As string

    private void ParseParametries( ArrayList ParametreList, string Line)
    {
      bool bFinish = false;
      int Position = 0;
      int Start = 0;
      bool Status = false;
      string TempString = String.Empty; 
      
      // parameters delimited by comma
      while (! bFinish)
      {        
        Parameter oParameter = new Parameter();
                
        // next word - control type
        TempString = GetWord(Line, ref Position);
        switch(TempString)
        {
          case "Optional":
            
            break;
          case "ByVal":
          case "ByRef":
            oParameter.Pass = TempString;
            break;
          // missing is byref
          default:
            oParameter.Pass = "ByRef";
            // variable name
            oParameter.Name = TempString;  
            Status = true;
            break;
        }
  
        // variable name
        if (!Status) 
        {
          Position ++;
          TempString = GetWord(Line, ref Position);
          oParameter.Name = TempString;  
        }
        // As 
        Position ++;
        TempString = GetWord(Line, ref Position);
        // parameter type
        Position ++;
        TempString = GetWord(Line, ref Position);
        oParameter.Type = TempString;

        ParametreList.Add(oParameter); 
        
        // next parameter
        Position ++;
        Start = Position;      
        Position = Line.IndexOf(",", Start);

        if (Position == -1) 
        {
          // end
          bFinish = true;
        }

      }
    }
    
    private void ParseProcedureName( Procedure oProcedure, string Line)
    {
      string TempString = String.Empty; 
      int iPosition = 0;
      int Start = 0;
      bool Status = false;

      //Private Sub cmdOk_Click()
      //private void cmdShow_Click(object sender, System.EventArgs e)
      
      //Private Sub Form_Load()
      //private void frmConvert_Load(object sender, System.EventArgs e)

      //Public Function Rozbor_DefaultFields(ByVal MKf As String) As String
      //public static bool ParseProcedures( Module SourceModule, Module TargetModule )

      TempString = GetWord(Line, ref iPosition);
      switch(TempString)
      {   
        case "Private":
          oProcedure.Scope = "private";
          Status = true;
          break;
        case "Public":
          oProcedure.Scope = "public";
          Status = true;
          break;
        default:
          oProcedure.Scope = "private";
          Status = true;
          break;
      }

      if (Status) 
      {
        // property
        iPosition ++;
        TempString = GetWord(Line, ref iPosition);
      }

      // procedure type
      switch(TempString)
      {   
        case "Sub":
          oProcedure.Type = PROCEDURE_TYPE.PROCEDURE_SUB;
          break;
        case "Function":
          oProcedure.Type = PROCEDURE_TYPE.PROCEDURE_FUNCTION;
          break;
        case "Event":
          oProcedure.Type = PROCEDURE_TYPE.PROCEDURE_EVENT;
          break;
      }

      // next is name
      iPosition ++;
      Start = iPosition;      
      iPosition = Line.IndexOf("(", Start);
      oProcedure.Name = Line.Substring(Start, iPosition - Start );  

      // next possible parameters
      iPosition ++;
      Start = iPosition;      
      iPosition = Line.IndexOf(")", Start);
      
      if ((iPosition - Start) > 0)
      {        
        TempString = Line.Substring( Start , iPosition - Start);       
        ArrayList ParameterList = new ArrayList();
        // process parametres
        ParseParametries(ParameterList, TempString );
        oProcedure.ParameterList = ParameterList;        
      }  

      // and return type of function
      if ( oProcedure.Type == PROCEDURE_TYPE.PROCEDURE_FUNCTION )
      {
        // as
        iPosition ++;
        TempString = GetWord(Line, ref iPosition);
        // function return type
        iPosition ++;
        oProcedure.ReturnType = GetWord(Line, ref iPosition);
      }

    }
    

    // generate result file
    // OutPath for pictures
    private string GetOutSourceCode(string OutPath)
    {
      StringBuilder oResult = new StringBuilder();

      string Temp = String.Empty;
      
      // convert source to target    
      mTargetModule = new Module();
      Tools.ParseModule(mSourceModule, mTargetModule); 
      
      // ********************************************************
      // common class
      // ********************************************************
      oResult.Append("using System;\r\n");

      // ********************************************************
      // only form class
      // ********************************************************
      if (mTargetModule.Type == "form")
      {	  
        oResult.Append("using System.Drawing;\r\n");
        oResult.Append("using System.Collections;\r\n");
        oResult.Append("using System.ComponentModel;\r\n");
        oResult.Append("using System.Windows.Forms;\r\n");
      }


      oResult.Append("\r\n");
      oResult.Append("namespace ProjectName\r\n");
      // start namepsace region
      oResult.Append("{\r\n");
      oResult.Append(Indent2 + "/// <summary>\r\n");
      oResult.Append(Indent2 + "/// Summary description for " + mSourceModule.Name + ".\r\n");
      oResult.Append(Indent2 + "/// </summary>\r\n");


      switch (mTargetModule.Type)
      {
        case "form":
          oResult.Append(Indent2 + "public class " + mSourceModule.Name + " : System.Windows.Forms.Form\r\n");
          break;
        case "module":
          oResult.Append(Indent2 + "sealed class " + mSourceModule.Name + "\r\n");
          // all procedures must be static
          break;
        case "class":
          oResult.Append(Indent2 + "public class " + mSourceModule.Name + "\r\n");
          break;
      }
      // start class region
      oResult.Append(Indent2 + "{\r\n");

      // ********************************************************
      // only form class
      // ********************************************************

      if (mTargetModule.Type == "form")
      {	  
        // list of controls
        foreach ( Control oControl in mTargetModule.ControlList )
        {
          if (!oControl.Valid)
          {
            oResult.Append("//");
          }
          oResult.Append(Indent2 + " private System.Windows.Forms." + oControl.Type + " " + oControl.Name + ";\r\n");        
        }
            
        oResult.Append(Indent4 + "/// <summary>\r\n");
        oResult.Append(Indent4 + "/// Required designer variable.\r\n");
        oResult.Append(Indent4 + "/// </summary>\r\n");
        oResult.Append(Indent4 + "private System.ComponentModel.Container components = null;\r\n");
        oResult.Append("\r\n");
        oResult.Append(Indent4 + "public " + mSourceModule.Name + "()\r\n");
        oResult.Append(Indent4 + "{\r\n");
        oResult.Append(Indent6 + "// Required for Windows Form Designer support\r\n");
        oResult.Append(Indent6 + "InitializeComponent();\r\n");
        oResult.Append("\r\n");
        oResult.Append(Indent6 + "// TODO: Add any constructor code after InitializeComponent call\r\n");
        oResult.Append(Indent4 + "}\r\n");
      
        oResult.Append(Indent4 + "/// <summary>\r\n");
        oResult.Append(Indent4 + "/// Clean up any resources being used.\r\n");
        oResult.Append(Indent4 + "/// </summary>\r\n");
        oResult.Append(Indent4 + "protected override void Dispose( bool disposing )\r\n");
        oResult.Append(Indent4 + "{\r\n");
        oResult.Append(Indent6 + "if( disposing )\r\n");
        oResult.Append(Indent6 + "{\r\n");
        oResult.Append(Indent6 + "  if (components != null)\r\n");
        oResult.Append(Indent6 + "  {\r\n");
        oResult.Append(Indent6 + "    components.Dispose();\r\n");
        oResult.Append(Indent6 + "  }\r\n");
        oResult.Append(Indent6 + "}\r\n");
        oResult.Append(Indent6 + "base.Dispose( disposing );\r\n");
        oResult.Append(Indent4 + "}\r\n");

        oResult.Append(Indent4 + "#region Windows Form Designer generated code\r\n");
        oResult.Append(Indent4 + "/// <summary>\r\n");
        oResult.Append(Indent4 + "/// Required method for Designer support - do not modify\r\n");
        oResult.Append(Indent4 + "/// the contents of this method with the code editor.\r\n");
        oResult.Append(Indent4 + "/// </summary>\r\n");
        oResult.Append(Indent4 + "private void InitializeComponent()\r\n");
        oResult.Append(Indent4 + "{\r\n");
        
		    // if form contain images
        if (mTargetModule.ImagesUsed) 
        {
          // System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(Form1));
          oResult.Append(Indent6 + "System.Resources.ResourceManager resources = " +
            "new System.Resources.ResourceManager(typeof(" + mTargetModule.Name + "));\r\n");
        }

        foreach ( Control oControl in mTargetModule.ControlList )
        {
          if (!oControl.Valid)
          {
            oResult.Append("//");
          }
          oResult.Append(Indent6 + "this." + oControl.Name 
            + " = new System.Windows.Forms." + oControl.Type 
            + "();\r\n" );
      
        }
      
        // SuspendLayout part
        oResult.Append(Indent6 + "this.SuspendLayout();\r\n" );
        // this.Frame1.ResumeLayout(false);
        // resume layout for each container
        foreach ( Control oControl in mTargetModule.ControlList )
        {
          // check if control is container
          // !! for menu controls
          if ((oControl.Container) && !(oControl.Type == "MenuItem") && !(oControl.Type == "MainMenu"))
          {
            if (!oControl.Valid)
            {
              oResult.Append("//");
            }
            oResult.Append(Indent6 + "this." + oControl.Name + ".SuspendLayout();\r\n");
          }
        }
      
        // each controls and his property		  
        foreach ( Control oControl in mTargetModule.ControlList )
        {          
          oResult.Append(Indent6 + "//\r\n"); 
          oResult.Append(Indent6 + "// " + oControl.Name + "\r\n");
          oResult.Append(Indent6 + "//\r\n");
        
          // unsupported control
          if (!oControl.Valid) { oResult.Append("/*"); } 

          // ImageList, Timer, Menu has't name property
          if ((oControl.Type != "ImageList") && (oControl.Type != "Timer") 
            && (oControl.Type != "MenuItem") && (oControl.Type != "MainMenu"))
          {
            // control name
            oResult.Append(Indent6 + "this." + oControl.Name + ".Name = " 
              + (char)34 + oControl.Name + (char)34 + ";\r\n");
          } 

          // write properties
          foreach ( ControlProperty oProperty in oControl.PropertyList )
          {         
            GetPropertyRow(oResult, oControl.Type, oControl.Name, oProperty, OutPath);            
          }  

          // if control is container for other controls
          Temp = String.Empty; 
          foreach ( Control oControl1 in mTargetModule.ControlList )
          {
            // all controls ownered by current control
            if ((oControl1.Owner == oControl.Name) && ( !oControl1.InvisibleAtRuntime ))
            {         
              Temp = Temp + Indent6 + Indent6 + "this." + oControl1.Name + ",\r\n";
            }
          }   
          if (Temp != String.Empty)
          {
            // exception for menu controls
            if (oControl.Type == "MainMenu" || oControl.Type == "MenuItem")
            {
              // this.mainMenu1.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] 
              oResult.Append(Indent6 + "this." + oControl.Name 
                + ".MenuItems.AddRange(new System.Windows.Forms.MenuItem[]\r\n");
            }
            else
            {
              // this. + oControl.Name + .Controls.AddRange(new System.Windows.Forms.Control[]
              oResult.Append(Indent6 + "this." + oControl.Name 
                + ".Controls.AddRange(new System.Windows.Forms.Control[]\r\n");
            }
            
            oResult.Append(Indent6 + "{\r\n");
            oResult.Append(Temp);
            // remove last comma, keep CRLF
            oResult.Remove( oResult.Length - 3 ,1);
            // close addrange part  
            oResult.Append(Indent6 + "});\r\n");        
          }
          // unsupported control
          if (!oControl.Valid) { oResult.Append("*/"); } 
        }

        oResult.Append(Indent6 + "//\r\n"); 
        oResult.Append(Indent6 + "// " + mSourceModule.Name + "\r\n"); 
        oResult.Append(Indent6 + "//\r\n"); 
        oResult.Append(Indent6 + "this.Controls.AddRange(new System.Windows.Forms.Control[]\r\n");
        oResult.Append(Indent6 + "{\r\n");
  

        // add control range to form
        foreach ( Control oControl in mTargetModule.ControlList )
        {
          if (!oControl.Valid)
          {
            oResult.Append("//");
          }        
          // all controls ownered by main form
          if ((oControl.Owner == mSourceModule.Name) && ( !oControl.InvisibleAtRuntime ))
          {
            oResult.Append(Indent6 + "      this." + oControl.Name + ",\r\n");
          }
        }   

        // remove last comma, keep CRLF
        oResult.Remove( oResult.Length - 3 ,1);
        // close addrange part  
        oResult.Append(Indent6 + "});\r\n");
      
        // form name
        oResult.Append(Indent6 + "this.Name = " + (char)34 + mTargetModule.Name + (char)34 + ";\r\n" );
        // exception for menu
        // this.Menu = this.mainMenu1;
        if (mTargetModule.MenuUsed)
        {
          foreach ( Control oControl in mTargetModule.ControlList )
          {
            if (oControl.Type == "MainMenu")
            {
              oResult.Append(Indent6 + "      this.Menu = " + oControl.Name + ";\r\n");
            }
          }
        }
        // form properties
        foreach ( ControlProperty oProperty in mTargetModule.FormPropertyList )
        {
          if (!oProperty.Valid)
          {
            oResult.Append("//");
          }
          GetPropertyRow( oResult, mTargetModule.Type, "", oProperty, OutPath );
        } 

        // this.CancelButton = this.cmdExit;
          
          
        // this.Frame1.ResumeLayout(false);
        // resume layout for each container
        foreach ( Control oControl in mTargetModule.ControlList )
        {
          // check if control is container
          if ((oControl.Container) && !(oControl.Type == "MenuItem") && !(oControl.Type == "MainMenu"))
          {
            if (!oControl.Valid)
            {
              oResult.Append("//");
            }
            oResult.Append(Indent6 + "this." + oControl.Name + ".ResumeLayout(false);\r\n");
          }
        }
        // form
        oResult.Append(Indent6 + "this.ResumeLayout(false);\r\n");

        oResult.Append(Indent4 + "}\r\n");
        oResult.Append(Indent4 + "#endregion\r\n");
      } // if (mTargetModule.Type = "form")


      // ********************************************************
      // enums
      // ********************************************************

      if (mTargetModule.EnumList.Count > 0)
      {
        oResult.Append("\r\n");
        foreach ( Enum oEnum in mTargetModule.EnumList )
        {
          // public enum VB_FILE_TYPE
          oResult.Append(Indent4 + oEnum.Scope + " enum " + oEnum.Name + "\r\n");
          oResult.Append(Indent4 + "{\r\n");

          foreach ( EnumItem oEnumItem in oEnum.ItemList  )
          {
            // name
            oResult.Append(Indent6 + oEnumItem.Name);

            if (oEnumItem.Value != String.Empty )
            {
              oResult.Append(" = " + oEnumItem.Value);
            }
            // enum items delimiter
            oResult.Append(",\r\n");

          }
          // remove last comma, keep CRLF
          oResult.Remove( oResult.Length - 3 ,1);
          // end enum
          oResult.Append(Indent4 + "};\r\n");
        }
      }

      // ********************************************************
      //  variables for al module types
      // ********************************************************

      if (mTargetModule.VariableList.Count > 0)
      {
        oResult.Append("\r\n");

        foreach ( Variable oVariable in mTargetModule.VariableList )
        {
          // string Result = null;
          oResult.Append(Indent4 + oVariable.Scope + " " + oVariable.Type + " " + oVariable.Name + ";\r\n");
        }
      }

      // ********************************************************
      // properties has only forms and classes
      // ********************************************************

      if ((mTargetModule.Type == "form") || (mTargetModule.Type == "class"))
      {
        // properties
        if (mTargetModule.PropertyList.Count > 0)
        {
          // new line
          oResult.Append("\r\n");
          //public string Comment  
          //{
          //  get { return mComment; }
          //  set { mComment = value; }
          //}
          foreach ( Property oProperty in mTargetModule.PropertyList )
          {
            // possible comment
            oResult.Append(oProperty.Comment + ";\r\n");
            // string Result = null;
            oResult.Append(Indent4 + oProperty.Scope + " " + oProperty.Type + " " + oProperty.Name + ";\r\n");
            oResult.Append(Indent4 + "{\r\n");
            oResult.Append(Indent6 + "get { return ; }\r\n");
            oResult.Append(Indent6 + "set {  = value; }\r\n");

            // lines
            foreach ( string Line in oProperty.LineList )
            {
              Temp = Line.Trim(); 
              if (Temp.Length > 0)
              {
                oResult.Append(Indent6 + Temp + ";\r\n");
              }
              else
              {
                oResult.Append("\r\n");
              }
            }
            oResult.Append(Indent4 + "}\r\n");
          }
        }
      }

      // ********************************************************
      // procedures
      // ********************************************************

      if (mTargetModule.ProcedureList.Count > 0)
      {
        oResult.Append("\r\n");
        foreach ( Procedure oProcedure in mTargetModule.ProcedureList )
        {
          // private void WriteResX ( ArrayList mImageList, string OutPath, string ModuleName )
          oResult.Append(Indent4 + oProcedure.Scope + " ");
          switch(oProcedure.Type)
          {
            case PROCEDURE_TYPE.PROCEDURE_SUB:
              oResult.Append("void");
              break;
            case PROCEDURE_TYPE.PROCEDURE_FUNCTION:
              oResult.Append(oProcedure.ReturnType );
              break;
            case PROCEDURE_TYPE.PROCEDURE_EVENT:
              oResult.Append("void");
              break;
          }
          // name
          oResult.Append(" " + oProcedure.Name);
          // parametres
          if ( oProcedure.ParameterList.Count > 0)
          {
          }
          else
          {
            oResult.Append("()\r\n");
          }

          // start body
          oResult.Append(Indent4 + "{\r\n");

          foreach ( string Line in oProcedure.LineList )
          {
            Temp = Line.Trim(); 
            if (Temp.Length > 0)
            {
              oResult.Append(Indent6 + Temp + ";\r\n");
            }
            else
            {
              oResult.Append("\r\n");
            }
          }
          // end procedure
          oResult.Append(Indent4 + "}\r\n");
        }
      }

      // end class
      oResult.Append(Indent2 + "}\r\n");
      // end namespace               
      oResult.Append("}\r\n");

      // return result      
      return oResult.ToString();
    }

    private void WriteResX ( ArrayList mImageList, string OutPath, string ModuleName )
    {
      string sResxName;

      if (mImageList.Count > 0)
      {
        // resx name
        sResxName = OutPath + ModuleName + ".resx";
        // open file
        ResXResourceWriter rsxw = new ResXResourceWriter(sResxName); 

        foreach ( string ResourceName in mImageList)
        {
          try
          {
            Image img = Image.FromFile(OutPath + ResourceName);
            rsxw.AddResource(ResourceName,img);
            img.Dispose();
          }
          catch
          {
          }
        }
        // rsxw.Generate();
        rsxw.Close();

        foreach ( string ResourceName in mImageList)
        {
          File.Delete(OutPath + ResourceName); 
        }
      }
    }

    private bool WriteImage ( Module SourceModule, string ResourceName, string Value, string OutPath )
    {
      string Temp = String.Empty;
      int Offset = 0;
      string FrxFile = String.Empty;
      string sResxName = String.Empty;
      int Position = 0;

      Position = Value.IndexOf(":",0);
      // "Form1.frx":0000;
      // old vb3 code has name without ""
      // CONTROL.FRX:0000

      if (SourceModule.Version == "5.00")
      {
        FrxFile = Value.Substring(1, Position - 2); 
      }
      else
      {
        FrxFile = Value.Substring(0, Position); 
      }

      Temp = Value.Substring(Position + 1, Value.Length - Position - 1); 
      Offset = System.Convert.ToInt32("0x" + Temp, 16);
      // exist file ?

      // get image
      byte[] ImageString;

      Tools.GetFRXImage(Path.GetDirectoryName(SourceModule.FileName) + @"\" + FrxFile, Offset, out ImageString );

      if ((ImageString.GetLength(0) - 8) > 0)
      {
        if (File.Exists(OutPath + ResourceName))
        {
          File.Delete(OutPath + ResourceName);
        }
        FileStream Stream = Stream = new FileStream( OutPath + ResourceName, FileMode.CreateNew, FileAccess.Write );
        BinaryWriter Writer = new BinaryWriter( Stream );
        Writer.Write(ImageString, 8, ImageString.GetLength(0) - 8);           
        Stream.Close();
        Writer.Close(); 

        // write
        mTargetModule.ImageList.Add(ResourceName);  
        return true;
      }
      else
      {
        return false;
      }
      // save it to resx
//      Debug.WriteLine(ModuleName + ", " + ResourceName + ", " + Temp + ", " + oImage.Width.ToString() );  
//      
//      ResXResourceWriter oWriter;
//      
//      sResxName = @"C:\temp\test\" + ModuleName + ".resx";
//      oWriter = new ResXResourceWriter(sResxName);
//     // oImage = Image.FromFile(myRow[sField].ToString());
//      oWriter.AddResource(ResourceName, oImage);
//      oWriter.Generate();
//      oWriter.Close();

//      MemoryStream s = new MemoryStream();
//      oImage.Save(s);
//      
//      byte[] b = s.ToArray();
//      String strContents = System.Security.Cryptography.EncodeAsBase64.EncodeBuffer(b);
//      myXmlTextNode.Value = strContents;

/* Add Convert it back using.... */
//
//byte[]  b =
//System.Security.Cryptography.DecodeBase64.DecodeBuffer(myXmlTextNode.Value);

//      ResourceWriter oWriter;
//      
//      sResxName = @"c:\temp\" + ModuleName + ".resx";
//      oWriter = new ResourceWriter(sResxName);
//     // oImage = Image.FromFile(myRow[sField].ToString());
//      oWriter.AddResource(ResourceName, oImage);
//      oWriter.Close();      
      
    }

    private string GetWord(string Line, ref int Position)
    {
      string Result = null;
      int End = 0;

      if (Position < Line.Length)
      {
        // seach for first space
        End = Line.IndexOf(" ", Position);
        if (End > -1)
        {
          Result = Line.Substring(Position, End - Position); 
          Position = End;
        }
        else
        {
          Result = Line.Substring(Position); 
        }
      }
      return Result;
    }
    
    private void GetPropertyRow( StringBuilder oResult, string Type, 
                                string Name, ControlProperty oProperty, string OutPath)
    {
      // exception for images
      if (oProperty.Name == "Icon" || oProperty.Name == "Image" || oProperty.Name == "BackgroundImage")
      {
        // generate resx file and write there image extracted from VB6 frx file
        string ResourceName = String.Empty;

        //.BackgroundImage = ((System.Drawing.Bitmap)(resources.GetObject("$this.BackgroundImage")));
        //.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
        //.Image = ((System.Drawing.Bitmap)(resources.GetObject("Command1.Image")));

        switch(oProperty.Name)
        {
          case "BackgroundImage":
            ResourceName = "$this.BackgroundImage";
            break;
          case "Icon":
            ResourceName = "$this.Icon";
            break;
          case "Image":
            ResourceName = Name + ".Image";
            break;
        }
        if (WriteImage ( mSourceModule, ResourceName, oProperty.Value, OutPath ))
        {
          switch(oProperty.Name)
          {
            case "BackgroundImage":
              oResult.Append(Indent6 + "this." 
                + oProperty.Name + " = ((System.Drawing.Bitmap)(resources.GetObject(" 
                + (char)34 + "$this.BackgroundImage" + (char)34 + ")));\r\n");
              break;
            case "Icon":
              oResult.Append(Indent6 + "this." 
                + oProperty.Name + " = ((System.Drawing.Icon)(resources.GetObject(" 
                + (char)34 + "$this.Icon" + (char)34 + ")));\r\n");
              break;
            case "Image":
              oResult.Append(Indent6 + "this." + Name + "." 
                + oProperty.Name + " = ((System.Drawing.Bitmap)(resources.GetObject(" 
                + (char)34 + Name + ".Image" + (char)34 + ")));\r\n");
              break;
          }
        }
      }
      else
      {
        // unsupported property
        if (!oProperty.Valid) { oResult.Append("//"); } 
        if (Type == "form")
        {
          // form properties
          oResult.Append(Indent6 + "this." 
            + oProperty.Name + " = " + oProperty.Value + ";\r\n");        
        }
        else
        {
          // control properties
          oResult.Append(Indent6 + "this." + Name + "." 
            + oProperty.Name + " = " + oProperty.Value + ";\r\n");
        }
      }    
    }
	}
}
