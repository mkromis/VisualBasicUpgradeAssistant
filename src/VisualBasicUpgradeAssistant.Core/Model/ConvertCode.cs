using System;
using System.Collections;
using System.Drawing;
using System.IO;
using System.Resources;
using System.Text;
using VisualBasicUpgradeAssistant.Core.DataClasses;

namespace VisualBasicUpgradeAssistant.Core.Model
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
        private VB_FILE_TYPE _fileType;
        private Module _sourceModule;
        private Module _targetModule;
        private readonly ArrayList _ownerStock;
        private readonly Tools _tools;
        private const String FORM_FIRST_LINE = "VERSION 5.00";
        private const String MODULE_FIRST_LINE = "ATTRIBUTE";
        private const String CLASS_FIRST_LINE = "1.0 CLASS";

        private const String Indent2 = "  ";
        private const String Indent4 = "    ";
        private const String Indent6 = "      ";

        public ConvertCode()
        {
            _ownerStock = new ArrayList();
            _tools = new Tools();
        }

        public String ActionResult { get; private set; }

        public String OutSourceCode { get; private set; }

        public Boolean ParseFile(String fileName, String outPath)
        {
            String line = String.Empty;
            String temp = String.Empty;
            String extension = String.Empty;
            String version = String.Empty;
            Boolean result = false;
            Int32 position = 0;

            // try recognize source code type depend by file extension
            extension = fileName.Substring(fileName.Length - 3, 3);
            _fileType = (extension.ToUpper()) switch
            {
                "FRM" => VB_FILE_TYPE.VB_FILE_FORM,
                "BAS" => VB_FILE_TYPE.VB_FILE_MODULE,
                "CLS" => VB_FILE_TYPE.VB_FILE_CLASS,
                _ => VB_FILE_TYPE.VB_FILE_UNKNOWN,
            };

            // open file
            FileStream stream = new FileStream(fileName, FileMode.Open, FileAccess.Read);
            StreamReader reader = new StreamReader(stream);
            // Start from begin.
            reader.BaseStream.Seek(0, SeekOrigin.Begin);
            line = reader.ReadLine();
            // verify type of file based on first line - form, module, class

            // get first word from first line
            position = 0;
            temp = GetWord(line, ref position);
            switch (temp.ToUpper())
            {
                // module first line
                // 'Attribute VB_Name = "ModuleName"'
                case MODULE_FIRST_LINE:
                    _fileType = VB_FILE_TYPE.VB_FILE_MODULE;
                    break;
                // form or class first line
                // 'VERSION 5.00' or 'VERSION 1.0 CLASS' 
                case "VERSION":
                    position++;
                    version = GetWord(line, ref position);

                    //            Debug.WriteLine (Line + " " + Version);

                    if (line.IndexOf(CLASS_FIRST_LINE, 0) > -1)
                        _fileType = VB_FILE_TYPE.VB_FILE_CLASS;
                    else
                        _fileType = VB_FILE_TYPE.VB_FILE_FORM;
                    break;
                default:
                    _fileType = VB_FILE_TYPE.VB_FILE_UNKNOWN;
                    break;
            }
            // if file is still unknown
            if (_fileType == VB_FILE_TYPE.VB_FILE_UNKNOWN)
            {
                ActionResult = "Unknown file type";
                return false;
            }

            _sourceModule = new Module
            {
                Version = version,
                FileName = fileName
            };

            // now parse specifics of each type
            switch (extension.ToUpper())
            {
                case "FRM":
                    _sourceModule.Type = "form";
                    result = ParseForm(reader);
                    break;
                case "BAS":
                    _sourceModule.Type = "module";
                    result = ParseModule(reader);
                    break;
                case "CLS":
                    _sourceModule.Type = "class";
                    result = ParseClass(reader);
                    break;
            }
            // parse remain - variables, functions, procedures
            result = ParseProcedures(reader);

            stream.Close();
            reader.Close();

            // generate output file
            OutSourceCode = GetOutSourceCode(outPath);

            // save result
            String OutFileName = outPath + _targetModule.FileName;
            stream = new FileStream(OutFileName, FileMode.OpenOrCreate);
            StreamWriter Writer = new StreamWriter(stream);
            Writer.Write(OutSourceCode);
            Writer.Close();
            // generate resx file if source form contain any images

            if (_targetModule.ImagesUsed)
                WriteResX(_targetModule.ImageList, outPath, _targetModule.Name);

            return result;
        }

        private Boolean ParseForm(StreamReader oReader)
        {
            Boolean bProcess = false;
            Boolean bFinish = false;
            Boolean bEnd = true;
            String sLine = null;
            String sName = null;
            String sOwner = null;
            String sTemp = null;
            Int32 iTemp = 0;
            Int32 iComment = 0;
            String sType = null;
            Int32 iPosition = 0;
            Int32 iLevel = 0;
            ControlType oControl = null;
            ControlProperty oNestedProperty = null;
            Boolean bNestedProperty = false;

            // parse only visual part of form
            while (!bFinish) // ( ( bFinish || (oReader.Peek() > -1)) )
            {
                sLine = oReader.ReadLine();
                sLine = sLine.Trim();
                iPosition = 0;
                // get first word in line
                sTemp = GetWord(sLine, ref iPosition);
                switch (sTemp)
                {
                    case "Begin":
                        bProcess = true;
                        // new level
                        iLevel++;
                        // next word - control type
                        iPosition++;
                        sType = GetWord(sLine, ref iPosition);
                        // next word - control name
                        iPosition++;
                        sName = GetWord(sLine, ref iPosition);
                        // detected missing end -> it indicate that next control is container
                        if (!bEnd)
                        {
                            // add container control to colection
                            if (!(oControl == null))
                            {
                                oControl.Container = true;
                                _sourceModule.ControlAdd(oControl);
                            }
                            // save name of previous control as owner for current and next controls
                            _ownerStock.Add(sOwner);
                        }
                        bEnd = false;

                        switch (sType)
                        {
                            case "Form":            // VERSION 2.00 - VB3
                            case "VB.Form":
                            case "VB.MDIForm":
                                _sourceModule.Name = sName;
                                // first owner
                                // save control name for possible next controls as owner
                                sOwner = sName;
                                break;
                            default:
                                // new control
                                oControl = new ControlType
                                {
                                    Name = sName,
                                    Type = sType
                                };
                                // save control name for possible next controls as owner
                                sOwner = sName;
                                // set current container name
                                oControl.Owner = (String)_ownerStock[_ownerStock.Count - 1];
                                break;
                        }
                        break;

                    case "End":
                        // double end - we leaving some container
                        if (bEnd)
                            // remove last item from stock
                            _ownerStock.Remove((String)_ownerStock[_ownerStock.Count - 1]);
                        else
                            // level 1 is form and all higher levels are controls
                            if (iLevel > 1)
                            // add control to colection
                            _sourceModule.ControlAdd(oControl);
                        // form or control end detected
                        bEnd = true;
                        // back to previous level
                        iLevel--;

                        break;

                    case "Object":
                        // used controls in form
                        break;

                    case "BeginProperty":
                        bNestedProperty = true;

                        oNestedProperty = new ControlProperty();
                        // next word - nested property name
                        iPosition++;
                        sName = GetWord(sLine, ref iPosition);
                        oNestedProperty.Name = sName;
                        //            Debug.WriteLine(sName); 
                        break;

                    case "EndProperty":
                        bNestedProperty = false;
                        // add property to control or form
                        if (iLevel == 1)
                            // add property to form
                            _sourceModule.FormPropertyAdd(oNestedProperty);
                        else
                            // to controls
                            oControl.PropertyAdd(oNestedProperty);
                        break;

                    default:
                        // parse property
                        ControlProperty oProperty = new ControlProperty();

                        iTemp = sLine.IndexOf("=");
                        if (iTemp > -1)
                        {
                            oProperty.Name = sLine.Substring(0, iTemp - 1).Trim();
                            iComment = sLine.IndexOf("'", iTemp);
                            if (iComment > -1)
                            {
                                oProperty.Value = sLine.Substring(iTemp + 1, iComment - iTemp - 1).Trim();
                                oProperty.Comment = sLine.Substring(iComment + 1, sLine.Length - iComment - 1).Trim();
                            }
                            else
                                oProperty.Value = sLine.Substring(iTemp + 1, sLine.Length - iTemp - 1).Trim();

                            if (bNestedProperty)
                                oNestedProperty.PropertyList.Add(oProperty);
                            else
                                // depend by level insert property to form or control
                                if (iLevel > 1)
                                // add property to control
                                oControl.PropertyAdd(oProperty);
                            else
                                // add property to form
                                _sourceModule.FormPropertyAdd(oProperty);
                        }
                        break;
                }

                if (iLevel == 0 && bProcess)
                    // visual part of form is finish
                    bFinish = true;

            }
            return true;
        }

        private Boolean ParseModule(StreamReader reader)
        {
            String line = String.Empty;
            Int32 position = 0;

            // name of module
            // Attribute VB_Name = "ModuleName"

            // Start from begin again
            reader.DiscardBufferedData();
            reader.BaseStream.Seek(0, SeekOrigin.Begin);
            // search for module name
            while (reader.Peek() > -1)
            {
                line = reader.ReadLine();
                position = line.IndexOf('"');
                _sourceModule.Name = line.Substring(position + 1, line.Length - position - 2);
                return true;
            }
            return false;
        }

        private Boolean ParseClass(StreamReader reader)
        {
            Int32 Position = 0;
            String line = null;
            String tempString = String.Empty;

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
            reader.DiscardBufferedData();
            reader.BaseStream.Seek(0, SeekOrigin.Begin);
            while (reader.Peek() > -1)
            {
                Position = 0;
                // verify type of file based on first line
                // form, module, class
                line = reader.ReadLine();
                // next word - control type
                tempString = GetWord(line, ref Position);
                if (tempString == "Attribute")
                {
                    Position++;
                    tempString = GetWord(line, ref Position);
                    switch (tempString)
                    {
                        case "VB_Name":
                            Position++;
                            tempString = GetWord(line, ref Position);
                            Position++;
                            _sourceModule.Name = GetWord(line, ref Position);
                            break;

                        case "VB_Exposed":
                            return true;
                            //break;
                    }
                }
            }
            return false;
        }

        private Boolean ParseProcedures(StreamReader reader)
        {
            String line = null;
            String tempString = null;
            String comments = null;
            String scope = null;

            Int32 position = 0;
            //bool bProcess = false;

            Boolean bEnum = false;
            Boolean bVariable = false;
            Boolean bProperty = false;
            Boolean bProcedure = false;
            Boolean bEnd = false;

            Variable oVariable = null;
            Property oProperty = null;
            Procedure oProcedure = null;
            EnumType oEnum = null;
            EnumItem oEnumItem = null;

            while (reader.Peek() > -1)
            {
                line = reader.ReadLine();
                //Line = Line.Trim();  

                position = 0;

                if (line != null && line != String.Empty)
                    // check if next line is same command, join it together ?
                    while (line.Substring(line.Length - 1, 1) == "_")
                        line = line + reader.ReadLine();
                // : is command delimiter


                //  Debug.WriteLine(Line); 

                // get first word in line
                tempString = GetWord(line, ref position);
                switch (tempString)
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
                        comments = comments + line + "\r\n";
                        break;

                    // next can be declaration of variables   

                    //Private mlParentID As Long
                    //Private mlOwnerType As ENUM_FORM_TYPE
                    //Private moAttachement As Attachement

                    case "Public":
                    case "Private":
                        // save it for later use
                        scope = tempString.ToLower();
                        // read next word
                        // next word - control type
                        position++;
                        tempString = GetWord(line, ref position);

                        switch (tempString)
                        {
                            // functions or procedures
                            case "Sub":
                            case "Function":

                                oProcedure = new Procedure
                                {
                                    Comment = comments
                                };
                                comments = String.Empty;
                                ParseProcedureName(oProcedure, line);

                                bProcedure = true;
                                break;

                            case "Enum":
                                oEnum = new EnumType
                                {
                                    Scope = scope
                                };
                                // next word is enum name
                                position++;
                                oEnum.Name = GetWord(line, ref position);
                                bEnum = true;
                                break;

                            case "Property":
                                oProperty = new Property
                                {
                                    Comment = comments
                                };
                                comments = String.Empty;
                                ParsePropertyName(oProperty, line);
                                bProperty = true;

                                break;
                            default:
                                // variable declaration
                                oVariable = new Variable();
                                ParseVariableDeclaration(oVariable, line);
                                bVariable = true;
                                break;
                        }

                        break;

                    case "Dim":
                        // variable declaration
                        oVariable = new Variable
                        {
                            Comment = comments
                        };
                        comments = String.Empty;
                        ParseVariableDeclaration(oVariable, line);
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
                            oEnumItem = new EnumItem
                            {
                                Comment = comments
                            };
                            comments = String.Empty;
                            ParseEnumItem(oEnumItem, line);
                            // add item
                            oEnum.ItemList.Add(oEnumItem);
                        }
                        if (bProperty)
                            // add line of property
                            oProperty.LineList.Add(line);
                        if (bProcedure)
                            oProcedure.LineList.Add(line);
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
                        _sourceModule.EnumList.Add(oEnum);
                        bEnum = false;
                    }
                    if (bProperty)
                    {
                        _sourceModule.PropertyAdd(oProperty);
                        bProperty = false;
                    }
                    if (bProcedure)
                    {
                        _sourceModule.ProcedureAdd(oProcedure);
                        bProcedure = false;
                    }
                    bEnd = false;
                }
                else
                    if (bVariable)
                    _sourceModule.VariableAdd(oVariable);

                bVariable = false;
            }

            return true;
        }

        //Public Enum ENUM_BUG_LEVEL
        //  BUG_LEVEL_PROJECT = 1
        //  BUG_LEVEL_VERSION = 2
        //End Enum
        private void ParseEnumItem(EnumItem oEnumItem, String line)
        {
            String TempString = String.Empty;
            Int32 iPosition = 0;

            line = line.Trim();
            // first word is ame
            oEnumItem.Name = GetWord(line, ref iPosition);
            iPosition++;
            // next word =
            TempString = GetWord(line, ref iPosition);
            iPosition++;
            // optional
            oEnumItem.Value = GetWord(line, ref iPosition);
        }


        //Private mlID As Long

        private void ParseVariableDeclaration(Variable oVariable, String line)
        {
            String TempString = String.Empty;
            Int32 iPosition = 0;
            Boolean Status = false;

            // next word - control type
            TempString = GetWord(line, ref iPosition);
            switch (TempString)
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
                iPosition++;
                TempString = GetWord(line, ref iPosition);
                oVariable.Name = TempString;
            }
            // As 
            iPosition++;
            TempString = GetWord(line, ref iPosition);
            // variable type
            iPosition++;
            TempString = GetWord(line, ref iPosition);
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

        private void ParsePropertyName(Property oProperty, String line)
        {
            String TempString = String.Empty;
            Int32 iPosition = 0;
            Int32 Start = 0;
            Boolean Status = false;

            // next word - control type
            TempString = GetWord(line, ref iPosition);
            switch (TempString)
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
                iPosition++;
                TempString = GetWord(line, ref iPosition);
            }

            // direction Let,Get, Set
            iPosition++;
            TempString = GetWord(line, ref iPosition);
            oProperty.Direction = TempString;

            //Public Property Let ParentID(ByVal lValue As Long)

            // name       
            Start = iPosition;
            iPosition = line.IndexOf("(", Start + 1);
            oProperty.Name = line.Substring(Start, iPosition - Start);

            // + possible parameters
            iPosition++;
            Start = iPosition;
            iPosition = line.IndexOf(")", Start);

            if (iPosition - Start > 0)
            {
                TempString = line.Substring(Start, iPosition - Start);
                ArrayList ParameterList = new ArrayList();
                // process parametres
                ParseParametries(ParameterList, TempString);
                oProperty.ParameterList = ParameterList;
            }

            // As 
            iPosition++;
            iPosition++;
            TempString = GetWord(line, ref iPosition);

            // type
            iPosition++;
            TempString = GetWord(line, ref iPosition);
            oProperty.Type = TempString;

        }

        // ByVal lValue As Long, ByVal sValue As string

        private void ParseParametries(ArrayList parametreList, String line)
        {
            Boolean bFinish = false;
            Int32 Position = 0;
            Int32 Start = 0;
            Boolean Status = false;
            String TempString = String.Empty;

            // parameters delimited by comma
            while (!bFinish)
            {
                Parameter oParameter = new Parameter();

                // next word - control type
                TempString = GetWord(line, ref Position);
                switch (TempString)
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
                    Position++;
                    TempString = GetWord(line, ref Position);
                    oParameter.Name = TempString;
                }
                // As 
                Position++;
                TempString = GetWord(line, ref Position);
                // parameter type
                Position++;
                TempString = GetWord(line, ref Position);
                oParameter.Type = TempString;

                parametreList.Add(oParameter);

                // next parameter
                Position++;
                Start = Position;
                Position = line.IndexOf(",", Start);

                if (Position == -1)
                    // end
                    bFinish = true;

            }
        }

        private void ParseProcedureName(Procedure procedure, String line)
        {
            String tempString = String.Empty;
            Int32 position = 0;
            Int32 start = 0;
            Boolean status = false;

            //Private Sub cmdOk_Click()
            //private void cmdShow_Click(object sender, System.EventArgs e)

            //Private Sub Form_Load()
            //private void frmConvert_Load(object sender, System.EventArgs e)

            //Public Function Rozbor_DefaultFields(ByVal MKf As String) As String
            //public static bool ParseProcedures( Module SourceModule, Module TargetModule )

            tempString = GetWord(line, ref position);
            switch (tempString)
            {
                case "Private":
                    procedure.Scope = "private";
                    status = true;
                    break;
                case "Public":
                    procedure.Scope = "public";
                    status = true;
                    break;
                default:
                    procedure.Scope = "private";
                    status = true;
                    break;
            }

            if (status)
            {
                // property
                position++;
                tempString = GetWord(line, ref position);
            }

            // procedure type
            switch (tempString)
            {
                case "Sub":
                    procedure.Type = ProcedureType.PROCEDURE_SUB;
                    break;
                case "Function":
                    procedure.Type = ProcedureType.PROCEDURE_FUNCTION;
                    break;
                case "Event":
                    procedure.Type = ProcedureType.PROCEDURE_EVENT;
                    break;
            }

            // next is name
            position++;
            start = position;
            position = line.IndexOf("(", start);
            procedure.Name = line.Substring(start, position - start);

            // next possible parameters
            position++;
            start = position;
            position = line.IndexOf(")", start);

            if (position - start > 0)
            {
                tempString = line.Substring(start, position - start);
                ArrayList ParameterList = new ArrayList();
                // process parametres
                ParseParametries(ParameterList, tempString);
                procedure.ParameterList = ParameterList;
            }

            // and return type of function
            if (procedure.Type == ProcedureType.PROCEDURE_FUNCTION)
            {
                // as
                position++;
                tempString = GetWord(line, ref position);
                // function return type
                position++;
                procedure.ReturnType = GetWord(line, ref position);
            }

        }


        // generate result file
        // OutPath for pictures
        private String GetOutSourceCode(String outPath)
        {
            StringBuilder oResult = new StringBuilder();

            String temp = String.Empty;

            // convert source to target    
            _targetModule = new Module();
            _tools.ParseModule(_sourceModule, _targetModule);

            // ********************************************************
            // common class
            // ********************************************************
            oResult.Append("using System;\r\n");

            // ********************************************************
            // only form class
            // ********************************************************
            if (_targetModule.Type == "form")
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
            oResult.Append(Indent2 + "/// Summary description for " + _sourceModule.Name + ".\r\n");
            oResult.Append(Indent2 + "/// </summary>\r\n");


            switch (_targetModule.Type)
            {
                case "form":
                    oResult.Append(Indent2 + "public class " + _sourceModule.Name + " : System.Windows.Forms.Form\r\n");
                    break;
                case "module":
                    oResult.Append(Indent2 + "sealed class " + _sourceModule.Name + "\r\n");
                    // all procedures must be static
                    break;
                case "class":
                    oResult.Append(Indent2 + "public class " + _sourceModule.Name + "\r\n");
                    break;
            }
            // start class region
            oResult.Append(Indent2 + "{\r\n");

            // ********************************************************
            // only form class
            // ********************************************************

            if (_targetModule.Type == "form")
            {
                // list of controls
                foreach (ControlType oControl in _targetModule.ControlList)
                {
                    if (!oControl.Valid)
                        oResult.Append("//");
                    oResult.Append(Indent2 + " private System.Windows.Forms." + oControl.Type + " " + oControl.Name + ";\r\n");
                }

                oResult.Append(Indent4 + "/// <summary>\r\n");
                oResult.Append(Indent4 + "/// Required designer variable.\r\n");
                oResult.Append(Indent4 + "/// </summary>\r\n");
                oResult.Append(Indent4 + "private System.ComponentModel.Container components = null;\r\n");
                oResult.Append("\r\n");
                oResult.Append(Indent4 + "public " + _sourceModule.Name + "()\r\n");
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
                if (_targetModule.ImagesUsed)
                    // System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(Form1));
                    oResult.Append(Indent6 + "System.Resources.ResourceManager resources = " +
                      "new System.Resources.ResourceManager(typeof(" + _targetModule.Name + "));\r\n");

                foreach (ControlType oControl in _targetModule.ControlList)
                {
                    if (!oControl.Valid)
                        oResult.Append("//");
                    oResult.Append(Indent6 + "this." + oControl.Name
                      + " = new System.Windows.Forms." + oControl.Type
                      + "();\r\n");

                }

                // SuspendLayout part
                oResult.Append(Indent6 + "this.SuspendLayout();\r\n");
                // this.Frame1.ResumeLayout(false);
                // resume layout for each container
                foreach (ControlType oControl in _targetModule.ControlList)
                    // check if control is container
                    // !! for menu controls
                    if (oControl.Container && !(oControl.Type == "MenuItem") && !(oControl.Type == "MainMenu"))
                    {
                        if (!oControl.Valid)
                            oResult.Append("//");
                        oResult.Append(Indent6 + "this." + oControl.Name + ".SuspendLayout();\r\n");
                    }

                // each controls and his property		  
                foreach (ControlType oControl in _targetModule.ControlList)
                {
                    oResult.Append(Indent6 + "//\r\n");
                    oResult.Append(Indent6 + "// " + oControl.Name + "\r\n");
                    oResult.Append(Indent6 + "//\r\n");

                    // unsupported control
                    if (!oControl.Valid)
                        oResult.Append("/*");
                    // ImageList, Timer, Menu has't name property
                    if (oControl.Type != "ImageList" && oControl.Type != "Timer"
                      && oControl.Type != "MenuItem" && oControl.Type != "MainMenu")
                        // control name
                        oResult.Append(Indent6 + "this." + oControl.Name + ".Name = "
                          + (Char)34 + oControl.Name + (Char)34 + ";\r\n");

                    // write properties
                    foreach (ControlProperty oProperty in oControl.PropertyList)
                        GetPropertyRow(oResult, oControl.Type, oControl.Name, oProperty, outPath);

                    // if control is container for other controls
                    temp = String.Empty;
                    foreach (ControlType oControl1 in _targetModule.ControlList)
                        // all controls ownered by current control
                        if (oControl1.Owner == oControl.Name && !oControl1.InvisibleAtRuntime)
                            temp = temp + Indent6 + Indent6 + "this." + oControl1.Name + ",\r\n";
                    if (temp != String.Empty)
                    {
                        // exception for menu controls
                        if (oControl.Type == "MainMenu" || oControl.Type == "MenuItem")
                            // this.mainMenu1.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] 
                            oResult.Append(Indent6 + "this." + oControl.Name
                              + ".MenuItems.AddRange(new System.Windows.Forms.MenuItem[]\r\n");
                        else
                            // this. + oControl.Name + .Controls.AddRange(new System.Windows.Forms.Control[]
                            oResult.Append(Indent6 + "this." + oControl.Name
                              + ".Controls.AddRange(new System.Windows.Forms.Control[]\r\n");

                        oResult.Append(Indent6 + "{\r\n");
                        oResult.Append(temp);
                        // remove last comma, keep CRLF
                        oResult.Remove(oResult.Length - 3, 1);
                        // close addrange part  
                        oResult.Append(Indent6 + "});\r\n");
                    }
                    // unsupported control
                    if (!oControl.Valid)
                        oResult.Append("*/");
                }

                oResult.Append(Indent6 + "//\r\n");
                oResult.Append(Indent6 + "// " + _sourceModule.Name + "\r\n");
                oResult.Append(Indent6 + "//\r\n");
                oResult.Append(Indent6 + "this.Controls.AddRange(new System.Windows.Forms.Control[]\r\n");
                oResult.Append(Indent6 + "{\r\n");


                // add control range to form
                foreach (ControlType oControl in _targetModule.ControlList)
                {
                    if (!oControl.Valid)
                        oResult.Append("//");
                    // all controls ownered by main form
                    if (oControl.Owner == _sourceModule.Name && !oControl.InvisibleAtRuntime)
                        oResult.Append(Indent6 + "      this." + oControl.Name + ",\r\n");
                }

                // remove last comma, keep CRLF
                oResult.Remove(oResult.Length - 3, 1);
                // close addrange part  
                oResult.Append(Indent6 + "});\r\n");

                // form name
                oResult.Append(Indent6 + "this.Name = " + (Char)34 + _targetModule.Name + (Char)34 + ";\r\n");
                // exception for menu
                // this.Menu = this.mainMenu1;
                if (_targetModule.MenuUsed)
                    foreach (ControlType oControl in _targetModule.ControlList)
                        if (oControl.Type == "MainMenu")
                            oResult.Append(Indent6 + "      this.Menu = " + oControl.Name + ";\r\n");
                // form properties
                foreach (ControlProperty oProperty in _targetModule.FormPropertyList)
                {
                    if (!oProperty.Valid)
                        oResult.Append("//");
                    GetPropertyRow(oResult, _targetModule.Type, "", oProperty, outPath);
                }

                // this.CancelButton = this.cmdExit;


                // this.Frame1.ResumeLayout(false);
                // resume layout for each container
                foreach (ControlType oControl in _targetModule.ControlList)
                    // check if control is container
                    if (oControl.Container && !(oControl.Type == "MenuItem") && !(oControl.Type == "MainMenu"))
                    {
                        if (!oControl.Valid)
                            oResult.Append("//");
                        oResult.Append(Indent6 + "this." + oControl.Name + ".ResumeLayout(false);\r\n");
                    }
                // form
                oResult.Append(Indent6 + "this.ResumeLayout(false);\r\n");

                oResult.Append(Indent4 + "}\r\n");
                oResult.Append(Indent4 + "#endregion\r\n");
            } // if (mTargetModule.Type = "form")


            // ********************************************************
            // enums
            // ********************************************************

            if (_targetModule.EnumList.Count > 0)
            {
                oResult.Append("\r\n");
                foreach (EnumType oEnum in _targetModule.EnumList)
                {
                    // public enum VB_FILE_TYPE
                    oResult.Append(Indent4 + oEnum.Scope + " enum " + oEnum.Name + "\r\n");
                    oResult.Append(Indent4 + "{\r\n");

                    foreach (EnumItem oEnumItem in oEnum.ItemList)
                    {
                        // name
                        oResult.Append(Indent6 + oEnumItem.Name);

                        if (oEnumItem.Value != String.Empty)
                            oResult.Append(" = " + oEnumItem.Value);
                        // enum items delimiter
                        oResult.Append(",\r\n");

                    }
                    // remove last comma, keep CRLF
                    oResult.Remove(oResult.Length - 3, 1);
                    // end enum
                    oResult.Append(Indent4 + "};\r\n");
                }
            }

            // ********************************************************
            //  variables for al module types
            // ********************************************************

            if (_targetModule.VariableList.Count > 0)
            {
                oResult.Append("\r\n");

                foreach (Variable oVariable in _targetModule.VariableList)
                    // string Result = null;
                    oResult.Append(Indent4 + oVariable.Scope + " " + oVariable.Type + " " + oVariable.Name + ";\r\n");
            }

            // ********************************************************
            // properties has only forms and classes
            // ********************************************************

            if (_targetModule.Type == "form" || _targetModule.Type == "class")
                // properties
                if (_targetModule.PropertyList.Count > 0)
                {
                    // new line
                    oResult.Append("\r\n");
                    //public string Comment  
                    //{
                    //  get { return mComment; }
                    //  set { mComment = value; }
                    //}
                    foreach (Property oProperty in _targetModule.PropertyList)
                    {
                        // possible comment
                        oResult.Append(oProperty.Comment + ";\r\n");
                        // string Result = null;
                        oResult.Append(Indent4 + oProperty.Scope + " " + oProperty.Type + " " + oProperty.Name + ";\r\n");
                        oResult.Append(Indent4 + "{\r\n");
                        oResult.Append(Indent6 + "get { return ; }\r\n");
                        oResult.Append(Indent6 + "set {  = value; }\r\n");

                        // lines
                        foreach (String Line in oProperty.LineList)
                        {
                            temp = Line.Trim();
                            if (temp.Length > 0)
                                oResult.Append(Indent6 + temp + ";\r\n");
                            else
                                oResult.Append("\r\n");
                        }
                        oResult.Append(Indent4 + "}\r\n");
                    }
                }

            // ********************************************************
            // procedures
            // ********************************************************

            if (_targetModule.ProcedureList.Count > 0)
            {
                oResult.Append("\r\n");
                foreach (Procedure oProcedure in _targetModule.ProcedureList)
                {
                    // private void WriteResX ( ArrayList mImageList, string OutPath, string ModuleName )
                    oResult.Append(Indent4 + oProcedure.Scope + " ");
                    switch (oProcedure.Type)
                    {
                        case ProcedureType.PROCEDURE_SUB:
                            oResult.Append("void");
                            break;
                        case ProcedureType.PROCEDURE_FUNCTION:
                            oResult.Append(oProcedure.ReturnType);
                            break;
                        case ProcedureType.PROCEDURE_EVENT:
                            oResult.Append("void");
                            break;
                    }
                    // name
                    oResult.Append(" " + oProcedure.Name);
                    // parametres
                    if (oProcedure.ParameterList.Count > 0)
                    {
                    }
                    else
                        oResult.Append("()\r\n");

                    // start body
                    oResult.Append(Indent4 + "{\r\n");

                    foreach (String Line in oProcedure.LineList)
                    {
                        temp = Line.Trim();
                        if (temp.Length > 0)
                            oResult.Append(Indent6 + temp + ";\r\n");
                        else
                            oResult.Append("\r\n");
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

        private void WriteResX(ArrayList imageList, String outPath, String moduleName)
        {
            String sResxName;

            if (imageList.Count > 0)
            {
                // resx name
                sResxName = outPath + moduleName + ".resx";
                // open file
                ResXResourceWriter rsxw = new ResXResourceWriter(sResxName);

                foreach (String ResourceName in imageList)
                    try
                    {
                        Image img = Image.FromFile(outPath + ResourceName);
                        rsxw.AddResource(ResourceName, img);
                        img.Dispose();
                    }
                    catch
                    {
                    }
                // rsxw.Generate();
                rsxw.Close();

                foreach (String ResourceName in imageList)
                    File.Delete(outPath + ResourceName);
            }
        }

        private Boolean WriteImage(Module sourceModule, String resourceName, String value, String outPath)
        {
            String Temp = String.Empty;
            Int32 Offset = 0;
            String FrxFile = String.Empty;
            String sResxName = String.Empty;
            Int32 Position = 0;

            Position = value.IndexOf(":", 0);
            // "Form1.frx":0000;
            // old vb3 code has name without ""
            // CONTROL.FRX:0000

            if (sourceModule.Version == "5.00")
                FrxFile = value.Substring(1, Position - 2);
            else
                FrxFile = value.Substring(0, Position);

            Temp = value.Substring(Position + 1, value.Length - Position - 1);
            Offset = Convert.ToInt32("0x" + Temp, 16);
            // exist file ?

            // get image

            _tools.GetFRXImage(Path.GetDirectoryName(sourceModule.FileName) + @"\" + FrxFile, Offset, out Byte[] imageString);

            if (imageString.GetLength(0) - 8 > 0)
            {
                if (File.Exists(outPath + resourceName))
                    File.Delete(outPath + resourceName);
                FileStream Stream = Stream = new FileStream(outPath + resourceName, FileMode.CreateNew, FileAccess.Write);
                BinaryWriter Writer = new BinaryWriter(Stream);
                Writer.Write(imageString, 8, imageString.GetLength(0) - 8);
                Stream.Close();
                Writer.Close();

                // write
                _targetModule.ImageList.Add(resourceName);
                return true;
            }
            else
                return false;
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

        private String GetWord(String line, ref Int32 position)
        {
            String Result = null;
            Int32 End = 0;

            if (position < line.Length)
            {
                // seach for first space
                End = line.IndexOf(" ", position);
                if (End > -1)
                {
                    Result = line.Substring(position, End - position);
                    position = End;
                }
                else
                    Result = line.Substring(position);
            }
            return Result;
        }

        private void GetPropertyRow(StringBuilder result, String type,
                                    String name, ControlProperty property, String outPath)
        {
            // exception for images
            if (property.Name == "Icon" || property.Name == "Image" || property.Name == "BackgroundImage")
            {
                // generate resx file and write there image extracted from VB6 frx file
                String ResourceName = String.Empty;

                //.BackgroundImage = ((System.Drawing.Bitmap)(resources.GetObject("$this.BackgroundImage")));
                //.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
                //.Image = ((System.Drawing.Bitmap)(resources.GetObject("Command1.Image")));

                switch (property.Name)
                {
                    case "BackgroundImage":
                        ResourceName = "$this.BackgroundImage";
                        break;
                    case "Icon":
                        ResourceName = "$this.Icon";
                        break;
                    case "Image":
                        ResourceName = name + ".Image";
                        break;
                }
                if (WriteImage(_sourceModule, ResourceName, property.Value, outPath))
                    switch (property.Name)
                    {
                        case "BackgroundImage":
                            result.Append(Indent6 + "this."
                              + property.Name + " = ((System.Drawing.Bitmap)(resources.GetObject("
                              + (Char)34 + "$this.BackgroundImage" + (Char)34 + ")));\r\n");
                            break;
                        case "Icon":
                            result.Append(Indent6 + "this."
                              + property.Name + " = ((System.Drawing.Icon)(resources.GetObject("
                              + (Char)34 + "$this.Icon" + (Char)34 + ")));\r\n");
                            break;
                        case "Image":
                            result.Append(Indent6 + "this." + name + "."
                              + property.Name + " = ((System.Drawing.Bitmap)(resources.GetObject("
                              + (Char)34 + name + ".Image" + (Char)34 + ")));\r\n");
                            break;
                    }
            }
            else
            {
                // unsupported property
                if (!property.Valid)
                    result.Append("//");
                if (type == "form")
                    // form properties
                    result.Append(Indent6 + "this."
                      + property.Name + " = " + property.Value + ";\r\n");
                else
                    // control properties
                    result.Append(Indent6 + "this." + name + "."
                      + property.Name + " = " + property.Value + ";\r\n");
            }
        }
    }
}
