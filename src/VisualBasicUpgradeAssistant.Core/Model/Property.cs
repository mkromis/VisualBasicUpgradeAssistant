using System;
using System.Collections;

namespace VB2C
{
	/// <summary>
	/// Summary description for Property.
	/// </summary>
	public class Property
	{
    private string mName;
	  private string mType;
    private string mDirection;
    private string mComment;
	  private string mScope;
	  private ArrayList mParameterList;
    private ArrayList mLineList;

    public Property()
    {
      mLineList = new ArrayList();
      mParameterList = new ArrayList();
    }

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

    public string Direction 
    {
      get { return mDirection; }
      set { mDirection = value; }
    }

    public string Comment  
    {
      get { return mComment; }
      set { mComment = value; }
    }

    public string Scope 
    {
      get { return mScope; }
      set { mScope = value; }
    }	 

    public ArrayList ParameterList 
    {
      get { return mParameterList; }
      set { mParameterList = value; }
    }  
    
    public ArrayList LineList 
    {
      get { return mLineList; }
      set { mLineList = value; }
    } 
	}
}
