using System;
using System.Collections;

namespace VB2C
{
	/// <summary>
	/// Summary description for Enum.
	/// </summary>
	public class Enum
	{
    private string mName;
    private string mScope;
    private ArrayList mItemList;

    public Enum()
    {
      mItemList = new ArrayList();
    }
	  
    public string Name 
    {
      get { return mName; }
      set { mName = value; }
    }

    public string Scope 
    {
      get { return mScope; }
      set { mScope = value; }
    }	 

    public ArrayList ItemList 
    {
      get { return mItemList; }
      set { mItemList = value; }
    } 
	}
}
