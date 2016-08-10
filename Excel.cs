 public static  String retrieve(String Col)
    {
        string path = CommonVariable. TESTDATA.Trim();
        String rowval = null;
	OleDbConnection conn;

    	try
            {
                conn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;" +
                    "Data Source=" + path + "\\TestData.xlsx;" +
                    "Extended Properties=Excel 12.0 Xml");
                conn.Open();
            }
            catch
            {
                try
                {
                    conn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.14.0;" +
                        "Data Source=" + path + "\\TestData.xlsx;" +
                        "Extended Properties=Excel 14.0 Xml");
                    conn.Open();
                }
                catch
                {
                    try
                    {
                        conn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.15.0;" +
                            "Data Source=" + path + "\\TestData.xlsx;" +
                            "Extended Properties=Excel 15.0 Xml");
                        conn.Open();
                    }
                    catch
                    {
                    }
                }
            }

           try
                {
                        try// _SelectValueList(BPS.drp_Program, ExcelRead.retrieve("Program"));    where TestName = "+CommonVariable.TestName+"
                        {
                            OleDbCommand command = new OleDbCommand("select * from ["+CommonVariable._Table+"$] where TestName = '"+CommonVariable.TestName+"'", conn);
                            using (OleDbDataReader dr = command.ExecuteReader())
                            {
                                while (dr.Read())
                                {
                                   rowval = dr[Col].ToString();
                                   // Console.WriteLine(row1Col0);
                                }
                            }
                      
                        }
                        catch(Exception)
                        { }
                        
                    }
                    catch(Exception)
                    {
                    }
                return rowval; 
    }

