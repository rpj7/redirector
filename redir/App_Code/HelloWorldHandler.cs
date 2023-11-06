using System;
using System.Collections.Generic;
using System.Web;
using System.Data;
using System.Data.OleDb;
using System.Data.Odbc;

public class HelloWorldHandler : IHttpHandler
{
    public HelloWorldHandler()
    {
    }
    public void ProcessRequest(HttpContext context)
    {
        HttpRequest Request = context.Request;
        HttpResponse Response = context.Response;
        String shortURL = Request.FilePath.ToString().Replace ("/","");
        String longURL = checkfile (shortURL.ToLower() );
        Response.ClearHeaders(); 
        Response.AppendHeader("Cache-Control", "no-cache"); //HTTP 1.1
        Response.AppendHeader("Cache-Control", "private"); // HTTP 1.1
        Response.AppendHeader("Cache-Control", "no-store"); // HTTP 1.1
        Response.AppendHeader("Cache-Control", "must-revalidate"); // HTTP 1.1
        Response.AppendHeader("Cache-Control", "max-stale=0"); // HTTP 1.1 
        Response.AppendHeader("Cache-Control", "post-check=0"); // HTTP 1.1 
        Response.AppendHeader("Cache-Control", "pre-check=0"); // HTTP 1.1 
        Response.AppendHeader("Pragma", "no-cache"); // HTTP 1.0 
        Response.AppendHeader("Expires", "Mon, 26 Jul 1997 05:00:00 GMT"); // HTTP 1.0
        Response.Write("<html><head>");
        Response.Write("<meta http-equiv='Refresh' content='0;url=" + longURL + "'></head>");
        Response.Write("</html>");
    }
    public bool IsReusable
    {
        // To enable pooling, return true here.
        // This keeps the handler in memory.
        get { return false; }
    }

    public String checkfile (String lookup){
        String res ="";
        String retval = "";
        String retvalx = "";

        if (lookup =="") {lookup="nomatch";}
        string file = "URL.csv";
        string dir = @"C:\Inetpub\redir";
//      string excelConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + dir + @";Extended Properties=""Text;HDR=Yes;FMT=Delimited""";
        string excelConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + dir + @";Extended Properties=""Text;HDR=No;FMT=Delimited""";
        OleDbConnection conn = new OleDbConnection(excelConn);
        conn.Open();
        string query = "SELECT * FROM [" + file+"] ";
        OleDbCommand cmd = new OleDbCommand(query, conn);
        OleDbDataReader reader = cmd.ExecuteReader();
        while (reader.Read() )
        {
            res=(reader[0].ToString().ToLower () );
            if (lookup == res) {
                    retval= reader[1].ToString();
           //         retvalx=reader[2].ToString();
            }

        }
        conn.Close();
	if (retvalx=="#") {retvalx="";}
        if (retval == "") { retval = "/nomatch"; } 
        return retval ;
    }

}