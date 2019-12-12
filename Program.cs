using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;
using System.Data.OleDb;    // needed to access databases
using System.IO;
using System.Net;           // needed for net communication

namespace MSA2CSV
{
    static class Constants
    {
        // programmnames and version
        public const String programname_short = "MSA2CSV";
        public const String programname_long  = "Microsoft Access to CSV-File";
        public const String programversion    = "0.7";
        public const String programauthor     = "Danilo Ludwig";
        // returncodes;
        public const int rc_OK = 0;
        public const int rc_noArguments = 1;
        public const int rc_noMDBArguments = 2;
        public const int rc_noSQLArguments = 3;
        public const int rc_dbAccessError = 50;
        public const int rc_POSTError = 100;

        public static void show_help(int rc)
        {
            // show programname and Version and help for arguments
            System.Console.WriteLine(Constants.programname_short + " v" + Constants.programversion + " - " + Constants.programname_long + " - (c) 2014 by " + Constants.programauthor);
            switch (rc)
            {
                case rc_noArguments:
                    System.Console.WriteLine("please add arguments");
                    break;
                case rc_noMDBArguments:
                    System.Console.WriteLine("please add argument MDB=");
                    break;
                case rc_noSQLArguments:
                    System.Console.WriteLine("please add argument SQL=");
                    break;
                default:
                    System.Console.WriteLine("something wrong with the arguments");
                    break;
            }
            System.Console.WriteLine("arguments are:");
            System.Console.WriteLine("MDB=\"MDB-filename\"");
            System.Console.WriteLine("MDW=\"MDW-filename\"");
            System.Console.WriteLine("UID=username");
            System.Console.WriteLine("PWD=password");
            System.Console.WriteLine("SQL=\"SELECT ... FROM ...\" or \"filename of file with SQL-statement\"");
            System.Console.WriteLine("printfieldnames");
            System.Console.WriteLine("note: the result is send to stdout");
            System.Console.WriteLine("note: if MDW (system database) is not set, then PWD is used as database password (UID is ignored in this case)");
        }//show_help()

    }//class Constants

    class MDBparams_class
    { //contains the given parameters for the MDB to open
        private String _MDBfilename = "";
        public String MDBfilename { get { return _MDBfilename; } set { _MDBfilename = value; } }
        private String _MDWfilename = "";
        public String MDWfilename { get { return _MDWfilename; } set { _MDWfilename = value; } }
        private String _MDBuid = "";
        public String MDBuid { get { return _MDBuid; } set { _MDBuid = value; } }
        private String _MDBpwd = "";
        public String MDBpwd { get { return _MDBpwd; } set { _MDBpwd = value; } }
        private String _MDBsql = "";
        public String MDBsql { get { return _MDBsql; } set { _MDBsql = value; } }
        private String _POSTto = "";
        public String POSTto { get { return _POSTto; } set { _POSTto = value; } }
        private bool _printFieldNames = false;
        public bool printFieldNames { get { return _printFieldNames; } set { _printFieldNames = value; } }
        private bool _plaintext = false;
        public bool plaintext { get { return _plaintext; } set { _plaintext = value; } }
        private String _CSVdelimiter = ";";
        public String CSVdelimiter { get { return _CSVdelimiter; } set { _CSVdelimiter = value; } }

        public MDBparams_class(string[] parameters) //constructor
        {   //set the parameters from the given program start arguments
            foreach (String param in parameters)
            {   
                String s = param.ToString().Replace("\"", "");
                if (s.Length > 4)
                {
                    switch (s.Substring(0, 4).ToUpper())
                    {
                        case "MDB=": // checking for MDB name
                            this.MDBfilename = s.Substring(4, s.Length - 4);
                            break;
                        case "MDW=": // checking for MDW name
                            this.MDWfilename = s.Substring(4, s.Length - 4);
                            break;
                        case "UID=": // checking for UID
                            this.MDBuid = s.Substring(4, s.Length - 4);
                            break;
                        case "PWD=": // checking for PWD
                            this.MDBpwd = s.Substring(4, s.Length - 4);
                            break;
                        case "SQL=": // checking for SQL
                            s = param.ToString();
                            this.MDBsql = s.Substring(4, s.Length - 4);
                            break;
                        default:
                            break;
                    } //end switch
                    if (s.ToLower().Equals("printfieldnames")) { this.printFieldNames = true; }
                    if (s.ToLower().Equals("plaintext")) { this.plaintext = true; }
                    if (s.Length > 7)
                    {
                        if (s.Substring(0, 6).ToUpper().Equals("POSTTO")) { this.POSTto = s.Substring(7, s.Length - 7); }
                    }
                }//if s.length >= 4
            } //end foreach
            //read the sql statement from sqlfile
            if (File.Exists(MDBsql))
            {
                String sqlline;
                System.IO.StreamReader sqlfile = new System.IO.StreamReader(@MDBsql);
                MDBsql = "";
                while ((sqlline = sqlfile.ReadLine()) != null)
                {
                    MDBsql += " " + sqlline;
                }
                sqlfile.Close();
            }
        } //end constructor MDBparams_class

        public String getConnectionString()
        {   //creates the Connection string out of MDBparams
            //set the connectionstring, see also http://www.connectionstrings.com/access-2003/
            String conns = @"Provider=Microsoft.Jet.OLEDB.4.0;";
            conns = @conns + "Data Source=" + this.@MDBfilename + ";";
            if (!this.MDWfilename.Equals(""))
            { //if MDW (system database) is set, then add uid and pwd as workgroup security
                conns = @conns + @"Jet OLEDB:System Database=" + this.@MDWfilename + ";";
                if (!this.MDBuid.Equals("")) { conns = @conns + @"User ID=" + this.@MDBuid + ";"; };
                if (!this.MDBpwd.Equals("")) { conns = @conns + @"Password=" + this.@MDBpwd +";"; };
            }
            else
            { //if MDW (system database) is not set, then add uid and pwd as database security (and ignore uid if given)
                //if (!MDBparameter.MDBuid.Equals("")) { conns = @conns + @"Jet OLEDB:User ID=" + @MDBparameter.MDBuid + ";"; };
                if (!this.MDBpwd.Equals("")) { conns = @conns + @"Jet OLEDB:Database Password=" + this.@MDBpwd +";"; };
            }//else (!MDBparameter.MDWfilename.Equals(""))
            return @conns;
        }//getConnectionString()
        
    } //end class MDBparams_class



    class Program //program starts here
    {
        static private String getCSVvalue(Object DBvalue, bool plaintext)
        {// transforms the database value into a string value for the csv output

            if (DBvalue.GetType().ToString().Equals("System.String"))
            {
                if (plaintext)
                    { return DBvalue.ToString(); }
                else
                    { return '"' + DBvalue.ToString() + '"'; }
            }
            else
            {
                return @DBvalue.ToString();//.Replace(",",".");
            }
        } //getCSVvalue
        
        static int Main(string[] args)
        {
            String s = "";
            String POSTstring = "";
            int i = 0;
            // check, if there are arguments
            if (args.Length == 0)
            {
                Constants.show_help(Constants.rc_noArguments); // show programname and Version and help for arguments
                return Constants.rc_noArguments; // end the program with returncode for missing arguments
            }
            else
            {
                // here is the real start
                // get the arguments and store in Object MDBparams
                MDBparams_class MDBparams = new MDBparams_class(args);
                
                // check if at least MDB-filename and SQL-statement is given
                if (MDBparams.MDBfilename.Equals(""))
                {
                    Constants.show_help(Constants.rc_noMDBArguments);
                    return Constants.rc_noMDBArguments;
                }
                if (MDBparams.MDBsql.Equals(""))
                {
                    Constants.show_help(Constants.rc_noSQLArguments);
                    return Constants.rc_noSQLArguments;
                }
                // now try to open the MDB Databasefile
                try
                {
                    OleDbConnection con = new OleDbConnection();
                    con.ConnectionString = MDBparams.getConnectionString();
                    //open DB
                    con.Open();
                    //execute the SQL statement
                    OleDbCommand comm = new OleDbCommand();
                    comm.Connection = con;
                    comm.CommandText = MDBparams.MDBsql;
                    OleDbDataReader reader = comm.ExecuteReader();
                    //print the column names on stdout
                    if (MDBparams.printFieldNames)
                    {
                        s = "";
                        for (i = 0; i < reader.FieldCount; i++)
                        {
                            if (!s.Equals("")) { s = @s + MDBparams.CSVdelimiter; }
                            s = @s + @reader.GetName(i);
                        }
                        System.Console.WriteLine(s);
                    }
                    //read the result (data) and print it on stdout
                    while (reader.Read())
                    {
                        s = "";
                        for (i = 0; i < reader.FieldCount; i++)
                        {
                            if (!s.Equals("")) { s = @s + MDBparams.CSVdelimiter; }
                            s += @getCSVvalue(reader.GetValue(i),MDBparams.plaintext);
                        }
                        System.Console.WriteLine(s);
                        if (!MDBparams.POSTto.Equals(""))
                        {
                            POSTstring = POSTstring + @s + "\r\n";
                        }
                    }
                    //close DB
                    con.Close();
                    //post the result to POSTto URI
                    if (!MDBparams.POSTto.Equals(""))
                    {
                        //string URI = "http://www.myurl.com/post.php";
                        //string myParameters = "param1=value1&param2=value2&param3=value3";
                        using (WebClient client = new WebClient())
                        {
                            // Optionally specify an encoding for uploading and downloading strings
                            client.Encoding = System.Text.Encoding.UTF8;
                            // set the HTTPRequestHeader
                            client.Headers[HttpRequestHeader.ContentType] = "application/x-www-form-urlencoded";
                            try
                            {
                                // Upload the data
                                string HtmlResult = client.UploadString(MDBparams.POSTto, POSTstring);
                                // Display the server's response
                                Console.WriteLine(HtmlResult);
                            }
                            catch (Exception e)
                            {
                                System.Console.WriteLine(e.Message);
                                return Constants.rc_POSTError;
                            }
                        } // using (WebClient...)
                    } // if POSTto is set
                } // try read database
                catch (Exception e)
                {
                    // if somethings went wrong show the error message
                    System.Console.WriteLine(e.Message);
                    return Constants.rc_dbAccessError;
                }
                
            }//else
            // end programm with returncode for OK
            return Constants.rc_OK;
        }//main
    }//class Program
}//namespace MSA2CSV
