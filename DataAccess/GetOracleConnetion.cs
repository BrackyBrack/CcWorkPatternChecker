using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Oracle.ManagedDataAccess.Client;


namespace DataAccess
{
    class GetOracleConnetion
    {
        public static OracleConnection GetDBConnection(string host, int port, string database, string username, string password)
        {
            // Connection String.
            OracleConnectionStringBuilder sb = new OracleConnectionStringBuilder();
            sb.DataSource = host + ":" + port + "/" + database;
            sb.UserID = username;
            sb.Password = password;
            OracleConnection conn = new OracleConnection(sb.ToString());

            return conn;
        }

        public static OracleConnection GetDBConnection()
        {
            return GetLiveConnection();
            //return GetTestConnection();
            //return GetTRNConnection();

        }

        private static OracleConnection GetTRNConnection()
        {
            string host = "10.145.52.36";
            int port = 1537;
            string database = "gostrn1";
            string username = "batch";
            string password = "batch123";

            return GetDBConnection(host, port, database, username, password);
        }

        private static OracleConnection GetTestConnection()
        {
            //Test connection settings
            string host = "10.145.128.68";
            int port = 1525;
            string database = "gosuat1";
            string username = "batch";
            string password = "batch123";



            return GetDBConnection(host, port, database, username, password);
        }

        public static OracleConnection GetLiveConnection()
        {
            //Prod connection settings
            string host = "idpsliveaix.britanniaairways.com";
            int port = 1539;
            string database = "goslive.britanniaairways.com";
            string username = "batch";
            string password = "batch5421";



            return GetDBConnection(host, port, database, username, password);
        }
    }
}
