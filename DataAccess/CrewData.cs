using Oracle.ManagedDataAccess.Client;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataAccess
{
    public class CrewData
    {
        public static DataTable GetPersNr()
        {
            string query = "SELECT * FROM CREW_ENDING_CONTRACT_FULL";
            DataTable dataTable = new DataTable();
            using (OracleConnection connection = GetOracleConnetion.GetLiveConnection())
            {
                OracleCommand command = new OracleCommand(query, connection);

                OracleDataAdapter dataAdapter = new OracleDataAdapter();
                dataAdapter.SelectCommand = command;

                connection.Open();
                dataAdapter.Fill(dataTable);
                connection.Close();
            }

            return dataTable;
        }

        public static DataTable GetOffPeriods()
        {
            string query = "SELECT * FROM JS_PANELS ";
            DataTable dataTable = new DataTable();
            using (OracleConnection connection = GetOracleConnetion.GetLiveConnection())
            {
                OracleCommand command = new OracleCommand(query, connection);

                OracleDataAdapter dataAdapter = new OracleDataAdapter();
                dataAdapter.SelectCommand = command;

                connection.Open();
                dataAdapter.Fill(dataTable);
                connection.Close();
            }

            return dataTable;
        }
    }
}