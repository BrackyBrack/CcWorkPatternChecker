using DataAccess;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CcWorkPatternChecker
{
    class Program
    {
        static void Main(string[] args)
        {
            DataTable crewPanels = CsvReader.GetCrewPanels();
            DataTable jsPanels = CsvReader.GetPanels();
            DataTable plannedJs = CrewData.GetOffPeriods();

            StringBuilder sb = new StringBuilder();
            foreach (DataRow crewRow in crewPanels.Rows)
            {
                string partTimeType = crewRow["Part Time Type"].ToString();

                if(partTimeType != "FT")
                {
                    string crewMember = crewRow[0].ToString();

                    var crewPlannedJs = plannedJs.Select($"P_LTR_CODE = '{crewMember}'");

                    List<DateTime> plannedDates = GetPlannedDates(crewPlannedJs);

                    foreach (DataRow jsRow in jsPanels.Rows)
                    {
                        if (jsRow[partTimeType].ToString() == "1")
                        {
                            DateTime panelDate = DateTime.Parse(jsRow[0].ToString());
                            if (plannedDates.Contains(panelDate) == false)
                            {
                                sb.AppendLine($"{crewMember} does not match panel {partTimeType}");
                                break;
                            }
                        }
                    }
                }
                else
                {
                    string crewMember = crewRow[0].ToString();

                    var crewPlannedJs = plannedJs.Select($"P_LTR_CODE = '{crewMember}'");

                    if(crewPlannedJs.Count() > 0)
                    {
                        sb.AppendLine($"{crewMember} does not match panel {partTimeType}");
                        break;
                    }
                }
            }

            System.IO.File.WriteAllText("C:\\Users\\david.bracken\\OneDrive - TUI\\Documents\\Furlough\\CC Infor\\Panel Errors.txt", sb.ToString());
        }

        private static List<DateTime> GetPlannedDates(DataRow[] crewPlannedJs)
        {
            List<DateTime> result = new List<DateTime>();

            foreach (DataRow row in crewPlannedJs)
            {
                DateTime date = DateTime.Parse(row[1].ToString());
                result.Add(date);
            }

            return result;
        }
    }
}
