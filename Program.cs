using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace TeamStats
{
    class Program
    {
        static void Main(string[] args)
        {
            string path = "E:\\Code\\VSCode\\Node\\CFB01\\2019CFPickem - Copy.xlsx";
            Application excel = new Application();
            Workbook wb = excel.Workbooks.Open(path);
            try
            {
                //Get data from the Offensive Stats worksheet
                Console.WriteLine("Transfer offensive stats to team sheets");
                Worksheet wsOff = wb.Worksheets["Offense"];

                int i = 2;
                while (wsOff.Cells[i, 2].Value != "N")
                {
                    Worksheet wsTeam = wb.Worksheets[wsOff.Cells[i, 2].Value];
                    //Copy values from team stats row 5 to row 6
                    wsTeam.Cells[6, 1].Value = wsTeam.Cells[5, 1].Value;
                    wsTeam.Cells[6, 2].Value = wsTeam.Cells[5, 2].Value;
                    wsTeam.Cells[6, 4].Value = wsTeam.Cells[5, 4].Value;
                    wsTeam.Cells[6, 6].Value = wsTeam.Cells[5, 6].Value;
                    wsTeam.Cells[6, 8].Value = wsTeam.Cells[5, 8].Value;
                    wsTeam.Cells[6, 10].Value = wsTeam.Cells[5, 10].Value;
                    wsTeam.Cells[6, 12].Value = wsTeam.Cells[5, 12].Value;
                    wsTeam.Cells[6, 14].Value = wsTeam.Cells[5, 14].Value;
                    wsTeam.Cells[6, 16].Value = wsTeam.Cells[5, 16].Value;
                    //Copy values from Offensive stats worksheet to team stats worksheet
                    wsTeam.Cells[5, 2].Value = wsOff.Cells[i, 3].Value;
                    wsTeam.Cells[5, 4].Value = wsOff.Cells[i, 4].Value;
                    wsTeam.Cells[5, 6].Value = wsOff.Cells[i, 5].Value;
                    wsTeam.Cells[5, 8].Value = wsOff.Cells[i, 6].Value;
                    //Determine if team had an off week
                    if (wsTeam.Cells[7, 4].Value != 0)
                    {
                        wsTeam.Cells[5, 1].Value = wsTeam.Cells[6, 1].Value + 1;
                    }
                }

                //Get data from Defensive Stats worksheet
                Console.WriteLine("Transfer defensive stats to team sheets");
                Worksheet wsDef = wb.Worksheets["Defense"];

                int j = 2;
                while (wsDef.Cells[j, 2].Value != "N")
                {
                    Worksheet wsTeam = wb.Worksheets[wsDef.Cells[j, 2].Value];
                    //Copy values from Defensive stats worksheet to team stats worksheet
                    wsTeam.Cells[5, 10].Value = wsDef.Cells[j, 3].Value;
                    wsTeam.Cells[5, 12].Value = wsDef.Cells[j, 4].Value;
                    wsTeam.Cells[5, 14].Value = wsDef.Cells[j, 5].Value;
                    wsTeam.Cells[5, 16].Value = wsDef.Cells[j, 6].Value;
                }
                wb.Save();
                excel.Quit();
            }
            catch (Exception)
            {
                excel.Quit();
                throw;
            }
            finally
            {
                excel.Quit();
            }
        }
    }
}
