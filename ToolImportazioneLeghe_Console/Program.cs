using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ToolImportazioneLeghe_Console.Utils;

namespace ToolImportazioneLeghe_Console
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                ServiceLocator.GetExcelService.OpenFileExcel(Constants.ExcelSourcePath, Constants.format_foglio_origin, Constants.ModalitaAperturaExcel.Lettura);
            }
            catch(Exception e)
            {

            }
        }
    }
}
