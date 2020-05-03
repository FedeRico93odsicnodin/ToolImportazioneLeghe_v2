using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ToolImportazioneLeghe_Console.Steps;
using ToolImportazioneLeghe_Console.Utils;

namespace ToolImportazioneLeghe_Console
{
    class Program
    {
        #region STEPS IMPORT EXCEL TO DATABASE

        /// <summary>
        /// Indica se si è conclusa la fase di validazione per l'excel corrente e posso passare alla lettura delle informazioni
        /// contenute nel file excel 
        /// </summary>
        private static bool _conclusioneSTEP1_ExcelToDatabase = false;


        /// <summary>
        /// Indica il concludersi dello STEP 2 per la validazione dei fogli sul file excel corrente
        /// </summary>
        private static bool _conclusioneSTEP2_ValidazioneFogliExcel = false;

        #endregion


        static void Main(string[] args)
        {
            try
            {
                // apertura del file excel corrente per la successiva validazione 
                _conclusioneSTEP1_ExcelToDatabase = FromExcelToDatabase.STEP1_OpenSourceExcel();

                if (_conclusioneSTEP1_ExcelToDatabase)
                    _conclusioneSTEP2_ValidazioneFogliExcel = FromExcelToDatabase.STEP2_ValidateSheetOnExcel();
            }
            catch(Exception e)
            {

            }
        }
    }
}
