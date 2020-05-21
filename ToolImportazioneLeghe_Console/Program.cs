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


        /// <summary>
        /// Indica la conclusione dello step relativo al recupero delle informazioni per la prima fase di validazione per i diversi fogli excel 
        /// in questa fase la validazione non passa se per il set delle informazioni di base obbligatorie per poter proseguire non sono stati letti 
        /// dei valori (le celle sono null)
        /// </summary>
        private static bool _conclusioneSTEP3_Lettura1InformazioniExcel = false;

        #endregion


        static void Main(string[] args)
        {
            try
            {
                // STEP 1: apertura del file excel corrente per la successiva validazione 
                _conclusioneSTEP1_ExcelToDatabase = FromExcelToDatabase.STEP1_OpenSourceExcel();

                // STEP 2: validazione delle diverse tipologie di foglio contenute all'interno del file excel aperto nello step precedente 
                if (_conclusioneSTEP1_ExcelToDatabase)
                    _conclusioneSTEP2_ValidazioneFogliExcel = FromExcelToDatabase.STEP2_ValidateSheetOnExcel();

                // STEP 3: recupero informazioni e validazione 1 per le informazioni contenute all'interno del file excel aperto e validato sugli steps precedenti
                if (_conclusioneSTEP2_ValidazioneFogliExcel)
                    _conclusioneSTEP3_Lettura1InformazioniExcel = FromExcelToDatabase.STEP3_GetAllInfoExcel();
            }
            catch(Exception e)
            {

            }
        }
    }
}
