using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ToolImportazioneLeghe_Console.Messaging_Console;
using ToolImportazioneLeghe_Console.Utils;

namespace ToolImportazioneLeghe_Console.Steps
{
    /// <summary>
    /// In questa classe sono inseriti tutti gli steps per leggere e successivamente inserire le informazioni a database
    /// </summary>
    public static class FromExcelToDatabase
    {
        /// <summary>
        /// STEP 1 di apertura del file excel di riferimento e che è trovato come sorgente 
        /// per l'import corrente 
        /// </summary>
        /// <returns></returns>
        public static bool STEP1_OpenSourceExcel()
        {
            ConsoleService.GetSeparatoreAttivita();

            // segnalazione inizio STEP 1 di apertura excel 
            ConsoleService.STEPS_FromExcelToDatabase.STEP1_InizioAperturaFoglioExcelSorgente(ServiceLocator.GetUtilityFunctions.GetFileName(Constants.ExcelSourcePath));

            if (ServiceLocator.GetExcelService.OpenFileExcel(Constants.ExcelSourcePath, Constants.format_foglio_origin, Constants.ModalitaAperturaExcel.Lettura))
            {
                ConsoleService.STEPS_FromExcelToDatabase.STEP1_FileExcelSorgenteApertoCorrettamente();
                return true;
            }
            else
            {
                throw new Exception(ExceptionMessages.ERRORESTEP1_APERTURAFILEEXCEL);
            }

        }


        /// <summary>
        /// Processo di validazione dei diversi fogli excel presenti all'interno del file in base al formato deciso all'interno del file di configurazioni
        /// </summary>
        /// <returns></returns>
        public static bool STEP2_ValidateSheetOnExcel()
        {
            ConsoleService.GetSeparatoreAttivita();

            // segnalazione inizio STEP 2 di validazione fogli excel
            ConsoleService.STEPS_FromExcelToDatabase.STEP2_InizioValidazioneFogliContenutiInExcel(ServiceLocator.GetUtilityFunctions.GetFileName(Constants.ExcelSourcePath));

            if (ServiceLocator.GetExcelService.GetExcelReaders.RecognizeSheetsOnExcel())
            {
                ConsoleService.STEPS_FromExcelToDatabase.STEP2_FineSorgenteValidatoCorrettamente();
                return true;
            }
            else
                throw new Exception(ExceptionMessages.ERRORESTEP2_VALIDAZIONEFOGLIEXCEL);
        }


        /// <summary>
        /// Processo di recupero di tutte le informazioni dai files in lettura corrente per il file excel
        /// </summary>
        /// <returns></returns>
        public static bool STEP3_GetAllInfoExcel()
        {
            ConsoleService.GetSeparatoreAttivita();

            // segnalazione inizio STEP 3 di recupero informazioni per il file excel corrente 
            ConsoleService.STEPS_FromExcelToDatabase.STEP3_InizioRecuperoDiTutteLeInformazioniPerExcelCorrente(ServiceLocator.GetUtilityFunctions.GetFileName(Constants.ExcelSourcePath));

            if (ServiceLocator.GetExcelService.GetExcelReaders.ReadExcelInformation())
            {
                ConsoleService.STEPS_FromExcelToDatabase.STEP3_RecuperoDelleInformazioniUltimato(ServiceLocator.GetUtilityFunctions.GetFileName(Constants.ExcelSourcePath));
                return true;
            }
            else
                throw new Exception(ExceptionMessages.ERRORESTEP3_LETTURAFOGLIOEXCELNONRIUSCITA);
        }
    }
}
