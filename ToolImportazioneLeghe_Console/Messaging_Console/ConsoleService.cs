﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ToolImportazioneLeghe_Console.Utils;

namespace ToolImportazioneLeghe_Console.Messaging_Console
{
    public static class ConsoleService
    {
        #region TUTTI I MESSAGGI PER LA CONSOLE 

        /// <summary>
        /// Servizio corrente per i messaggi di console provenienti dal servizio di logging
        /// </summary>
        private static ConsoleService_UtilityFunctionsMessages _consoleService_LoggingMessage; 


        /// <summary>
        /// Servizio corrente per i messaggi di console provenienti dal servizio excel
        /// </summary>
        private static ConsoleService_Excel _consoleService_Excel;

        #endregion


        #region METODI PER L'EFFETTIVO UTILIZZO DELLA CONSOLE 

        /// <summary>
        /// Permette la formattazione per il messaggio visualizzato in console e il suo log nel file 
        /// di log globale (nel caso sia specificato come parametro di input)
        /// </summary>
        /// <param name="currentMessage"></param>
        /// <param name="logMessage"></param>
        public static void FormatMessageConsole(string currentMessage, bool logMessage)
        {
            if (logMessage)
                ServiceLocator.GetLoggingService.LogMessage(currentMessage, Constants.GlobalLoggingPath);
        }

        #endregion


        #region TUTTI I SERVIZI DI CONSOLE

        public static ConsoleService_UtilityFunctionsMessages ConsoleLogging
        {
            get
            {
                if (_consoleService_LoggingMessage == null)
                    _consoleService_LoggingMessage = new ConsoleService_UtilityFunctionsMessages();

                return _consoleService_LoggingMessage;
            }
        }


        /// <summary>
        /// Getter per messaging console excel services
        /// </summary>
        public static ConsoleService_Excel ConsoleExcel
        {
            get
            {
                if (_consoleService_Excel == null)
                    _consoleService_Excel = new ConsoleService_Excel();

                return _consoleService_Excel;
            }
        }

        #endregion
    }


    /// <summary>
    /// Tutti i serivizi di messaging console per il servizio di logging
    /// </summary>
    public class ConsoleService_UtilityFunctionsMessages
    {
        /// <summary>
        /// Marker per i messaggi provienti dal servizio di log 
        /// </summary>
        private const string utilityFunctions_Marker = "UTILITY: ";


        /// <summary>
        /// Segnalazione di esistenza per il file di log attuale, la procedura verrà accodata 
        /// a questo file di log 
        /// </summary>
        /// <param name="logFilePath"></param>
        public void EsistenzaFileLog_Message(string logFilePath)
        {
            string currentMessage = String.Format(utilityFunctions_Marker + "il file di LOG al path \n'{0}'\nesiste già, la procedura verrà loggata in accodamento.", logFilePath);
            ConsoleService.FormatMessageConsole(currentMessage, false);
        }


        /// <summary>
        /// Segnalazione di ricreazione della folder per inserire successivamente il log nel percorso 
        /// esplicitato
        /// </summary>
        /// <param name="analyzedFolder"></param>
        /// <param name="targetLogFile"></param>
        public void RicreazioneFolder(string analyzedFolder, string targetLogFile)
        {
            string currentMessage = String.Format(utilityFunctions_Marker + "ho appena creato la seguente cartella '{0}' per l'inserimneto del log '{1}'", analyzedFolder, targetLogFile);
            ConsoleService.FormatMessageConsole(currentMessage, false);
        }


        /// <summary>
        /// Segnalazione di una ricreazione del file nel quale verranno loggate parti della procedura 
        /// </summary>
        /// <param name="targetLogFile"></param>
        public void RicreazioneLogFile(string targetLogFile)
        {
            string currentMessage = String.Format(utilityFunctions_Marker + "ho appena ricreato il seguente file di log '{0}', la procedura verrà loggata da 0 in questo file", targetLogFile);
            ConsoleService.FormatMessageConsole(currentMessage, false);
        }
    }


    /// <summary>
    /// Tutti i servizi di messaging console per il servizio excel 
    /// </summary>
    public class ConsoleService_Excel
    {
        #region ATTRIBUTI CREAZIONE INTESTAZIONE MESSAGGIO

        /// <summary>
        /// Excel per il marker principale per i servizi excel
        /// </summary>
        private const string excelService_Marker = "EXCEL - ";


        /// <summary>
        /// Marker per i reader di inserimento excel 
        /// </summary>
        private const string readerExcel_Marker = "READER: ";


        /// <summary>
        /// Marker per identificare il formato 1 proveniente dal database di leghe 
        /// </summary>
        private const string intestazioneFormat1 = "FORMAT 1 (DATABASE LEGHE) - ";


        /// <summary>
        /// Marker per identificare il formato 2 proveninete dal cliente 
        /// </summary>
        private const string intestazioneFormat2 = "FORMAT 2 (CLIENTE) - ";

        #endregion


        /// <summary>
        /// Segnalazione che il foglio è stato trovato di una certa tipologia per l'istanza di reader corrente e per il formato corrente
        /// questo metodo è riferito al primo formato proveniente dal database di tutte le leghe
        /// </summary>
        /// <param name="currentFoglio"></param>
        /// <param name="currentPosizione"></param>
        /// <param name="currentTipologia"></param>
        public void ExcelReaders_Message_RiconoscimentoSeguenteTipologia_Format1(string currentFoglio, int currentPosizione, string currentTipologia)
        {
            string currentMessage = String.Format(excelService_Marker + intestazioneFormat1 + readerExcel_Marker + "il foglio di nome '{0}' in posizione in excel {1} è stato riconosciuto come {2}", currentFoglio, currentPosizione, currentTipologia);
            ConsoleService.FormatMessageConsole(currentMessage, true);
        }


        /// <summary>
        /// Segnalazione che il foglio è stato trovato come contenente delle informazioni valide per la lettura di leghe e concentrazioni 
        /// inerentemente la seconda tipologia di foglio
        /// </summary>
        /// <param name="currentFoglio"></param>
        /// <param name="currentPosizione"></param>
        public void ExcelReaders_Message_RiconoscimentoSeguenteTipologia_Format2(string currentFoglio, int currentPosizione)
        {
            string currentMessage = String.Format(excelService_Marker + intestazioneFormat2 + readerExcel_Marker + "il foglio di nome '{0}' in posizione excel {1} è stato riconosciuto come contenere delle informazioni valide", currentFoglio, currentPosizione);
            ConsoleService.FormatMessageConsole(currentMessage, true);
        }


        /// <summary>
        /// Segnalazione che il foglio non è stato riconosciuto come contenere delle informazioni valide per il file corrente e per le informazioni 
        /// di lega o di concentrazione
        /// </summary>
        /// <param name="currentFoglio"></param>
        /// <param name="currentPosizione"></param>
        public void ExcelReaders_Message_FoglioNonRiconosciuto(string currentFoglio, int currentPosizione)
        {
            string currentMessage = String.Format(excelService_Marker + readerExcel_Marker + "il seguente foglio '{0}' in posizione {1} non è stato riconosciuto come foglio di informazioni valide, si prega di controllarne nuovamente il contenuto", currentFoglio, currentPosizione);
            ConsoleService.FormatMessageConsole(currentMessage, true);
        }





    }
}
