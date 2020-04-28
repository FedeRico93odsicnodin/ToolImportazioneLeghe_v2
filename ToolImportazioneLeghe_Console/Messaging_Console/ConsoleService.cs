using System;
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
            string currentMessage = String.Format(utilityFunctions_Marker + "ho appena creato la seguente cartella '{0}' per l'inserimneto del log '{1}'");
            ConsoleService.FormatMessageConsole(currentMessage, false);
        }


        /// <summary>
        /// Segnalazione di una ricreazione del file nel quale verranno loggate parti della procedura 
        /// </summary>
        /// <param name="targetLogFile"></param>
        public void RicreazioneLogFile(string targetLogFile)
        {
            string currentMessage = String.Format(utilityFunctions_Marker + "ho appena ricreato il seguente file di log '{0}', la procedura verrà loggata da 0 in questo file");
            ConsoleService.FormatMessageConsole(currentMessage, false);
        }
    }


    /// <summary>
    /// Tutti i servizi di messaging console per il servizio excel 
    /// </summary>
    public class ConsoleService_Excel
    {
        
    }
}
