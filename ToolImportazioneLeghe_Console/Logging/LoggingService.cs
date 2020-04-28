using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ToolImportazioneLeghe_Console.Messaging_Console;
using ToolImportazioneLeghe_Console.Utils;

namespace ToolImportazioneLeghe_Console.Logging
{
    /// <summary>
    /// Servizi per il log all'interno dei diversi files di log 
    /// </summary>
    public class LoggingService
    {
        #region COSTRUTTORE - INIZIALIZZAZIONE DEL SERVIZIO DI LOG

        public LoggingService()
        {
            // ricreazione del log globale 
            BuildLogFilePath(Constants.GlobalLoggingPath);
        }

        #endregion

        /// <summary>
        /// Permette la ricostruzione del file di log nel caso in cui non esista nel percorso 
        /// passato come input nelle configurazioni
        /// </summary>
        /// <param name="logFilePath"></param>
        private void BuildLogFilePath(string logFilePath)
        {
            ServiceLocator.GetUtilityFunctions.BuildFilePath(logFilePath);
        }


        /// <summary>
        /// Servizio di log in base a stringa e path del percorso per il file passato in input
        /// </summary>
        /// <param name="lines"></param>
        /// <param name="logPath"></param>
        public void LogMessage(string lines, string logPath)
        {
            if (logPath == Constants.GlobalLoggingPath)
            {
                System.IO.StreamWriter file = new System.IO.StreamWriter(logPath, true);
                file.WriteLine(lines);

                file.Close();
            }
            else
                throw new Exception(String.Format(ExceptionMessages.LOGGING_INVALIDLOGPATH, logPath));
        }
    }


    
}
