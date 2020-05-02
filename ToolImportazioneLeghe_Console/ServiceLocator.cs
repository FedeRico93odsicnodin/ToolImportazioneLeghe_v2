using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ToolImportazioneLeghe_Console.Excel;
using ToolImportazioneLeghe_Console.Logging;
using ToolImportazioneLeghe_Console.Utils;

namespace ToolImportazioneLeghe_Console
{
    /// <summary>
    /// Locazione dei servizi del tool
    /// </summary>
    public static class ServiceLocator
    {
        #region ATTRIBUTI PRIVATI - TUTTI I SERVIZI

        /// <summary>
        /// Funzioni di utilità globale per il tool
        /// </summary>
        private static UtilityFunctions _utilityFunctions;


        /// <summary>
        /// Servizio di log per il tool
        /// </summary>
        private static LoggingService _loggingService;


        /// <summary>
        /// Serizio corrente excel 
        /// </summary>
        private static ExcelService _excelService;

        #endregion


        #region PROPRIETA PUBBLICHE - GETTERS SERVIZI

        /// <summary>
        /// Servizio di getters per le funzionalita comuni a tutto il tool
        /// </summary>
        public static UtilityFunctions GetUtilityFunctions
        {
            get
            {
                if (_utilityFunctions == null)
                    _utilityFunctions = new UtilityFunctions();

                return _utilityFunctions;
            }
        }


        /// <summary>
        /// Servizio di log per il tool corrente 
        /// </summary>
        public static LoggingService GetLoggingService
        {
            get
            {
                if (_loggingService == null)
                    _loggingService = new LoggingService();

                return _loggingService;
            }
        }


        public static ExcelService GetExcelService
        {
            get
            {
                if (_excelService == null)
                    _excelService = new ExcelService();

                return _excelService;
            }
        }

        #endregion

    }
}
