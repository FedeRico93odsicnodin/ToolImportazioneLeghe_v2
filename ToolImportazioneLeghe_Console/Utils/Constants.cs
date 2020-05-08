using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ToolImportazioneLeghe_Console.Utils
{
    /// <summary>
    /// Costanti di configurazione per l'istanza di tool corrente 
    /// </summary>
    public static class Constants
    {
        #region EXCEL PARAMS

        /// <summary>
        /// Indicazione del fatto che il file excel in apertura venga utilizzato in modalità di lettura o 
        /// scrittura 
        /// </summary>
        public enum ModalitaAperturaExcel
        {
            Lettura = 1,
            Scrittura = 2
        }


        /// <summary>
        /// Indicazione di quale sia il format per il file excel in apertura corrente
        /// se corrisponde allo standard adottato da cliente o per la lettura delle informazioni 
        /// dal database delle leghe 
        /// </summary>
        public enum FormatFileExcel
        {
            NotDefined = 0,
            Cliente = 1,
            DatabaseLeghe = 2
        }


        /// <summary>
        /// Fomrat per il file excel utilizzato come eventuale sorgente 
        /// </summary>
        public static FormatFileExcel format_foglio_origin = FormatFileExcel.Cliente;


        /// <summary>
        /// Format per il file excel utilizzato come eventuale destinazione
        /// </summary>
        public static FormatFileExcel format_foglio_destination = FormatFileExcel.NotDefined;


        /// <summary>
        /// Path per l'eventuale file excel di origine 
        /// </summary>
        public static string ExcelSourcePath = "C:\\Users\\Fede\\Desktop\\Alloy_test.xlsx";


        /// <summary>
        /// Path per l'eventuale file excel di destinazione 
        /// </summary>
        public static string ExcelDestinationPath = String.Empty;

        #endregion


        #region GLOBAL

        /// <summary>
        /// Stringa di logging per il log globale sull'intera 
        /// procedura di importazione 
        /// </summary>
        public static string GlobalLoggingPath = "C:\\Users\\Fede\\Desktop\\log.txt";

        
        /// <summary>
        /// Modalita di inserimento delle informazioni nella destinazione 
        /// in particolare se si trovano delle corrispondenze, si puo decidere di:
        /// 1) sovrascrivere tutto il set delle informazioni presenti nella destinazione con quelle di cui si è deciso l'inserimento 
        /// 2) inserire in accodamento queste informazioni (inserimento delle informazioni completando il set che era già presente)
        /// 3) fermare il tool se le informazioni in inserimento non sono completamente nuove per la destinazione 
        /// </summary>
        public enum ModalitaInserimentoInformazioni
        {
            Stop = 1,
            Sovrascrittura = 2,
            Accodamento = 3
        }


        /// <summary>
        /// Modalita di inserimento delle informazioni letta dalle configurazioni 
        /// </summary>
        public static ModalitaInserimentoInformazioni CurrentModalitaInserimentoInformazioni = ModalitaInserimentoInformazioni.Stop;

        #endregion
    }
}
