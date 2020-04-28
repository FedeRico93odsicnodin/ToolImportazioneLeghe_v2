using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ToolImportazioneLeghe.Excel
{
    /// <summary>
    /// Tutti i serivizi excel necessari all'analisi
    /// </summary>
    public class ExcelService
    {
        #region ATTRIBUTI PRIVATI

        //private ExcelPackage


        /// <summary>
        /// Servizio excel di lettura 
        /// </summary>
        private ExcelReaders _excelReader;


        /// <summary>
        /// Servizio excel di scrittura 
        /// </summary>
        private ExcelWriters _excelWriter;

        #endregion


        #region COSTRUTTORE

        /// <summary>
        /// Inizializzazione dei servizi di lettura scrittura per excel corrente 
        /// </summary>
        public ExcelService()
        {
            _excelReader = new ExcelReaders();
            _excelWriter = new ExcelWriters();
        }

        #endregion


        #region METODI PUBBLICI 

        /// <summary>
        /// Permette l'apertura del file excel corrente
        /// </summary>
        /// <param name="pathExcel"></param>
        /// <returns></returns>
        public bool OpenFileExcel(string pathExcel)
        {
            return false;
        }

        #endregion
    }


    /// <summary>
    /// Servizi necessari alla lettura di informazioni per il file excel corrente 
    /// </summary>
    public class ExcelReaders
    {
        /// <summary>
        /// Validazione per la prima tipologia di foglio excel 
        /// </summary>
        /// <returns></returns>
        private bool ValidateExcel_Type1()
        {
            return false;
        }
    }


    /// <summary>
    /// Servizi necessari alla scrittura di informazioni per il file excel corrente 
    /// </summary>
    public class ExcelWriters
    {

    }
}
