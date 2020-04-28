using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ToolImportazioneLeghe_Console.Excel.Model_Excel
{
    /// <summary>
    /// Formato per le informazioni correntemente analizzate dalla seconda tipologia per il foglio excel in analisi
    /// </summary>
    public class Excel_Format2_Sheet
    {

        #region ATTRIBUTI PRIVATI

        /// <summary>
        /// Lista di tutte le informazioni di lega lette per la seconda tipologia di formato excel disponibile
        /// </summary>
        private List<Excel_Format2_Row> _listRowLega_Foglio2;


        /// <summary>
        /// Nome foglio corrente
        /// </summary>
        private string _sheetName;


        /// <summary>
        /// Posizione foglio corrente 
        /// </summary>
        private int _posInExcel;

        #endregion


        #region COSTRUTTORE

        /// <summary>
        /// Inizializzazione nome e posizione 
        /// </summary>
        /// <param name="sheetName"></param>
        /// <param name="posInExcel"></param>
        public Excel_Format2_Sheet(string sheetName, int posInExcel)
        {
            _sheetName = sheetName;
            _posInExcel = posInExcel;

            _listRowLega_Foglio2 = new List<Excel_Format2_Row>();
        }

        #endregion


        #region GETTERS

        /// <summary>
        /// Ritorno / aggiorno la lista relativa alle informazioni prelevate per la prima tipologia di foglio per il formato 1
        /// </summary>
        public List<Excel_Format2_Row> GetListInfoLeghe
        {
            get { return _listRowLega_Foglio2; }
            set { _listRowLega_Foglio2 = value; }
        }
        

        /// <summary>
        /// Nome per il foglio corrente 
        /// </summary>
        public string GetSheetName { get { return _sheetName; } }


        /// <summary>
        /// Posizione nel file per il foglio corrente 
        /// </summary>
        public int GetPosSheet { get { return _posInExcel; } }

        #endregion
    }
}
