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


        }

        #endregion


        #region GETTERS
        
        /// <summary>
        /// Nome per il foglio corrente 
        /// </summary>
        public string GetSheetName { get { return _sheetName; } }


        /// <summary>
        /// Posizione nel file per il foglio corrente 
        /// </summary>
        public int GetPosSheet { get { return _posInExcel; } }


        /// <summary>
        /// Informazione riga di inizio per la lettura delle concentrazioni
        /// </summary>
        public int StartingRow_Leghe { get; set; }


        /// <summary>
        /// Informazione riga di fine lettura per le informazioni generali di lega 
        /// </summary>
        public int EndingCol_Leghe { get; set; }


        /// <summary>
        /// Informazione colonna di inizio per la lettura delle concentrazioni
        /// </summary>
        public int StartingCol_Leghe { get; set; }


        /// <summary>
        /// Informazioni di carattere generale prese in lettura per le informazioni di lega di riga corrente
        /// in questa lista sono contenute tutte le informazioni per le proprieta obbligatorie / opzionali per le informazioni
        /// generali di lega 
        /// </summary>
        public List<Excel_Format2_Row_LegaProperties> AllInfoLeghe { get; set; }


        /// <summary>
        /// Indica se il foglio corrente ha passato la validazione finale inerente il recupero di tutte le informazioni contenute al suo interno
        /// cio significa che per almeno una riga di lega, per questo iniziale controllo, sia stato letto tutto il set di proprieta e concentrazioni 
        /// da eventualmente trasferire
        /// </summary>
        public bool RecuperoCorrettoInformazioni { get; set; }

        #endregion
    }
}
