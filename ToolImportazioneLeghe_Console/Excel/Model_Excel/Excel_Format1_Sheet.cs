using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ToolImportazioneLeghe_Console.Excel.Model_Excel
{
    /// <summary>
    /// Tipologia di foglio correntemente in analisi per il caso in cui il foglio excel preso sia relativo al formato 1 di lettura delle informazioni 
    /// dal database originale di leghe 
    /// </summary>
    public class Excel_Format1_Sheet
    {
        #region ATTRIBUTI PRIVATI
        
        /// <summary>
        /// lista di tutte le informazioni relative alle concentrazioni in lettura per la seconda tipologia di foglio e il formato 1
        /// </summary>
        private List<Excel_Format1_Sheet2_ConcQuadrant> _listConcQuadrants_FoglioType2;


        /// <summary>
        /// Nome per il foglio correntemente in analisi
        /// </summary>
        private string _sheetName;


        /// <summary>
        /// Tipologia letta per il foglio corrente rispetto al primo formato disponibile
        /// </summary>
        private Constants_Excel.TipologiaFoglio_Format1 _tipologiaFoglioCorrente;


        /// <summary>
        /// Posizione del foglio corrente nel file di origine
        /// </summary>
        private int _posInExcel;

        #endregion

        
        #region COSTRUTTORE 

        /// <summary>
        /// Inizializzazione con tutti i parametri letti per il foglio in analisi
        /// </summary>
        /// <param name="sheetName"></param>
        /// <param name="tipologiaFoglioFormato1"></param>
        /// <param name="posInExcel"></param>
        public Excel_Format1_Sheet(string sheetName, Constants_Excel.TipologiaFoglio_Format1 tipologiaFoglioFormato1, int posInExcel)
        {
            _sheetName = sheetName;
            _tipologiaFoglioCorrente = tipologiaFoglioFormato1;
            _posInExcel = posInExcel;

            
        }

        #endregion


        #region GETTERS

        /// <summary>
        /// Ritorno / aggiorno l'oggetto contenente tutte le proprieta lette per le leghe dalla prima tipologia di foglio per il 
        /// primo formato excel 
        /// </summary>
        public Excel_Format1_Sheet1_Rows GetListInfoLega_Type1 { get; set; }


        /// <summary>
        /// Ritorno / aggiorno la lista relativa alle informazioni di concentrazione lette per la seconda tipologia di foglio
        /// </summary>
        public List<Excel_Format1_Sheet2_ConcQuadrant> GetConcQuadrants_Type2
        {
            get { return _listConcQuadrants_FoglioType2; }
            set { _listConcQuadrants_FoglioType2 = value; }
        }


        /// <summary>
        /// Nome per il foglio corrente 
        /// </summary>
        public string GetSheetName { get { return _sheetName; } }


        /// <summary>
        /// Tipologia per il foglio corrente 
        /// </summary>
        public Constants_Excel.TipologiaFoglio_Format1 GetTipologiaFoglio { get { return _tipologiaFoglioCorrente; } }


        /// <summary>
        /// Posizione nel file per il foglio corrente 
        /// </summary>
        public int GetPosSheet { get { return _posInExcel; } }


        /// <summary>
        /// Posizione iniziale di riga per la lettura delle leghe, nel caso in cui si trattasse 
        /// di un foglio per la lettura delle informazioni principali di lega 
        /// </summary>
        public int StartingRow_letturaLeghe { get; set; }


        /// <summary>
        /// Posizione iniziale di colonna per la lettura delle leghe, nel caso in cui si trattasse
        /// di un foglio per la lettura delle informazioni principali di lega
        /// </summary>
        public int StartingCol_letturaLeghe { get; set; }


        /// <summary>
        /// Posizione di fine lettura colonne per le proprieta di lega corrente
        /// </summary>
        public int EndingCol_letturaLeghe { get; set; }

        #endregion


        #region STRINGHE MESSAGGISTICA IN USCITA 

        /// <summary>
        /// Indicazione degli errori che sono emersi durante la lettura e l'analisi del file excel
        /// nei diversi steps di read
        /// questa stringa è per il tentativo di recupero delle informazioni relative alla lega 
        /// </summary>
        public string ErrorMessages_ExcelAnalyzer_ProprietaLeghe { get; set; }


        /// <summary>
        /// indicazione dei messaggi di warnings emersi durante la lettura e l'analisi del file excel 
        /// nei diversi steps di read
        /// questa stringa è per il tentativo di recupero delle informazioni relative alla lega 
        /// </summary>
        public string WarningMessages_ExcelAnalyzer_ProprietaLeghe { get; set; }


        /// <summary>
        /// indicazione dei messaggi di errore emersi durante la lettura e l'analisi del file excel 
        /// nei diversi steps di read
        /// questa stringa è per il tentativo di recupero delle informazioni relative alle concentrazioni
        /// </summary>
        public string ErrorMessages_ExcelAnalyzer_ProprietaConcentrazioni { get; set; }


        /// <summary>
        /// Indicazione dei messaggi di warnings emersi durante la lettura e l'analisi del file excel 
        /// nei diversi steps di read
        /// questa stringa è per il tentativo di recupero delle informazioni relative alle concentrazioni
        /// </summary>
        public string WarningMessages_ExcelAnalyzer_ProprietaConcentrazioni { get; set; }

        #endregion

    }
}
