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
        /// lista di tutte le informazioni di lega che vengono lette dalla prima tipologia di foglio per il formato 1
        /// </summary>
        private List<Excel_Format1_Sheet1_Row> _listInfoLega_FoglioType1;


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


            if (tipologiaFoglioFormato1 == Constants_Excel.TipologiaFoglio_Format1.FoglioLeghe)
                _listInfoLega_FoglioType1 = new List<Excel_Format1_Sheet1_Row>();
            else if (tipologiaFoglioFormato1 == Constants_Excel.TipologiaFoglio_Format1.FoglioConcentrazioni)
                _listConcQuadrants_FoglioType2 = new List<Excel_Format1_Sheet2_ConcQuadrant>();
        }

        #endregion


        #region GETTERS

        /// <summary>
        /// Ritorno / aggiorno la lista relativa alle informazioni prelevate per la prima tipologia di foglio per il formato 1
        /// </summary>
        public List<Excel_Format1_Sheet1_Row> GetListInfoLega_Type1
        {
            get { return _listInfoLega_FoglioType1; }
            set { _listInfoLega_FoglioType1 = value; }
        }


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

        #endregion

    }
}
