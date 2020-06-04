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
    public class Excel_AlloyInfo_Sheet
    {
        

        #region COSTRUTTORE

        /// <summary>
        /// Inizializzazione nome e posizione 
        /// </summary>
        /// <param name="sheetName"></param>
        /// <param name="posInExcel"></param>
        public Excel_AlloyInfo_Sheet(string sheetName, int posInExcel)
        {
            GetSheetName = sheetName;
            GetPosSheet = posInExcel;
        }

        #endregion


        #region PROPRIETA PUBBLICHE
        
        /// <summary>
        /// Nome per il foglio corrente 
        /// </summary>
        public string GetSheetName { get; }


        /// <summary>
        /// Posizione nel file per il foglio corrente 
        /// </summary>
        public int GetPosSheet { get; }


        /// <summary>
        /// Tipologia per il foglio in lettura corrente, puo essere relativo a un foglio di sole leghe e concentrazioni (caso 1)
        /// oppure a un foglio con informazioni sia di leghe che di concentrazioni come nel formato 2
        /// </summary>
        public Constants_Excel.TipologiaFoglio_Format GetTipologiaFoglio { get; set; }
        

        /// <summary>
        /// Istanze di associazione alloys e proprieta di concentrazioni per il foglio corrente, questa definizione vale direttamente 
        /// per il secondo formato di foglio 
        /// </summary>
        public List<AlloyConcentrations_Association> AlloyInstances { get; set; }


        /// <summary>
        /// Istanze per i valori di concentrazione che vengono recuperati per il foglio corrente 
        /// </summary>
        public List<Excel_PropertiesContainer> ConcentrationsPropertiesInstances { get; set; }


        /// <summary>
        /// Istanze per i valori di lega che vengono recuperati per il foglio corrente 
        /// </summary>
        public List<Excel_PropertiesContainer> AlloyPropertiesInstances { get; set; }


        /// <summary>
        /// Validazione finale sul foglio correntemente in analisi
        /// </summary>
        public bool Validation_OK { get; set; }


        /// <summary>
        /// Messaggio di errore associato alla lettura del foglio corrente 
        /// </summary>
        public string ErrorMessageSheet { get; set; }


        /// <summary>
        /// Messaggio di warnings associato alla lettura del foglio corrente 
        /// </summary>
        public string WarningMessageSheet { get; set; }

        #endregion
    }


    /// <summary>
    /// Classe che indica l'associazione diretta, ove possibile, tra l'elemento con tutte le proprieta di lega che è stato letto 
    /// e quelli relativi alle concentrazioni per questo elemento di lega.
    /// FORMATO 1: questa associazione sarà fatta durante lo step 3 di controllo correttezza per le informazioni dell'una e dell'altra istanza 
    /// FORMATO 2: l'associazione è già possibile in quanto per l'elemento di lega esistono già tutti gli elementi di concentrazioni su stessa riga 
    /// (NB: stare attenti pero alle 2 validazioni precedenti)
    /// </summary>
    public class AlloyConcentrations_Association
    {
        /// <summary>
        /// Riferimento alla proprieta di lega corrente, questo è l'unico elemento per questo tipo di contenitore
        /// </summary>
        public Excel_PropertiesContainer RifProprietaLega { get; set; }


        /// <summary>
        /// Diversi elementi di concentrazione che sono stati letti per la lega inserita nell'oggetto sopra
        /// seguendo le diverse definizioni per le tipologie di formato
        /// </summary>
        public List<Excel_PropertiesContainer> RifConcentrationsCurrentLega { get; set; }


        /// <summary>
        /// Mi dice se rispetto alla validazione 4 di informazioni già presenti per il set di destinazione 
        /// passo la validazione 4 con la quale posso andare poi effettivamente a inserire la proprieta all'interno del set
        /// </summary>
        public bool Validation4_OK { get; set; }
    }
}
