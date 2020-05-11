using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ToolImportazioneLeghe_Console.Excel.Messages;
using ToolImportazioneLeghe_Console.Excel.Model_Excel;
using ToolImportazioneLeghe_Console.Utils;
using static ToolImportazioneLeghe_Console.Excel.Constants_Excel;

namespace ToolImportazioneLeghe_Console.Excel.Excel_Algorithms
{
    /// <summary>
    /// Inserimento di tutti gli algoritmi per la lettura delle informazioni dai file excel per i formati 1-2
    /// tutti gli oggetti che sono stati predisposti attraverso il primo metodo di validazione vengono quindi riempiti giustamente con le informazioni
    /// contenute nel foglio in base alla casistica
    /// </summary>
    public static class ExcelReaderInfo
    {
        #region ATTRIBUTI PRIVATI PER IL RECUPERO DELLE INFORMAZIONI SU FOGLIO CORRENTE

        /// <summary>
        /// Foglio excel correntemente in analisi e ricevuto in input
        /// per il caso di richiamo corrente 
        /// </summary>
        private static ExcelWorksheet _currentFoglioExcel;


        /// <summary>
        /// Indice di riga corrente 
        /// </summary>
        private static int _currentRowIndex = 0;


        /// <summary>
        /// Indice di colonna corrente 
        /// </summary>
        private static int _currentColIndex = 0;


        /// <summary>
        /// Mappatura degli indici di colonna iniziali sui quali vado a leggere le proprieta per le righe sottostanti
        /// </summary>
        private static Dictionary<int, string> PropertiesColMapper;


        /// <summary>
        /// Lista dei warnings eventualmente restituiti nel caso in cui durante il recupero dei valori qualche validazione non passa
        /// </summary>
        private static string _listaWarnings_LetturaFoglio = String.Empty;


        /// <summary>
        /// Lista degli errori eventualmente restituiti nel caso in cui durante il recupero dei valori qualche validazione non passa provocando errori
        /// per i quali non è possibile continuare con l'analisi 
        /// </summary>
        private static string _listaErrori_LetturaFoglio = String.Empty;

        #endregion


        #region RECUPERO INFORMAZIONI PER IL FORMATO 1 EXCEL

        /// <summary>
        /// Permette di inserire tutte le informazioni di lega all'interno dell'oggetto predisposto per contenere tutti i valori per le informazioni di lega 
        /// in lettura corrente e per il foglio in analisi
        /// Si tratta della prima tipologia di foglio per il primo formato excel disponibile, ovvero quello di lettura delle leghe
        /// In particolare vengono anche restituiti in output la lista degli errori e warnings per l'iterazione corrente, sui quali devono essere configurati eventualmente dei logs appositi 
        /// </summary>
        /// <param name="currentFoglioExcel"></param>
        /// <param name="emptyLegheInfo"></param>
        /// <param name="filledLegheInfo"></param>
        /// <param name="listaWarnings_LetturaFoglio"></param>
        /// <param name="listaErrori_LetturaFoglio"></param>
        /// <returns></returns>
        public static EsitoRecuperoInformazioniFoglio ReadLegheInfo(ExcelWorksheet currentFoglioExcel, Excel_Format1_Sheet emptyLegheInfo, out Excel_Format1_Sheet filledLegheInfo, out string listaWarnings_LetturaFoglio, out string listaErrori_LetturaFoglio)
        {
            // inizializzazione delle 2 liste di errori warnings per l'iterazione sul foglio corrente 
            _listaWarnings_LetturaFoglio = String.Empty;
            _listaErrori_LetturaFoglio = String.Empty;

            // validazione e inserimento del foglio in lettura corrente 
            if (currentFoglioExcel == null)
                throw new Exception(ExceptionMessages.EXCEL_READERINFO_FOGLIONULLPERLETTURA);

            _currentFoglioExcel = currentFoglioExcel;

            // inizializzazione istanza di primo foglio
            Excel_Format1_Sheet1_Rows currentRowPropertiesFoglio = new Excel_Format1_Sheet1_Rows();


            // validazioni su indici di lettura iniziali e finali per l'header di colonna 
            if (emptyLegheInfo.StartingCol_letturaLeghe == 0 || emptyLegheInfo.EndingCol_letturaLeghe == 0)
                throw new Exception(ExceptionMessages.EXCEL_READERINFO_INDICIDILETTURAZERO);
            if (emptyLegheInfo.StartingCol_letturaLeghe >= emptyLegheInfo.EndingCol_letturaLeghe)
                throw new Exception(ExceptionMessages.EXCEL_READERINFO_INDICINONVALIDI);

            // validazioni su indici di lettura iniziali e finali per l'header di riga 
            if (emptyLegheInfo.StartingRow_letturaLeghe == 0)
                throw new Exception(ExceptionMessages.EXCEL_READERINFO_INDICIDILETTURAZERO);

            // fisso indice di riga alla lettura per le proprieta correnti
            _currentRowIndex = emptyLegheInfo.StartingRow_letturaLeghe;

            // inserisco tutte le proprieta necessarie per la futura lettura corrente 
            FillPropertyMapper(emptyLegheInfo.StartingCol_letturaLeghe, emptyLegheInfo.EndingCol_letturaLeghe);

            // genero errore se non ho letto neanche un header per le proprieta per il foglio corrente 
            // (se il dizionario riempito allo step precedente ha una dimensione minore di quella della lista delle proprieta obbligatorie per il caso corrente)
            if (PropertiesColMapper == null)
                throw new Exception(ExceptionMessages.EXCEL_READERINFO_PROPRIETAINTERNANULLA);
            if (PropertiesColMapper.Count() < Constants_Excel.PROPRIETAOBBLIGATORIE_FORMAT1_SHEET1.ToList().Count())
                throw new Exception(ExceptionMessages.EXCEL_READERINFO_VINCOLILETTURAPROPRIETANONRISPETTATI);


            // inizio iterazione per il recupero dei valori
            while (_currentRowIndex <= _currentFoglioExcel.Dimension.End.Row)
            {
                Excel_PropertyWrapper currentSetProperties;

                // inserisco la riga solamente se sono rispettati i vincoli sulle proprieta obbligatorie
                if (ProprietaLega(out currentSetProperties))
                    currentRowPropertiesFoglio.ExcelSheet1_LegheProperties.Add(currentSetProperties);

                _currentRowIndex++;
            }

            // attribuzione dei valori finali
            emptyLegheInfo.GetListInfoLega_Type1 = currentRowPropertiesFoglio;
            filledLegheInfo = emptyLegheInfo;

            // se non è stata letta alcuna proprieta allora ritorno il recupero scorretto
            if (filledLegheInfo.GetListInfoLega_Type1.ExcelSheet1_LegheProperties.Count() == 0)
            {

                _listaErrori_LetturaFoglio += String.Format(Excel_ErrorMessages.Formato1_Foglio1_Leghe.ERRORE_NESSUNA_INFORMAZIONE_LETTA, _currentFoglioExcel.Name);

                listaErrori_LetturaFoglio = _listaErrori_LetturaFoglio;
                listaWarnings_LetturaFoglio = _listaWarnings_LetturaFoglio;

                return EsitoRecuperoInformazioniFoglio.RecuperoConErrori;
            }




            listaErrori_LetturaFoglio = _listaErrori_LetturaFoglio;
            listaWarnings_LetturaFoglio = _listaWarnings_LetturaFoglio;

            filledLegheInfo = emptyLegheInfo;
            return EsitoRecuperoInformazioniFoglio.RecuperoCorretto;
        }


        /// <summary>
        /// Mi permette di mappare una sola e unica volta le proprieta lette per il foglio excel corrente 
        /// questo mapper mi servira per andare a dare un effettivo valore ai diversi elementi di riga letti per 
        /// i casi seguenti
        /// </summary>
        /// <param name="startingColIndex"></param>
        /// <param name="endingColIndex"></param>
        private static void FillPropertyMapper(int startingColIndex, int endingColIndex)
        {
            int currentIndex = startingColIndex;

            // inizializzazione della lista di mapper
            PropertiesColMapper = new Dictionary<int, string>();

            while(currentIndex <= endingColIndex)
            {
                if (_currentFoglioExcel.Cells[_currentRowIndex, currentIndex].Value != null)
                    PropertiesColMapper.Add(currentIndex, _currentFoglioExcel.Cells[_currentRowIndex, currentIndex].Value.ToString().ToUpper());
            }
        }


        /// <summary>
        /// Permette il recupero di tutte le proprieta per la riga corrente e il loro eventuale inserimento in una lista di proprieta 
        /// in lettura corrente per la riga per la prima tipologia di foglio per il PRIMO FORMATO
        /// </summary>
        /// <param name="readProperties"></param>
        /// <returns></returns>
        private static bool ProprietaLega(out Excel_PropertyWrapper readProperties)
        {
            
            // istanza per le proprieta correntemente lette per la lega 
            readProperties = new Excel_PropertyWrapper(Constants_Excel.PROPRIETAOBBLIGATORIE_FORMAT1_SHEET1, Constants_Excel.PROPRIETAOPZIONALI_FORMAT1_SHEET1, TipologiaPropertiesFoglio.Format1_Foglio1_Leghe);


            // lettura proprieta di riga 
            foreach (KeyValuePair<int, string> currentProperty in PropertiesColMapper)
            {
                // lettura proprieta corrente 
                if(_currentFoglioExcel.Cells[_currentRowIndex, currentProperty.Key].Value != null)
                {
                    // inserimento di una proprieta obbligatoria
                    if (Constants_Excel.PROPRIETAOBBLIGATORIE_FORMAT1_SHEET1.Contains(currentProperty.Value.ToUpper()))
                        readProperties.InsertMandatoryValue(currentProperty.Value.ToUpper(), _currentFoglioExcel.Cells[_currentRowIndex, currentProperty.Key].Value.ToString());


                    // inserimento di una proprieta opzionale
                    if (Constants_Excel.PROPRIETAOPZIONALI_FORMAT1_SHEET1.Contains(currentProperty.Value.ToUpper()))
                        readProperties.InsertOptionalValue(currentProperty.Value.ToUpper(), _currentFoglioExcel.Cells[_currentRowIndex, currentProperty.Key].Value.ToString());
                }
            }

            // le proprieta obbligatorie non sono state lette correttamente per il foglio corrente 
            if (readProperties.CounterMandatoryProperties < Constants_Excel.PROPRIETAOBBLIGATORIE_FORMAT1_SHEET1.Count())
                return false;

            return true;
        }



        /// <summary>
        /// Permette di recuperare tutte le informazioni di concentrazione all'interno dell'oggetto predisposto per contenere tutti i valori per le informazioni di concentrazioni
        /// in lettura corrente e per il foglio in analisi
        /// </summary>
        /// <param name="currentFoglioExcel"></param>
        /// <param name="emptyConcentrationsInfo"></param>
        /// <param name="filledConcentrationsInfo"></param>
        /// <returns></returns>
        public static EsitoRecuperoInformazioniFoglio ReadConcentrationsInfo(ExcelWorksheet currentFoglioExcel, Excel_Format1_Sheet emptyConcentrationsInfo, out Excel_Format1_Sheet filledConcentrationsInfo)
        {
            filledConcentrationsInfo = emptyConcentrationsInfo;
            return EsitoRecuperoInformazioniFoglio.RecuperoCorretto;
        }

        #endregion


        #region RECUPERO INFORMAZIONI PER IL FORMATO 2 EXCEL 

        /// <summary>
        /// Permette il recupero di tutte le informazioni per leghe e concentrazioni all'interno dell'oggetto predisposto per contenenere tutti i vlaori per le informazioni di
        /// leghe e concentrazioni in lettura corrente e per il foglio in analisi
        /// </summary>
        /// <param name="currentFoglioExcel"></param>
        /// <param name="emptyInfo"></param>
        /// <param name="filledInfo"></param>
        /// <returns></returns>
        public static EsitoRecuperoInformazioniFoglio ReadInfoFormat2(ExcelWorksheet currentFoglioExcel, Excel_Format2_Sheet emptyInfo, Excel_Format2_Sheet filledInfo)
        {
            filledInfo = emptyInfo;
            return EsitoRecuperoInformazioniFoglio.RecuperoCorretto;
        }

        #endregion

    }
}
