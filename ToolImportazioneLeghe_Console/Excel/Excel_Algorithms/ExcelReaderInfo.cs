using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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

        #endregion
        

        #region RECUPERO INFORMAZIONI PER IL FORMATO 1 EXCEL

        /// <summary>
        /// Permette di inserire tutte le informazioni di lega all'interno dell'oggetto predisposto per contenere tutti i valori per le informazioni di lega 
        /// in lettura corrente e per il foglio in analisi
        /// </summary>
        /// <param name="currentFoglioExcel"></param>
        /// <param name="emptyLegheInfo"></param>
        /// <param name="filledLegheInfo"></param>
        /// <returns></returns>
        public static EsitoRecuperoInformazioniFoglio ReadLegheInfo(ExcelWorksheet currentFoglioExcel, Excel_Format1_Sheet emptyLegheInfo, out Excel_Format1_Sheet filledLegheInfo)
        {
            bool hoLettoAlmenoUnaLega = false;

            // validazione e inserimento del foglio in lettura corrente 
            if (currentFoglioExcel == null)
                throw new Exception(ExceptionMessages.EXCEL_READERINFO_FOGLIONULLPERLETTURA);

            _currentFoglioExcel = currentFoglioExcel;


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

            // TODO : recupero di tutte le informazioni


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
