using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ToolImportazioneLeghe_Console.Excel.Model_Excel;

namespace ToolImportazioneLeghe_Console.Excel.Excel_Algorithms
{
    /// <summary>
    /// Inserimento di tutti gli algoritmi per la lettura delle informazioni dai file excel per i formati 1-2
    /// tutti gli oggetti che sono stati predisposti attraverso il primo metodo di validazione vengono quindi riempiti giustamente con le informazioni
    /// contenute nel foglio in base alla casistica
    /// </summary>
    public static class ExcelReaderInfo
    {
        #region RECUPERO INFORMAZIONI PER IL FORMATO 1 EXCEL

        /// <summary>
        /// Permette di inserire tutte le informazioni di lega all'interno dell'oggetto predisposto per contenere tutti i valori per le informazioni di lega 
        /// in lettura corrente e per il foglio in analisi
        /// </summary>
        /// <param name="currentFoglioExcel"></param>
        /// <param name="emptyLegheInfo"></param>
        /// <param name="filledLegheInfo"></param>
        /// <returns></returns>
        public static bool ReadLegheInfo(ExcelWorksheet currentFoglioExcel, Excel_Format1_Sheet emptyLegheInfo, out Excel_Format1_Sheet filledLegheInfo)
        {
            filledLegheInfo = emptyLegheInfo;
            return false;
        }


        /// <summary>
        /// Permette di recuperare tutte le informazioni di concentrazione all'interno dell'oggetto predisposto per contenere tutti i valori per le informazioni di concentrazioni
        /// in lettura corrente e per il foglio in analisi
        /// </summary>
        /// <param name="currentFoglioExcel"></param>
        /// <param name="emptyConcentrationsInfo"></param>
        /// <param name="filledConcentrationsInfo"></param>
        /// <returns></returns>
        public static bool ReadConcentrationsInfo(ExcelWorksheet currentFoglioExcel, Excel_Format1_Sheet emptyConcentrationsInfo, out Excel_Format1_Sheet filledConcentrationsInfo)
        {
            filledConcentrationsInfo = emptyConcentrationsInfo;
            return false;
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
        public static bool ReadInfoFormat2(ExcelWorksheet currentFoglioExcel, Excel_Format2_Sheet emptyInfo, Excel_Format2_Sheet filledInfo)
        {
            filledInfo = emptyInfo;
            return false;
        }

        #endregion

    }
}
