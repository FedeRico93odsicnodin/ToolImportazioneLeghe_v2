using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ToolImportazioneLeghe_Console.Utils;

namespace ToolImportazioneLeghe_Console.Excel.Excel_Algorithms
{
    /// <summary>
    /// Classe contenente tutti gli algoritmi per il riconoscimento corretto delle 3 tipologie di foglio excel 
    /// le prime 2 riguardano il primo formato per il quale si puo individuare rispettivamente un foglio relativo alle concentrazioni o uno relativo alle informazioni di lega 
    /// la terza riguarda invece la tipologia relativa al secondo formato per il quale si potranno leggere sia delle informazioni di lega che delle concentrazioni 
    /// </summary>
    public static class ExcelRecognizers
    {
        #region ATTRIBUTI PRIVATI

        /// <summary>
        /// Attribuzione al momenti di richiamo di uno dei diversi metodi in analisi, mappatura di tutte le informazioni per il foglio 
        /// excel correntemente in analisi
        /// </summary>
        private static ExcelWorksheet _foglioExcelCorrente;


        /// <summary>
        /// Limite nella lettura delle righe prima che non sia stata ancora trovata nessuna informazione utile al 
        /// fine del riconoscimento
        /// </summary>
        private static int LIMIT_ROW = 20;


        /// <summary>
        /// Limite nella lettura delle colonne prima che non sia stata ancora trovata nessuna informazione utile al 
        /// fine del riconoscimento
        /// </summary>
        private static int LIMIT_COL = 20;


        /// <summary>
        /// Lista degli headers obbligatori per il foglio di leghe per il formato 1
        /// </summary>
        private static List<string> _mandatoryInfo_format1_sheet1 = Constants_Excel.PROPRIETAOBBLIGATORIE_FORMAT1_SHEET1.ToList();


        /// <summary>
        /// Lista degli headers opzionali per il foglio di leghe per il formato 1
        /// </summary>
        private static List<string> _optionalInfo_format1_sheet1 = Constants_Excel.PROPRIETAOPZIONALI_FORMAT1_SHEET1.ToList();

        #endregion


        /// <summary>
        /// Mi permette di riconoscere se il foglio corrente appartiene alla categoria relativa alle informazioni di lega 
        /// per il primo formato di foglio excel disponibile
        /// </summary>
        /// <param name="currentWorksheet"></param>
        /// <param name="startingRow"></param>
        /// <param name="startingCol"></param>
        /// <returns></returns>
        public static bool Recognize_Format1_InfoLeghe(ref ExcelWorksheet currentWorksheet, out int startingRow, out int startingCol)
        {
            // validazioni di partenza 
            if (currentWorksheet == null)
                throw new Exception(ExceptionMessages.EXCEL_FILENOTINMEMORY);

            _foglioExcelCorrente = currentWorksheet;

            startingRow = 0;
            startingCol = 0;

            int indexRow_Max = 1;
            int intexCol_Max = 1;

            int currentRow = 0;
            int currentCol = 0;
            

            // inserimento dei valori per il limite massimo di riga / colonna entro il quale devo riconoscere l'informazione 
            indexRow_Max = (currentWorksheet.Dimension.End.Row <= LIMIT_ROW) ? currentWorksheet.Dimension.End.Row : LIMIT_ROW;
            intexCol_Max = (currentWorksheet.Dimension.End.Column <= LIMIT_COL) ? currentWorksheet.Dimension.End.Column : LIMIT_COL;

            do
            {
                currentCol++;

                do
                {
                    currentRow++;

                    if(HoRiconosciutoHeader_Format1_Leghe(currentRow, currentCol))
                    {
                        startingRow = currentRow;
                        startingCol = currentCol;

                        return true;
                    }

                }
                while (currentRow <= indexRow_Max);

            }
            while (currentCol <= intexCol_Max);


            return false;
        }


        /// <summary>
        /// Mi dice se ho riconosciuto l'header relativo alle informazioni per le leghe sul primo foglio per il primo 
        /// formato excel
        /// </summary>
        /// <param name="startingRow"></param>
        /// <param name="startingCol"></param>
        /// <returns></returns>
        private static bool HoRiconosciutoHeader_Format1_Leghe(int startingRow, int startingCol)
        {

            List<string> recognizedMandatoryProperties = new List<string>(); ;
            int currentRow = startingRow;
            int currentCol = startingCol;


            while (!(_foglioExcelCorrente.Cells[currentRow, currentCol].Value == null))
            {
                if (_mandatoryInfo_format1_sheet1.Contains(_foglioExcelCorrente.Cells[currentRow, currentCol].Value) && !(recognizedMandatoryProperties.Contains(_foglioExcelCorrente.Cells[currentRow, currentCol].Value)))
                    recognizedMandatoryProperties.Add(_foglioExcelCorrente.Cells[currentRow, currentCol].Value.ToString());

                currentCol++;
            }

            if (recognizedMandatoryProperties.Count() == _mandatoryInfo_format1_sheet1.Count())
            {
                return true;
            }
                
            return false;
        }


        /// <summary>
        /// Mi permette di riconoscere se il foglio corrente appartiene alla categoria relativa alle informazioni per le concentrazioni
        /// per il primo formato di foglio excel disponibile
        /// </summary>
        /// <param name="currentWorksheet"></param>
        /// <returns></returns>
        public static bool Recognize_Format1_InfoConcentrations(ExcelWorksheet currentWorksheet)
        {
            return false;
        }


        /// <summary>
        /// Mi permette di riconoscere se il foglio corrente appartiene alla categoria relativa alle informazioni per concentrazioni leghe 
        /// in lettura dal secondo formato excel disponibile
        /// </summary>
        /// <param name="currentWorksheet"></param>
        /// <returns></returns>
        public static bool Recognize_Format2_InfoLegheConcentrazioni(ExcelWorksheet currentWorksheet)
        {
            return false;
        }
    }
}
