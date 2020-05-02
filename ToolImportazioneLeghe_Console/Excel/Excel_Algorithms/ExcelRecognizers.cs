using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ToolImportazioneLeghe_Console.Excel.Model_Excel;
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


        /// <summary>
        /// Traccia di riga correntemenete in analisi
        /// </summary>
        private static int _currentRowIndex = 0;


        /// <summary>
        /// Traccia di colonna correntemente in analisi
        /// </summary>
        private static int _currentColIndex = 0;


        /// <summary>
        /// Lista per l'eventuale riconoscimento di un quadrante delle concentrazioni per la seconda tipologia di foglio
        /// per il primo formato
        /// </summary>
        private static List<Excel_Format1_Sheet2_ConcQuadrant> _listaQuadrantiConcentrazioni;


        /// <summary>
        /// Proprieta obbligatorie per il riconoscimento header per il quadrante delle concentrazioni corrente 
        /// </summary>
        private static List<string> _mandatoryInfo_format1_sheet2 = Constants_Excel.PROPRIETAOBBLIGATORIE_FORMAT1_SHEET2.ToList();


        /// <summary>
        /// Proprieta opzionali per il riconoscimento header per il quadrante delle concentrazioni corrente
        /// </summary>
        private static List<string> _optionalInfo_format1_sheet2 = Constants_Excel.PROPRIETAOPZIONALI_FORMAT1_SHEET2.ToList();


        /// <summary>
        /// Indica il numero di righe vuote massimo che posso leggere prima di incontrare l'header per il quadrante di concentrazioni 
        /// a partire dal primo riconoscimento fatto per il title
        /// </summary>
        private static int LIMIT_ROW_HEADERCONCENTRATION_RECOGNITION = 2;


        /// <summary>
        /// Indicazione della colonna dei CRITERI per la quale devo andare a riconoscere sulle righe successive la presenza di un certo elemento
        /// definito
        /// </summary>
        private static int _colCriteriIndex = 0;


        /// <summary>
        /// Traccia del nome di colonna criteri per l'eventuale lettura delle definizioni degli elementi sottostanti
        /// </summary>
        private const string COLCRITERI_HEADER = "CRITERI";

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

            _currentRowIndex = 0;
            _currentColIndex = 0;
            

            // inserimento dei valori per il limite massimo di riga / colonna entro il quale devo riconoscere l'informazione 
            indexRow_Max = (currentWorksheet.Dimension.End.Row <= LIMIT_ROW) ? currentWorksheet.Dimension.End.Row : LIMIT_ROW;
            intexCol_Max = (currentWorksheet.Dimension.End.Column <= LIMIT_COL) ? currentWorksheet.Dimension.End.Column : LIMIT_COL;

            do
            {
                _currentColIndex++;

                do
                {
                    _currentRowIndex++;

                    if(HoRiconosciutoHeader_Format1_Leghe())
                    {
                        startingRow = _currentRowIndex;
                        startingCol = _currentColIndex;

                        return true;
                    }

                }
                while (_currentRowIndex <= indexRow_Max);

            }
            while (_currentColIndex <= intexCol_Max);


            return false;
        }


        /// <summary>
        /// Mi dice se ho riconosciuto l'header relativo alle informazioni per le leghe sul primo foglio per il primo 
        /// formato excel
        /// </summary>
        /// <param name="startingRow"></param>
        /// <param name="startingCol"></param>
        /// <returns></returns>
        private static bool HoRiconosciutoHeader_Format1_Leghe()
        {
           
            List<string> recognizedMandatoryProperties = new List<string>(); ;
           

            while (!(_foglioExcelCorrente.Cells[_currentRowIndex, _currentColIndex].Value == null))
            {
                if (_mandatoryInfo_format1_sheet1.Contains(_foglioExcelCorrente.Cells[_currentRowIndex, _currentColIndex].Value) && !(recognizedMandatoryProperties.Contains(_foglioExcelCorrente.Cells[_currentRowIndex, _currentColIndex].Value)))
                    recognizedMandatoryProperties.Add(_foglioExcelCorrente.Cells[_currentRowIndex, _currentColIndex].Value.ToString());

                _currentColIndex++;
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
        /// Viene restituito in output la lista dei quadranti excel letti nel caso in cui sui sia effettivamente riconosciuto il foglio come 
        /// foglio per le concentrazioni
        /// </summary>
        /// <param name="currentWorksheet"></param>
        /// <param name="listaQuadrantiConcentrazioni"></param>
        /// <returns></returns>
        public static bool Recognize_Format1_InfoConcentrations(ref ExcelWorksheet currentWorksheet, out List<Excel_Format1_Sheet2_ConcQuadrant> listaQuadrantiConcentrazioni)
        {
            // validazioni di partenza 
            if (currentWorksheet == null)
                throw new Exception(ExceptionMessages.EXCEL_FILENOTINMEMORY);

            // ritorno eccezione anche se incontro una colonna definita per il ricoscimento degli elementi ma che non appartiene 
            // alle definizioni per le proprieta obbligatorie di riconoscimento delle concentrazioni
            if (!_mandatoryInfo_format1_sheet2.Contains(COLCRITERI_HEADER))
                throw new Exception(ExceptionMessages.EXCEL_COLCRITERINONPRESENTE);


            _foglioExcelCorrente = currentWorksheet;

            _listaQuadrantiConcentrazioni = new List<Excel_Format1_Sheet2_ConcQuadrant>();
            
            _currentRowIndex= 0;
            _currentColIndex = 0;

            int indexRow_Max = (currentWorksheet.Dimension.End.Row <= LIMIT_ROW) ? currentWorksheet.Dimension.End.Row : LIMIT_ROW;
            int indexCol_Max = (currentWorksheet.Dimension.End.Column <= LIMIT_COL) ? currentWorksheet.Dimension.End.Column : LIMIT_COL;

            do
            {
                _currentColIndex++;

                do
                {
                    _currentRowIndex++;

                    // riposizionamento indice di riga 
                    if(HoRiconosciutoFormat1_Concentrazioni())
                        indexRow_Max = (currentWorksheet.Dimension.End.Row <= _currentRowIndex + LIMIT_ROW) ? currentWorksheet.Dimension.End.Row : _currentRowIndex + LIMIT_ROW;

                }
                while (_currentRowIndex <= indexRow_Max);

                // ricalcolo eventuale indice solonna 
                _currentColIndex = RicalcolaIndiceColonna();

                // ricalcolo il limite per la lettura su colonna 
                indexCol_Max = (currentWorksheet.Dimension.End.Column <= _currentColIndex + LIMIT_COL) ? currentWorksheet.Dimension.End.Column : _currentColIndex + LIMIT_COL;
            }
            while (_currentColIndex <= indexCol_Max);

            // attribuzione con gli eventuali quadranti di concentrazione letti
            listaQuadrantiConcentrazioni = _listaQuadrantiConcentrazioni;


            // ritorno true solo se ho riconosciuto almeno un quadrante di concentrazioni per il foglio excel corrente 
            if (listaQuadrantiConcentrazioni.Count() > 0)
                return true;

            return false;

        }


        /// <summary>
        /// Riconoscimento vero e proprio per l'eventuale quadrante delle concentrazioni per il foglio corrente 
        /// vengono anche ricalcolati gli indici di spostamento per riga e colonna correnti
        /// </summary>
        /// <returns></returns>
        private static bool HoRiconosciutoFormat1_Concentrazioni()
        {
            Excel_Format1_Sheet2_ConcQuadrant riconoscimentoQuadranteCorrente = new Excel_Format1_Sheet2_ConcQuadrant();

            #region VERIFICA ESISTENZA TITOLO DI LEGA

            // verifico esistenza del titolo
            if (_foglioExcelCorrente.Cells[_currentRowIndex, _currentColIndex].Value == null)
                return false;

            riconoscimentoQuadranteCorrente.NomeMateriale = _foglioExcelCorrente.Cells[_currentRowIndex, _currentColIndex].Value.ToString();
            riconoscimentoQuadranteCorrente.StartingRow_Title = _currentRowIndex;
            riconoscimentoQuadranteCorrente.StartigCol = _currentColIndex;

            #endregion

            // riconoscimento header dopo iterazione corrente 
            bool riconoscimentoHeader = false;

            // incremento posizione riga 
            _currentRowIndex++;

            // attribuzione riga massima per il riconoscimento dell'header delle concentrazioni
            int maxHeader_rowIndex = _currentRowIndex + LIMIT_ROW_HEADERCONCENTRATION_RECOGNITION;

            // indice di colonna massimo per il quadrante di concentrazioni corrente (corrispondente a ultima lettura header)
            int maxColIndex = 0;

            while((!riconoscimentoHeader) || _currentRowIndex <= maxHeader_rowIndex)
            {
                riconoscimentoHeader = RecognizeHeaderConcentrations(out maxColIndex);
                if (riconoscimentoHeader)
                {
                    riconoscimentoQuadranteCorrente.StartingRow_Headers = _currentRowIndex;
                    riconoscimentoQuadranteCorrente.EndingCol = maxColIndex;
                    break;
                }
                    

                // incremento per questa iterazione solamente nel caso in cui non abbia ancora riconosciuto l'header corrente di concentrazioni
                _currentRowIndex++;
            }

            // se non ho riconosciuto l'header allora esco senza aver riconosciuto il quadrante 
            if (!riconoscimentoHeader)
                return false;

            #region RICONOSCIMENTO HEADERS CONCENTRAZIONI

            // riconoscimento del set di concentrazioni per il quadrante corrente 
            bool riconoscimentoConcentrationi = false;

            // incremento posizione riga 
            _currentRowIndex++;
            // inserimento della eventuale posizione di partenza per la lettura delle concentrazioni
            int startingPosConc = _currentRowIndex;

            int maxConc_RowIndex = _currentRowIndex + LIMIT_ROW_HEADERCONCENTRATION_RECOGNITION;

            while((!riconoscimentoConcentrationi) || _currentRowIndex <= maxConc_RowIndex)
            {
                riconoscimentoConcentrationi = RecognizeContentConcentrations();
                if (riconoscimentoConcentrationi)
                {
                    riconoscimentoQuadranteCorrente.StartingRow_Concentrations = startingPosConc;
                    riconoscimentoQuadranteCorrente.EndingRow_Concentrations = _currentRowIndex;
                    break;
                }
                    

                // incremento perché non sono ancora riuscito a trovare le concentrazioni per questa iterazione
                _currentRowIndex++;
                startingPosConc = _currentRowIndex;
            }

            #endregion


            #region AGGIUNTA NEL NUOVO QUADRANTE NELLE DEFINIZIONI E RIITORNO VERO

            if(riconoscimentoHeader && riconoscimentoConcentrationi)
            {
                _listaQuadrantiConcentrazioni.Add(riconoscimentoQuadranteCorrente);
                return true;
            }

            #endregion


            return false;
        }


        /// <summary>
        /// Permette il riconoscimento per l'header delle proprieta di concentrazioni corrente 
        /// viene restituito il set di tutte le proprieta riconosciute
        /// In questa fase viene anche calcolato il massimo indice di colonna per il quadrante corrente 
        /// (corrispondente all'ultima colonna per la lettura dell'header)
        /// </summary>
        /// <param name="maxColIndex"></param>
        /// <returns></returns>
        private static bool RecognizeHeaderConcentrations(out int maxColIndex)
        {
            // lista di tutte le proprieta riconosciute
            List<string> recognizedMandatoryProperties = new List<string>();
            
            

            int currentRowIndexCopy = _currentRowIndex;
            int currentColIndexCopy = _currentColIndex;

            maxColIndex = _currentColIndex;

            if (_foglioExcelCorrente.Cells[currentRowIndexCopy, currentColIndexCopy].Value == null)
                return false;

            while(_foglioExcelCorrente.Cells[currentRowIndexCopy, currentColIndexCopy].Value != null)
            {
                if(_mandatoryInfo_format1_sheet2.Contains(_foglioExcelCorrente.Cells[currentRowIndexCopy, currentColIndexCopy].Value.ToString().ToUpper()))
                    recognizedMandatoryProperties.Add(_foglioExcelCorrente.Cells[currentRowIndexCopy, currentColIndexCopy].Value.ToString().ToUpper());

                // tengo traccia dell'indice di colonna dei CRITERI per la successiva eventuale lettura degli elementi sottostanti
                if (_foglioExcelCorrente.Cells[currentRowIndexCopy, currentColIndexCopy].Value.ToString().ToUpper() == COLCRITERI_HEADER)
                    _colCriteriIndex = currentColIndexCopy;

                // incremento indice di colonna relativo agli headers
                maxColIndex++;
            }

            if (recognizedMandatoryProperties.Count() == _mandatoryInfo_format1_sheet2.Count())
            {
                return true;
            }
                

            
            return false;
        }


        /// <summary>
        /// Riconoscimento posizione per gli elementi correnti all'interno del foglio 
        /// mi fermo solamente quando non riconosco piu un elemento 
        /// </summary>
        /// <returns></returns>
        private static bool RecognizeContentConcentrations()
        {
            bool hoLettoAlmenoUnPossibileValoreElemento = false;

            while(_foglioExcelCorrente.Cells[_currentRowIndex, _colCriteriIndex].Value != null)
            {
                if (!_foglioExcelCorrente.Cells[_currentRowIndex, _colCriteriIndex + 1].Merge == true)
                {
                    hoLettoAlmenoUnPossibileValoreElemento = true;
                    _currentRowIndex++;
                }
                   
            }

            if (hoLettoAlmenoUnPossibileValoreElemento)
                return true;
            
            return false;
        }


        /// <summary>
        /// Permette di calcolare l'indice per il riposizionamento eventuale della colonna per il riconoscimento
        /// di altri quadranti all'interno del foglio excel delle concentrazioni
        /// </summary>
        /// <returns></returns>
        private static int RicalcolaIndiceColonna()
        {
            int newColIndex = _currentColIndex++;

            if (_listaQuadrantiConcentrazioni != null)
                if (_listaQuadrantiConcentrazioni.Count() > 0)
                    newColIndex = _listaQuadrantiConcentrazioni.Select(x => x.EndingCol).Max();

            return newColIndex;
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
