using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ToolImportazioneLeghe_Console.Excel.Messages;
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

        #region RICONOSCIMENTO DEI PRIMI 2 FOGLI PER LA PRIMA TIPOLOGIA DI FORMATO

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


        #region RICONOSCIMENTO DELLA TERZA TIPOLOGIA DI FOGLIO PER IL FORMATO 2 NEL QUALE SONO PRESENTI SIA LE INFORMAZIONI DI LEGHE CHE DI CONCENTRAZIONI

        /// <summary>
        /// Mi darà indicazione finale rispetto a dove iniziare a leggere le proprieta per la lega 
        /// </summary>
        private static int _startReadingProperties_row = 0;


        /// <summary>
        /// Indicazione della colonna di partenza dalla quale iniziare a leggere per le proprieta di lega
        /// </summary>
        private static int _startReadingLegheProperties_col = 0;


        /// <summary>
        /// Indicazione degli oggetti di concentrazione da riempire rispetto alle posizioni excel individuate per questi
        /// in questa fases
        /// </summary>
        private static List<Excel_Format2_ConcColumns> _currentConcentrations;


        /// <summary>
        /// Proprieta obbligatorie leghe per il secondo formato excel
        /// </summary>
        private static List<string> _mandatoryInfo_Leghe_Format2 = Constants_Excel.PROPRIETAOBBLIGATORIE_FORMAT2_LEGHE.ToList();


        /// <summary>
        /// Proprieta opzionali leghe per il secondo formato excel
        /// </summary>
        private static List<string> _optionalInfo_Leghe_Format2 = Constants_Excel.PROPRIETAOPZIONALI_FORMAT2_leghe.ToList();


        /// <summary>
        /// Proprieta obbligatorie di elemento per la seconda tipologia di formato excel 
        /// </summary>
        private static List<string> _mandatoryInfo_Concentrations_Format2 = Constants_Excel.PROPRIETAOBBLIGATORIE_ELEM_FORMAT2.ToList();


        /// <summary>
        /// Proprieta opzionali di elemento per la seconda tipologia di formato excel 
        /// </summary>
        private static List<string> _optionalInfo_Concentrations_Format2 = Constants_Excel.PROPRIETAOPZIONALI_ELEM_FORMAT2.ToList();


        /// <summary>
        /// Stringa dei possibili errori emersi durante la lettura per il file excel corrente 
        /// </summary>
        private static string _error_Messages_ReadExcel = String.Empty;


        /// <summary>
        /// Stringa dei possibili warnings durante la procedura di lettura per il file excel corrente 
        /// </summary>
        private static string _warning_Messages_ReadExcel = String.Empty;


        /// <summary>
        /// Stringa relativa al messaggio finale di errore in restituzione per la lettura di un quadrante di concentrazione 
        /// per il formato 2, prima tipologia di foglio 
        /// </summary>
        private static string _finalError_Message_ConcentrationQuadrant = String.Empty;


        /// <summary>
        /// Stringa relativa al messaggio finale di warning in restituzione per la lettura di un quadrante di concentrazione 
        /// per il formato 2, prima tipologia di foglio
        /// </summary>
        private static string _finalWarning_Message_ConcentrationQuadrant = String.Empty;

        #endregion

        #endregion


        /// <summary>
        /// Mi permette di riconoscere se il foglio corrente appartiene alla categoria relativa alle informazioni di lega 
        /// per il primo formato di foglio excel disponibile
        /// se viene riconosciuto correttamente per l'header delle leghe viene anche restituito come indice di colonna l'ultimo indice sul quale 
        /// è stata letta la proprieta da riconoscere nel foglio
        /// </summary>
        /// <param name="currentWorksheet"></param>
        /// <param name="startingRow"></param>
        /// <param name="startingCol"></param>
        /// <param name="endingColIndexHeaders"></param>
        /// <param name="possibleErrorMessages"></param>
        /// <param name="possibleWarningMessages"></param>
        /// <returns></returns>
        public static Constants_Excel.EsitoRecuperoInformazioniFoglio Recognize_Format1_InfoLeghe(ref ExcelWorksheet currentWorksheet, 
            out int startingRow, 
            out int startingCol, 
            out int endingColIndexHeaders,
            out string possibleErrorMessages,
            out string possibleWarningMessages)
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

            // indicazione di ultimo indice di lettura letto per gli header di proprieta che vengono riconosciuti per il foglio excel corrente 
            endingColIndexHeaders = 0;


            // inizializzazione delle 2 stringhe relative ai messaggi di segnalazione warnings / errori per il file excel correntemente in analisi
            _error_Messages_ReadExcel = String.Empty;
            _warning_Messages_ReadExcel = String.Empty;


            // inserimento dei valori per il limite massimo di riga / colonna entro il quale devo riconoscere l'informazione 
            indexRow_Max = (currentWorksheet.Dimension.End.Row <= LIMIT_ROW) ? currentWorksheet.Dimension.End.Row : LIMIT_ROW;
            intexCol_Max = (currentWorksheet.Dimension.End.Column <= LIMIT_COL) ? currentWorksheet.Dimension.End.Column : LIMIT_COL;

            do
            {
                _currentColIndex++;

                // azzero nuovamente indice di riga per la nuova iterazione su colonna
                _currentRowIndex = 0;

                do
                {
                    _currentRowIndex++;

                    if(HoRiconosciutoHeader_Format1_Leghe(out endingColIndexHeaders))
                    {
                        startingRow = _currentRowIndex;
                        startingCol = _currentColIndex;

                        // attribuzione delle 2 stinghe per i messaggi di errore e warnings finali
                        possibleErrorMessages = _error_Messages_ReadExcel;
                        possibleWarningMessages = _warning_Messages_ReadExcel;

                        if (possibleWarningMessages != String.Empty)
                            return Constants_Excel.EsitoRecuperoInformazioniFoglio.RecuperoConWarnings;


                        return Constants_Excel.EsitoRecuperoInformazioniFoglio.RecuperoCorretto;
                    }

                }
                while (_currentRowIndex <= indexRow_Max);

            }
            while (_currentColIndex <= intexCol_Max);


            // attribuzione delle 2 stinghe per i messaggi di errore e warnings finali
            possibleErrorMessages = _error_Messages_ReadExcel;
            possibleWarningMessages = _warning_Messages_ReadExcel;

            return Constants_Excel.EsitoRecuperoInformazioniFoglio.RecuperoConErrori;
        }


        /// <summary>
        /// Mi dice se ho riconosciuto l'header relativo alle informazioni per le leghe sul primo foglio per il primo 
        /// formato excel
        /// se viene riconosciuto correttamente per l'header delle leghe viene anche restituito come output l'indice per la colonna corrente 
        /// </summary>
        /// <param name="nextColIndex"></param>
        /// <returns></returns>
        private static bool HoRiconosciutoHeader_Format1_Leghe(out int nextColIndex)
        {
            // lista di tutte le proprieta obbligatorie riconosciute per le proprieta correnti
            List<string> recognizedMandatoryProperties = new List<string>();

            // lista di tutte le proprieta che vengono riconosciute per l'iterazione corrente 
            List<string> allRecognizedProperties = new List<string>();

            // tiene traccia delle proprieta che sto leggendo
            nextColIndex = _currentColIndex;
           

            while (!(_foglioExcelCorrente.Cells[_currentRowIndex, nextColIndex].Value == null))
            {
                if (_mandatoryInfo_format1_sheet1.Contains(_foglioExcelCorrente.Cells[_currentRowIndex, nextColIndex].Value.ToString().ToUpper()) && !(recognizedMandatoryProperties.Contains(_foglioExcelCorrente.Cells[_currentRowIndex, nextColIndex].Value.ToString().ToUpper())))
                    recognizedMandatoryProperties.Add(_foglioExcelCorrente.Cells[_currentRowIndex, nextColIndex].Value.ToString());

                // inserimento in tutte le proprieta di header in riconoscimento corrente per eventualmente andare a calcolare i warnings per il caso corrente 
                allRecognizedProperties.Add(_foglioExcelCorrente.Cells[_currentRowIndex, nextColIndex].Value.ToString());

                nextColIndex++;
            }

            // calcolo della stringa relativa ai possibili warnings (le proprieta non obbligatorie che eventualmente non vengono riconosciute sul foglio in lettura corrente)
            CompleteListWarnings_ReadLegheInfo_Format1(allRecognizedProperties);


            if (recognizedMandatoryProperties.Count() == _mandatoryInfo_format1_sheet1.Count())
            {
                return true;
            }

            // completamento della lista degli errori per le possibili proprieta di header di cui è mancata la lettura
            CompleteListError_ReadLegheInfo_Format1(recognizedMandatoryProperties);


            nextColIndex = 0;
            return false;
        }


        /// <summary>
        /// Permette di completare la stringa di errori relativa alla lettura degli headers per il riconoscimento delle proprieta di lega per 
        /// il foglio relativo alle leghe per il primo formato excel disponibile
        /// </summary>
        /// <param name="partialProperties"></param>
        private static void CompleteListError_ReadLegheInfo_Format1(List<string> partialProperties)
        {
            foreach(string mandatoryProperty in _mandatoryInfo_format1_sheet1)
            {
                // se la proprieta non è stata letta per l'header la aggiungo alla lista di tutte le proprieta mancate per il foglio excel correntemente in analisi dal tool
                if (!partialProperties.Contains(mandatoryProperty))
                    _error_Messages_ReadExcel += String.Format(Excel_ErrorMessages.Formato1_Foglio1_Leghe.ERRORE_MANCATORICONOSCIMENTOPROPRIETAHEADERLEGHE, mandatoryProperty, _foglioExcelCorrente.Name);
            }
        }


        /// <summary>
        /// Permette il completamento con la segnalazione di tutti gli headers non obbligatori che non sono comunque riuscito a riconoscere all'interno 
        /// del foglio corrente per le leghe sul primo formato excel disponibile
        /// </summary>
        /// <param name="partialProperties"></param>
        private static void CompleteListWarnings_ReadLegheInfo_Format1(List<string> partialProperties)
        {
            // inserimento della proprieta opzionale di cui è mancata la lettura per gli headers correnti
            foreach (string optionalProperty in _optionalInfo_format1_sheet1)
                _warning_Messages_ReadExcel += String.Format(Excel_WarningMessages.Formato1_Foglio1_Leghe.WARNING_MANCATORICONOSCIMENTOPROPRIETAHEADER_LEGA, optionalProperty, _foglioExcelCorrente.Name);
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
        public static Constants_Excel.EsitoRecuperoInformazioniFoglio Recognize_Format1_InfoConcentrations(
            ref ExcelWorksheet currentWorksheet, 
            out List<Excel_Format1_Sheet2_ConcQuadrant> listaQuadrantiConcentrazioni,
            out string errorRecognizingConcentrationsSheet,
            out string warningsRecognizingConcentrationsSheet)
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


            // inizializzazione delle stringhe relative ai messaggi di errori warnings per la lettura del foglio corrente 
            _error_Messages_ReadExcel = String.Empty;
            _warning_Messages_ReadExcel = String.Empty;
            
            
            do
            {
                _currentColIndex++;

                do
                {
                    _currentRowIndex++;

                    // se ci sono eventuali parti da accodare per il mancato riconosciemento di un quadrante lo vado a fare in questa fase
                    if(!HoRiconosciutoFormat1_Concentrazioni())
                    {
                        _error_Messages_ReadExcel += _finalError_Message_ConcentrationQuadrant;
                        _warning_Messages_ReadExcel += _finalWarning_Message_ConcentrationQuadrant;

                        // azzeramento stringhe di errori warnings per la lettura di quadrante
                        // passaggio a quadrante successivo
                        _finalError_Message_ConcentrationQuadrant = String.Empty;
                        _finalWarning_Message_ConcentrationQuadrant = String.Empty;
                    } 


                }
                while (_currentRowIndex <= currentWorksheet.Dimension.End.Row);

                int _colIndexIterazionePrecedente = _currentColIndex;


                // ricalcolo eventuale indice colonna 
                _currentColIndex = RicalcolaIndiceColonna();

                if (_currentColIndex == 0)
                    break;


                // riazzero indice di riga 
                _currentRowIndex = 0;
            }
            while (_currentColIndex <= currentWorksheet.Dimension.End.Column);

            // attribuzione con gli eventuali quadranti di concentrazione letti
            listaQuadrantiConcentrazioni = _listaQuadrantiConcentrazioni;


            // attribuzione dei valori in uscita per i messaggi di warnings e di errore
            errorRecognizingConcentrationsSheet = _error_Messages_ReadExcel;
            warningsRecognizingConcentrationsSheet = _warning_Messages_ReadExcel;

            // ritorno true solo se ho riconosciuto almeno un quadrante di concentrazioni per il foglio excel corrente 
            if (listaQuadrantiConcentrazioni.Count() > 0)
            {
                // se la lista dei warnings per il riconoscimento corrente è comunque maggiore di 0, devo ritornare il riconoscimento per warnings 
                if (warningsRecognizingConcentrationsSheet != String.Empty)
                    return Constants_Excel.EsitoRecuperoInformazioniFoglio.RecuperoConWarnings;

                return Constants_Excel.EsitoRecuperoInformazioniFoglio.RecuperoCorretto;
            }
            // inserisco un messaggio di errore indicante che la tipologia per il foglio è stata letta correttamente ma che non sono riuscito a leggere nessun quadrante per una analisi
            else
                errorRecognizingConcentrationsSheet += Excel_ErrorMessages.Formato1_Foglio2_Concentrazioni.ERRORE_NESSUNQUADRANTECONCENTRAZIONIUTILEPERANALISI;

            return Constants_Excel.EsitoRecuperoInformazioniFoglio.RecuperoConErrori;

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
            {
                // inserisco nel messaggio di errori il mancato riconoscimento del title per il quadrante corrente 
                _error_Messages_ReadExcel += String.Format(Excel_ErrorMessages.Formato1_Foglio2_Concentrazioni.ERRORE_MANCATOTITLEQUADRANTE, _currentRowIndex);

                return false;
            }
                

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

            while((!riconoscimentoHeader))
            {
                riconoscimentoHeader = RecognizeHeaderConcentrations(out maxColIndex, out _finalError_Message_ConcentrationQuadrant, out _finalWarning_Message_ConcentrationQuadrant);
                if (riconoscimentoHeader)
                {
                    riconoscimentoQuadranteCorrente.StartingRow_Headers = _currentRowIndex;
                    riconoscimentoQuadranteCorrente.EndingCol = maxColIndex;
                    break;
                }
                    

                // incremento per questa iterazione solamente nel caso in cui non abbia ancora riconosciuto l'header corrente di concentrazioni
                _currentRowIndex++;

                if (_currentRowIndex > maxHeader_rowIndex)
                    break;
            }

            // se non ho riconosciuto l'header allora esco senza aver riconosciuto il quadrante 
            if (!riconoscimentoHeader)
            {
                // ricalcolo eventuale indice prima di coontinuare a leggere altro quadrante 
                while (_foglioExcelCorrente.Cells[_currentRowIndex, _currentColIndex].Value != null)
                {

                    if (_foglioExcelCorrente.Cells[_currentRowIndex, _currentColIndex].Merge == true)
                    {
                        _currentRowIndex--;
                        break;
                    }

                    _currentRowIndex++;
                }

                return false;
            }
                

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
                    riconoscimentoQuadranteCorrente.EndingRow_Concentrations = _currentRowIndex - 1;
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
        /// <param name="errorHeaderMessage"></param>
        /// <param name="warningHeaderMessage"></param>
        /// <returns></returns>
        private static bool RecognizeHeaderConcentrations(out int maxColIndex, out string errorHeaderMessage, out string warningHeaderMessage)
        {
            // lista di tutte le proprieta riconosciute
            List<string> recognizedMandatoryProperties = new List<string>();

            // lista di tutti gli elementi not null riconosciuti
            List<string> allRecognizedHeaders = new List<string>();
            

            int currentRowIndexCopy = _currentRowIndex;
            int currentColIndexCopy = _currentColIndex;

            maxColIndex = _currentColIndex;

            if (_foglioExcelCorrente.Cells[currentRowIndexCopy, currentColIndexCopy].Value == null)
            {
                // per questo caso nel quale non si riesce neanche a distinguere una possibile proprieta, non viene calcolato nessun messaggio relativo a errori e warnings 
                errorHeaderMessage = String.Empty;
                warningHeaderMessage = String.Empty;

                return false;
            }
                

            while(_foglioExcelCorrente.Cells[currentRowIndexCopy, currentColIndexCopy].Value != null)
            {
                if(_mandatoryInfo_format1_sheet2.Contains(_foglioExcelCorrente.Cells[currentRowIndexCopy, currentColIndexCopy].Value.ToString().ToUpper()))
                    recognizedMandatoryProperties.Add(_foglioExcelCorrente.Cells[currentRowIndexCopy, currentColIndexCopy].Value.ToString().ToUpper());

                // tengo traccia dell'indice di colonna dei CRITERI per la successiva eventuale lettura degli elementi sottostanti
                if (_foglioExcelCorrente.Cells[currentRowIndexCopy, currentColIndexCopy].Value.ToString().ToUpper() == COLCRITERI_HEADER)
                    _colCriteriIndex = currentColIndexCopy;

                // aggiungo al set di tutti gli headers letti per eventualmente calcolare i messaggi di warnings per il caso corrente 
                allRecognizedHeaders.Add(_foglioExcelCorrente.Cells[currentRowIndexCopy, currentColIndexCopy].Value.ToString().ToUpper());

                // incremento indice di colonna relativo agli headers
                maxColIndex++;
                currentColIndexCopy++;
            }

            // calcolo del messaggio di warning per l'iterazione corrente 
            warningHeaderMessage = CheckHeaderConcentrationsOptionalPropertiesForQuadrant(allRecognizedHeaders);

            if (recognizedMandatoryProperties.Count() == _mandatoryInfo_format1_sheet2.Count())
            {
                errorHeaderMessage = String.Empty;
                return true;
            }

            // calcolo del messaggio di errore per l'iterazione corrente 
            errorHeaderMessage = CheckHeaderConcentrationsMandatoryPropertiesForQuadrant(recognizedMandatoryProperties);
            
            return false;
        }


        /// <summary>
        /// Permette di stabilire e inserire all'interno del messaggio di errore per il quadrante corrente quali siano state le proprieta non riconosciute 
        /// per il quadrante concentrazioni corrente 
        /// </summary>
        /// <param name="partialProperties"></param>
        private static string CheckHeaderConcentrationsMandatoryPropertiesForQuadrant(List<string> partialProperties)
        {
            // messaggio finale di errore dato per la lettura delle proprieta per le concentrazioni e sulla riga corrente 
            string finalEventualErrors = String.Empty;

            foreach (string mandatoryProperty in Constants_Excel.PROPRIETAOBBLIGATORIE_FORMAT1_SHEET2)
                if (!partialProperties.Contains(mandatoryProperty))
                    finalEventualErrors = String.Format(Excel_ErrorMessages.Formato1_Foglio2_Concentrazioni.ERRORE_MANCATORICONOSCIMENTOPROPRIETAHEADEROBBLIGATORIA, _currentRowIndex, mandatoryProperty);

            return finalEventualErrors;
        }


        /// <summary>
        /// Permette di stabilire e inserire all'interno del messaggio di warnings per il quadrante corrente quali siano state le proprieta non obbligatorie non riconosciute 
        /// per il quadrante concentrazioni corrente 
        /// </summary>
        /// <param name="readProperties"></param>
        private static string CheckHeaderConcentrationsOptionalPropertiesForQuadrant(List<string> readProperties)
        {
            // messaggio finale di warning dato per la lettura delle proprieta opzionali per le concentrazioni e sulla riga corrente 
            string finalEventualWarnings = String.Empty;

            foreach (string optionalProperty in Constants_Excel.PROPRIETAOPZIONALI_FORMAT1_SHEET2)
                if (!readProperties.Contains(optionalProperty))
                    finalEventualWarnings += String.Format(Excel_WarningMessages.Formato1_Foglio2_Concentrazioni.WARNING_MANCATORICONOSCIMENTOPROPRIETAOPZIONALIQUADRANTE, _currentRowIndex, optionalProperty);

            return finalEventualWarnings;

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
                else
                {
                    if (hoLettoAlmenoUnPossibileValoreElemento)
                    {
                        _currentRowIndex--;
                        return true;
                    }
                        
                }
                   
            }

            if (hoLettoAlmenoUnPossibileValoreElemento)
                return true;

            // segnalo che non sono riuscito a riconoscere nessun elemento per il quadrante delle concentrazioni corrente 
            _error_Messages_ReadExcel += String.Format(Excel_ErrorMessages.Formato1_Foglio2_Concentrazioni.ERRORE_NESSUNRICONOSCIMENTOPERELEMENTO, _currentRowIndex);
            return false;
        }


        /// <summary>
        /// Permette di calcolare l'indice per il riposizionamento eventuale della colonna per il riconoscimento
        /// di altri quadranti all'interno del foglio excel delle concentrazioni
        /// </summary>
        /// <returns></returns>
        private static int RicalcolaIndiceColonna()
        {
            int newColIndex = _currentColIndex;

            if (_listaQuadrantiConcentrazioni != null)
                if (_listaQuadrantiConcentrazioni.Count() > 0)
                {
                    newColIndex = _listaQuadrantiConcentrazioni.Select(x => x.EndingCol).Max();
                    if (newColIndex == _currentColIndex - 1)
                        return 0;
                }
                    
            

            return newColIndex;
        }

        
        /// <summary>
        /// 
        /// </summary>
        /// <param name="currentWorksheet"></param>
        /// <param name="startingRow"></param>
        /// <param name="leghe_start_col"></param>
        /// <param name="colonneElementi"></param>
        /// <returns></returns>
        public static bool Recognize_Format2_InfoLegheConcentrazioni(ref ExcelWorksheet currentWorksheet, out int startingRow, out int leghe_start_col, out List<Excel_Format2_ConcColumns> colonneElementi)
        {
            // validazioni di partenza 
            if (currentWorksheet == null)
                throw new Exception(ExceptionMessages.EXCEL_FILENOTINMEMORY);

            _foglioExcelCorrente = currentWorksheet;
            
            int indexRow_Max = 1;
            int intexCol_Max = 1;

            _currentRowIndex = 0;
            _currentColIndex = 0;

            startingRow = 0;
            leghe_start_col = 0;
            colonneElementi = null;

            // inserimento dei valori per il limite massimo di riga / colonna entro il quale devo riconoscere l'informazione 
            indexRow_Max = (currentWorksheet.Dimension.End.Row <= LIMIT_ROW) ? currentWorksheet.Dimension.End.Row : LIMIT_ROW;
            intexCol_Max = (currentWorksheet.Dimension.End.Column <= LIMIT_COL) ? currentWorksheet.Dimension.End.Column : LIMIT_COL;

            do
            {
                _currentColIndex++;

                do
                {
                    _currentRowIndex++;

                    // riconoscimento delle prime informazioni di lega 
                    if (RiconosciQuadrante_SecondoFormato())
                    {
                        // attribuzione dei primi parametri conosciuti
                        startingRow = _startReadingProperties_row;
                        leghe_start_col = _startReadingLegheProperties_col;

                        if (RiconoscimentoColonneConcentrazioni_SecondoFormato())
                        {
                            colonneElementi = _currentConcentrations;
                            return true;
                        }
                    }

                }
                while (_currentRowIndex <= indexRow_Max);

            }
            while (_currentColIndex <= intexCol_Max);
            

            return false;
        }


        /// <summary>
        /// Riconoscimento del quadrante complessivo per il secondo formato excel 
        /// </summary>
        /// <param name="currentWorksheet"></param>
        /// <returns></returns>
        private static bool RiconosciQuadrante_SecondoFormato()
        {
            List<string> readMandatoryProperties_Leghe = new List<string>();

            int currentColCopy = _currentColIndex;

            while(readMandatoryProperties_Leghe.Count() < _mandatoryInfo_Leghe_Format2.Count())
            {

                if (_foglioExcelCorrente.Cells[_currentRowIndex, _currentColIndex].Value != null)
                {
                    if (_mandatoryInfo_Leghe_Format2.Contains(_foglioExcelCorrente.Cells[_currentRowIndex, _currentColIndex].Value.ToString().ToUpper()) && _foglioExcelCorrente.Cells[_currentRowIndex, _colCriteriIndex + 1].Merge == true)
                        readMandatoryProperties_Leghe.Add(_foglioExcelCorrente.Cells[_currentRowIndex, _currentColIndex].Value.ToString().ToUpper());

                    _currentColIndex++;
                }
                else
                    break;
            }

            if(readMandatoryProperties_Leghe.Count() == _mandatoryInfo_Leghe_Format2.Count())
            {
                _startReadingProperties_row = _currentRowIndex;
                _startReadingLegheProperties_col = currentColCopy;

                return true;
            }
            
            return false;
        }


        /// <summary>
        /// Permette di riconoscere le colonne relative alle concentrazioni per il secondo formato 
        /// </summary>
        /// <returns></returns>
        private static bool RiconoscimentoColonneConcentrazioni_SecondoFormato()
        {
            // header nel quale si trovano le eventuali proprieta in lettura per gli elementi
            int rowHeadersProperty = _currentRowIndex + 1;

            // istanza del quadrante di concentrazioni sul quale andare a inserire le proprieta lette 
            Excel_Format2_ConcColumns currentColumnsConcentrations = new Excel_Format2_ConcColumns();

            // non posso continuare l'analisi
            if (_foglioExcelCorrente.Cells[_currentRowIndex, _currentColIndex].Value == null)
                return false;

            // definizione per il primo elemento
            string currentElem = _foglioExcelCorrente.Cells[_currentRowIndex, _currentColIndex].Value.ToString();

            string nextElem = String.Empty;

            if (_foglioExcelCorrente.Cells[_currentRowIndex, _currentColIndex + 1].Value != null)
                nextElem = _foglioExcelCorrente.Cells[_currentRowIndex, _currentColIndex + 1].Value.ToString();

            // lista che alla fine conterrà tutte le proprieta obbligatorie per l'elemento
            List<string> readMandatoryProperties = new List<string>();
            // lista che alla fine conterrà tutte le proprieta opzionali per l'elemento
            List<string> readOptionalProperties = new List<string>();

            // indice di colonna di inizio per il primo elemento in lettura eventuale corrente
            int startingColIndex = _currentColIndex;

            // indice di colonna di fine iterazione per il primo elemento in lettura 
            int endingColIndex = _currentColIndex;
            
            // significa che per questo caso mi trovo ancora nella lettura delle proprieta per l'elemento precedente
            while(_currentColIndex <= _foglioExcelCorrente.Dimension.End.Column)
            {
                if (_foglioExcelCorrente.Cells[_currentRowIndex + 1, _currentColIndex].Value != null)
                {
                    // riconoscimento della proprieta corrente
                    if (
                        _mandatoryInfo_Concentrations_Format2.Contains(_foglioExcelCorrente.Cells[_currentRowIndex + 1, _currentColIndex].Value.ToString().ToUpper()) ||
                        _mandatoryInfo_Concentrations_Format2.Contains(_foglioExcelCorrente.Cells[_currentRowIndex + 1, _currentColIndex].Value.ToString().ToUpper() + "."))
                    {
                        readMandatoryProperties.Add(_foglioExcelCorrente.Cells[_currentRowIndex + 1, _currentColIndex].Value.ToString().ToUpper());
                    }
                    else if (_optionalInfo_Concentrations_Format2.Contains(_foglioExcelCorrente.Cells[_currentRowIndex + 1, _currentColIndex].Value.ToString().ToUpper()))
                    {
                        readOptionalProperties.Add(_foglioExcelCorrente.Cells[_currentRowIndex + 1, _currentColIndex].Value.ToString().ToUpper());
                    }

                    // incremenento indice di colonna per iterazione corrente
                    _currentColIndex++;

                    // prendo la posizione per l'indice di colonna letto in modo finale per l'elemento corrente e le sue proprieta 
                    endingColIndex = _currentColIndex - 1;

                    // verifico che la prossima colonna contiene ancora la definizione per l'elemento corrente 
                    // (deve corrispondere ad un valore NULL per l'elemento corrente)
                    if (_foglioExcelCorrente.Cells[_currentRowIndex, _currentColIndex].Value != null)
                        nextElem = _foglioExcelCorrente.Cells[_currentRowIndex, _currentColIndex].Value.ToString();

                    // finisco di leggere l'elemento corrente solo se ho trovato l'elemento successivos
                    if(nextElem != String.Empty)
                    {
                        if (nextElem != currentElem)
                        {

                            // validazione delle proprieta lette fino ad ora per l'elemento corrente
                            if (readMandatoryProperties.Count() < _mandatoryInfo_Concentrations_Format2.Count())
                                return false;

                            // verifica che l'elemento corrente non sia nelle definizioni già date per gli elementi già inseriti
                            if (_currentConcentrations != null)
                            {
                                if (_currentConcentrations.Where(x => x.NomeElemento == currentElem).Count() > 0)
                                    return false;

                                // inserimento degli indici per il riconoscimento dell'elemento corrente
                                _currentConcentrations.Add(new Excel_Format2_ConcColumns()
                                {
                                    startingCol_Header = startingColIndex,
                                    endingCol_Header = endingColIndex,
                                    startingRow_Elemento = _currentRowIndex,
                                }


                                    );


                                // calcolo eventualmente i nuovi indici di inzio e fine colonna di lettura
                                startingColIndex = endingColIndex;
                                endingColIndex = _currentColIndex;
                            }
                            else
                            {
                                _currentConcentrations = new List<Excel_Format2_ConcColumns>();

                                // inserimento degli indici per il riconoscimento dell'elemento corrente
                                _currentConcentrations.Add(new Excel_Format2_ConcColumns()
                                {
                                    startingCol_Header = startingColIndex,
                                    endingCol_Header = endingColIndex,
                                    startingRow_Elemento = _currentRowIndex,
                                }


                                    );


                                // calcolo eventualmente i nuovi indici di inzio e fine colonna di lettura
                                startingColIndex = endingColIndex;
                                endingColIndex = _currentColIndex;
                            }

                            // imposto il prossimo elemento come l'elemento corrente di analisi
                            currentElem = nextElem;
                        }
                    }

                }
                else
                {
                    _currentColIndex++;
                    endingColIndex = _currentColIndex;
                }

                // capisco se sono comunque arrivato a fine lettura per l'elemento corrente
                // facendo la differenza tra la colonna massima di lettura e quella minima per le proprieta questo valore deve essere al massimo uguale al conteggio delle proprieta sulle 2 liste
                int differenzaProprietaLette = endingColIndex - startingColIndex;
                if (differenzaProprietaLette > (_mandatoryInfo_Concentrations_Format2.Count() + _optionalInfo_Concentrations_Format2.Count()))
                    break;


            }

            if (_currentConcentrations != null)
                if (_currentConcentrations.Count() > 0)
                    return true;

            return false;
        }
    }
}
