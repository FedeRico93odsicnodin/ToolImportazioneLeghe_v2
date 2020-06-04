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
        
        /// <summary>
        /// Attribuzione al momenti di richiamo di uno dei diversi metodi in analisi, mappatura di tutte le informazioni per il foglio 
        /// excel correntemente in analisi
        /// </summary>
        private static ExcelWorksheet _foglioExcelCorrente;
        

        /// <summary>
        /// Traccia di riga correntemenete in analisi
        /// </summary>
        private static int _currentRowIndex = 0;


        /// <summary>
        /// Traccia di colonna correntemente in analisi
        /// </summary>
        private static int _currentColIndex = 0;
        

        /// <summary>
        /// Mapper per gli indici di colonna relativi ai diversi titoli letti per il riconoscimento di 
        /// un determinato header - riconoscimento delle proprieta obbligatorie
        /// </summary>
        private static Dictionary<int, string> _mandatoryTitlesColumnMapper;


        /// <summary>
        /// Mapper per gli indici di colonna relativi ai diversi titoli letti per il riconoscimento di 
        /// un determinato header - riconoscimento delle proprieta opzionali
        /// </summary>
        private static Dictionary<int, string> _optionalTitlesColumnMapper;


        /// <summary>
        /// Indica l'indice massimo di colonna dal quale provare eventualmente a iterare per il riconoscimento dei quadranti di concentrazione
        /// orizzontali successivi rispetto a quelli incolonnati il cui riconoscimento è ultimato
        /// </summary>
        private static int _maxColIndexPreviousQuadrantsRecognition = 0;
        

        /// <summary>
        /// Stringa dei possibili errori emersi durante la lettura per il file excel corrente 
        /// </summary>
        private static string _error_Messages_ReadExcel = String.Empty;


        /// <summary>
        /// Stringa dei possibili warnings durante la procedura di lettura per il file excel corrente 
        /// </summary>
        private static string _warning_Messages_ReadExcel = String.Empty;
        

        /// <summary>
        /// Limite sul numero di righe consentite in lettura vuota prima del riconoscimento dell'header per il quadrante delle concentrazioni
        /// in analisi corrente (lettura concentrazioni per il primo formato disponibile)
        /// </summary>
        private const int _LIMITHEADERCONCENTRATIONRECOGNITION = 5;
        
        
        /// <summary>
        /// Oggetto per la mappatura relativa alle proprieta degli elementi letti per il secondo formato e la terza tipologia di foglio 
        /// </summary>
        public class MapperElementFormat2
        {
            /// <summary>
            /// Nome per l'elemento (intestazione a inizio colonna)
            /// </summary>
            public string NameElement { get; set; } 


            /// <summary>
            /// Indicazione di riga per l'istanza di proprieta per le concentrazioni in lettura corrente 
            /// </summary>
            public int CurrentRowIndex { get; set; }


            /// <summary>
            /// Mappatura delle proprieta obbligatorie e in particolare degli indici 
            /// di colonna e relativi titoli delle proprieta per questi
            /// </summary>
            public Dictionary<int, string> MandatoryProperties { get; set; }


            /// <summary>
            /// Mappatura delle proprieta opzionali e in particolare degli indici
            /// di colonna e relativi titoli delle proprieta per questi 
            /// </summary>
            public Dictionary<int, string> OptionalProperties { get; set; }
        }


        /// <summary>
        /// Lista relativa alla mappatura degli header per le proprieta degli elementi per il terzo formato a disposizione
        /// </summary>
        private static List<MapperElementFormat2> _mapperElementsFormat2Sheet3 { get; set; }

        #endregion
        

        /// <summary>
        /// Riempimento di tutti gli oggetti per il foglio 1 e il formato 1 correnti
        /// questo foglio viene cosi anche eventualmente validato rispetto alle informazioni di lega che si dovranno ritrovare al suo interno
        /// </summary>
        /// <param name="currentWorksheet"></param>
        /// <param name="recognizedInfo"></param>
        /// <param name="possibleErrorMessages"></param>
        /// <param name="possibleWarningMessages"></param>
        /// <returns></returns>
        public static Constants_Excel.EsitoRecuperoInformazioniFoglio Recognize_Format1_InfoLeghe(
            ref ExcelWorksheet currentWorksheet, 
            out Excel_AlloyInfo_Sheet recognizedInfo)
        {

            // parametri fondamentali per il riconoscimento corrente 
            _foglioExcelCorrente = currentWorksheet;
            _currentRowIndex = 0;
            _currentColIndex = 0;
            
            // inizializzazione delle 2 stringhe relative ai messaggi di segnalazione warnings / errori per il file excel correntemente in analisi
            _error_Messages_ReadExcel = String.Empty;
            _warning_Messages_ReadExcel = String.Empty;


            // istanza per il foglio corrente 
            recognizedInfo = new Excel_AlloyInfo_Sheet(_foglioExcelCorrente.Name, _foglioExcelCorrente.Index);


            // riconoscimento header
            bool riconoscimentoHeaderCorrente = false;

            #region RICONOSCIMENTO HEADER PER IL FOGLIO DELLE LEGHE CORRENTE 

            do
            {
                // riconoscimento header per il formato corrente 
                _currentColIndex++;

                // azzero nuovamente indice di riga per la nuova iterazione su colonna
                _currentRowIndex = 0;

                do
                {
                    _currentRowIndex++;

                    // tentativo di riconoscimento per l'header corrente 
                    if(RecognizeHeaderOnRow(Constants_Excel.PROPRIETAOBBLIGATORIE_FORMAT1_SHEET1.ToList(), Constants_Excel.PROPRIETAOPZIONALI_FORMAT1_SHEET1.ToList()))
                    {
                        riconoscimentoHeaderCorrente = true;
                        break;
                    }

                }
                while (_currentRowIndex <= _foglioExcelCorrente.Dimension.End.Row);

                // se ho riconosciuto gli headers agli steps precedenti esco dal ciclo
                if (riconoscimentoHeaderCorrente)
                    break;
            }
            while (_currentColIndex <= _foglioExcelCorrente.Dimension.End.Column);

            #endregion


            #region RIEMPIMENTO OGGETTI CON PROPRIETA

            if(riconoscimentoHeaderCorrente)
            {
                // istanza di tutte le proprieta di lega riconosciute per il foglio corrente 
                List<Excel_PropertyWrapper> allPropertiesLeghe = new List<Excel_PropertyWrapper>();

                _currentRowIndex++;

                while(_currentRowIndex <= _foglioExcelCorrente.Dimension.End.Row)
                {
                    // istanza proprieta di lega per la riga corrente 
                    List<Excel_PropertyWrapper> allPropertiesLega_currentRow;

                    InsertNewRowPossibleValue(out allPropertiesLega_currentRow);

                    // aggiunta alle proprieta globali di tutto il foglio corrente 
                    if (allPropertiesLega_currentRow.Count() > 0)
                        allPropertiesLeghe.AddRange(allPropertiesLega_currentRow);

                    _currentRowIndex++;
                }

                // aggiunta dell'eventuale messaggio di errore per il fatto di non aver riconosciuto correttamente tutte le proprieta di lega 
                if (allPropertiesLeghe.Count() == 0)
                {
                    _error_Messages_ReadExcel += String.Format("non ho riconosciuto nessuna proprieta di lega per il foglio '{0}' in posizione {1}", _foglioExcelCorrente.Name, _foglioExcelCorrente.Index);
                    recognizedInfo.Validation_OK = false;
                }
                    
                // altrimenti imposto la proprieta per il foglio in uscita e le leghe lette 
                else
                {
                    // costruzione dell'oggetto relativo alle proprieta riconosciute
                    List<Excel_PropertiesContainer> groupedProperties = BuildPropertiesContainerForRecognition_Format1Sheet1_Leghe(allPropertiesLeghe);

                    // inserimento nel foglio custom in restituzione
                    recognizedInfo.AlloyPropertiesInstances = groupedProperties;


                    // attribuzione delle 2 stinghe per i messaggi di errore e warnings finali
                    recognizedInfo.ErrorMessageSheet = _error_Messages_ReadExcel;
                    recognizedInfo.WarningMessageSheet = _warning_Messages_ReadExcel;

                    // se ho dei warnings ritorno con warnings 
                    if (recognizedInfo.WarningMessageSheet != String.Empty)
                        return Constants_Excel.EsitoRecuperoInformazioniFoglio.RecuperoConWarnings;

                    return Constants_Excel.EsitoRecuperoInformazioniFoglio.RecuperoCorretto;

                }

            }

            #endregion

            // attribuzione delle 2 stinghe per i messaggi di errore e warnings finali
            recognizedInfo.ErrorMessageSheet = _error_Messages_ReadExcel;
            recognizedInfo.WarningMessageSheet = _warning_Messages_ReadExcel;

            return Constants_Excel.EsitoRecuperoInformazioniFoglio.RecuperoConErrori;
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
            out Excel_AlloyInfo_Sheet recognizedInfoConcentrations)
        {

            _foglioExcelCorrente = currentWorksheet;
            _currentRowIndex= 0;
            _currentColIndex = 0;


            // inizializzazione delle stringhe relative ai messaggi di errori warnings per la lettura del foglio corrente 
            _error_Messages_ReadExcel = String.Empty;
            _warning_Messages_ReadExcel = String.Empty;


            // inizializzazione per la lettura di tutti i possibili quadranti concentrazioni
            recognizedInfoConcentrations = new Excel_AlloyInfo_Sheet(_foglioExcelCorrente.Name, _foglioExcelCorrente.Index);

            // inizializzazione wrapper per le proprieta di concentrazione
            recognizedInfoConcentrations.ConcentrationsPropertiesInstances = new List<Excel_PropertiesContainer>();


            do
            {
                _currentColIndex++;

                int _maxColPrev = _maxColIndexPreviousQuadrantsRecognition;

                do
                {
                    _currentRowIndex++;

                    Excel_PropertiesContainer possibleQuadrant;

                    // se ci sono eventuali parti da accodare per il mancato riconosciemento di un quadrante lo vado a fare in questa fase
                    if(RecognizeQuadrantConcentration_Format1(out possibleQuadrant))
                    {
                        // aggiunta del quadrante riconosciuto alla lista di tutti i quadranti per il foglio corrente 
                        recognizedInfoConcentrations.ConcentrationsPropertiesInstances.Add(possibleQuadrant);
                    } 

                }
                while (_currentRowIndex <= currentWorksheet.Dimension.End.Row);

                // attribuzione indice di colonna rispetto al quadrante con indice piu grande di colonna rispetto al quale andare a leggere successivamente 
                if (_maxColPrev != _maxColIndexPreviousQuadrantsRecognition)
                {
                    _currentColIndex = _maxColIndexPreviousQuadrantsRecognition;
                }

                _currentRowIndex = 0;
                    
            }
            while (_currentColIndex <= currentWorksheet.Dimension.End.Column);

            // attribuzione dei valori in uscita per i messaggi di warnings e di errore
            recognizedInfoConcentrations.ErrorMessageSheet = _error_Messages_ReadExcel;
            recognizedInfoConcentrations.WarningMessageSheet = _warning_Messages_ReadExcel;

            // verifico quanti quadranti ho letto 
            if (recognizedInfoConcentrations.ConcentrationsPropertiesInstances.Count() == 0)
            {
                _error_Messages_ReadExcel += String.Format("per il foglio '{0}' in posizione {1} non sono riuscito a riconoscere nessun quadrante di concentrazione", _foglioExcelCorrente.Name, _foglioExcelCorrente.Index);

                return Constants_Excel.EsitoRecuperoInformazioniFoglio.RecuperoConErrori;
            }

            if (recognizedInfoConcentrations.WarningMessageSheet != String.Empty)
                return Constants_Excel.EsitoRecuperoInformazioniFoglio.RecuperoConWarnings;

            return Constants_Excel.EsitoRecuperoInformazioniFoglio.RecuperoCorretto;

        }

        
        /// <summary>
        /// 
        /// </summary>
        /// <param name="currentWorksheet"></param>
        /// <param name="recognizedInfoFormat2"></param>
        /// <param name="errorMessages"></param>
        /// <param name="warningMessages"></param>
        /// <returns></returns>
        public static Constants_Excel.EsitoRecuperoInformazioniFoglio Recognize_Format2_InfoLegheConcentrazioni(
            ref ExcelWorksheet currentWorksheet, 
            out Excel_AlloyInfo_Sheet recognizedInfoFormat2)
        {
            // foglio excel corrente 
            _foglioExcelCorrente = currentWorksheet;


            // inizializzazione del foglio in restituzione per il foglio excel corrente 
            recognizedInfoFormat2 = new Excel_AlloyInfo_Sheet(_foglioExcelCorrente.Name, _foglioExcelCorrente.Index);


            // informazioni relative alla valorizzazione dei diversi elementi per le concentrazioni
            List<Excel_ConcColumns_Definition> informazioniColonneConcentrazioniCorrenti = new List<Excel_ConcColumns_Definition>();


            // indici massimo e minimo di iterazione
            int indexRow_Max = 1;
            int intexCol_Max = 1;

            // indici di colonna e di riga su iterazione corrente 
            _currentRowIndex = 0;
            _currentColIndex = 0;
            

            // inizializzazione per le stringhe di messaggi errori / warnings in uscita 
            _error_Messages_ReadExcel = String.Empty;
            _warning_Messages_ReadExcel = String.Empty;

            // riconoscimento header principale
            bool hoRiconosciutoHeaderLeghe = false;
            bool hoRiconosciutoHeaderConcentrazioni = false;


            #region RICONOSCIMENTO HEADERS 

            while (_currentRowIndex <= indexRow_Max) 
            {

                _currentRowIndex++;

                while (_currentColIndex <= intexCol_Max) 
                {

                    int maxIndexConcentrations = 0;

                    // riconoscimento proprieta per header leghe - riconoscimento della colonna dalla quale iniziare a leggere per le concentrazioni
                    if (RecognizeHeaderOnRow(Constants_Excel.PROPRIETAOBBLIGATORIE_FORMAT2_LEGHE.ToList(), Constants_Excel.PROPRIETAOPZIONALI_FORMAT2_LEGHE.ToList(), out maxIndexConcentrations))
                    {
                        _currentColIndex = maxIndexConcentrations;
                        hoRiconosciutoHeaderLeghe = true;
                    }
                    else
                        continue;

                    if (hoRiconosciutoHeaderLeghe && RecognizeColumnsConcentrationsFormat2())
                    {
                        hoRiconosciutoHeaderConcentrazioni = true;
                        break;
                    }
                        

                }

                _currentColIndex = 0;

            }

            #endregion




            // attribuzione messaggi errori / warnings in uscita 
            recognizedInfoFormat2.ErrorMessageSheet = _error_Messages_ReadExcel;
            recognizedInfoFormat2.WarningMessageSheet = _warning_Messages_ReadExcel;


            // inizializzazione dei valori per la lettura di tutte le proprieta di riga 
            if (hoRiconosciutoHeaderLeghe && hoRiconosciutoHeaderConcentrazioni)
            {
                int _contentRowIndex = _currentRowIndex + 2;
                
                // inizializzazione dei contenitori finali per tutte le proprieta di lega - concentrazioni
                recognizedInfoFormat2.AlloyPropertiesInstances = new List<Excel_PropertiesContainer>();
                recognizedInfoFormat2.ConcentrationsPropertiesInstances = new List<Excel_PropertiesContainer>();

                // inizializzazione del contenitore che associa leghe e concentrazioni per i fogli
                recognizedInfoFormat2.AlloyInstances = new List<AlloyConcentrations_Association>();

                while (_contentRowIndex <= _foglioExcelCorrente.Dimension.End.Row)
                {
                    #region INIZIALIZZAZIONE CON PROPRIETA DI LEGA

                    // recupero degli spazi di proprieta per la lega su riga corrente 
                    Excel_PropertiesContainer currentPropertiesAlloy = BuildPropertiesContainerForRecognition_Format2Sheet_Leghe(_contentRowIndex);

                    // recupero degli spazi di proprieta per ogni elemento disponibile per la lega corrente 
                    List<Excel_PropertiesContainer> currentPropertiesConcentrations = BuildPropertiesContainerForRecognition_Format2Sheet_ConcentrationsOnRow(_contentRowIndex);

                    // aggiunta delle possibili proprieta di lega per tutti i contenitori relativi alle leghe 
                    recognizedInfoFormat2.AlloyPropertiesInstances.Add(currentPropertiesAlloy);

                    // aggiunta di tutti i contenitori relativi alle concentrazioni per tutti i raccoglitori di concentrazione presi
                    recognizedInfoFormat2.ConcentrationsPropertiesInstances.AddRange(currentPropertiesConcentrations);

                    // creazione e aggiunta dell'associazione 
                    AlloyConcentrations_Association currentAssociationLegheConcentrazioni = new AlloyConcentrations_Association();
                    currentAssociationLegheConcentrazioni.RifProprietaLega = currentPropertiesAlloy;
                    currentAssociationLegheConcentrazioni.RifConcentrationsCurrentLega = currentPropertiesConcentrations;

                    recognizedInfoFormat2.AlloyInstances.Add(currentAssociationLegheConcentrazioni);
                    
                    #endregion
                }

                if (recognizedInfoFormat2.WarningMessageSheet != String.Empty)
                    return Constants_Excel.EsitoRecuperoInformazioniFoglio.RecuperoConWarnings;

                return Constants_Excel.EsitoRecuperoInformazioniFoglio.RecuperoCorretto;
            }
            // inserimento per i messaggi di errore di nessuna proprieta di riga letta 
            else
                recognizedInfoFormat2.ErrorMessageSheet += String.Format(Excel_ErrorMessages.Formato2_Foglio1_LegheConcentrazioni.ERRORE_NESSUNAINFORMAZIONEDIRIGALETTA, _foglioExcelCorrente.Name);
            

            return Constants_Excel.EsitoRecuperoInformazioniFoglio.RecuperoConErrori;
        }
        

        #region NUOVI METODI

        /// <summary>
        /// Riconoscimento di un quadrante di concentratiozi per il primo formato 
        /// questo riconoscimento avviene a partire dagli indici di riga e colonna correnti tracciano i nuovi indici di riga e colonna nel caso in cui sia 
        /// effettivamente riconosciuto il quadrante 
        ///</summary>
        ///<param name="recognizedConcQuadrant"></param>
        /// <returns></returns>
        private static bool RecognizeQuadrantConcentration_Format1(out Excel_PropertiesContainer recognizedConcQuadrant)
        {
            // inizializzazione del quadrante corrente 
            recognizedConcQuadrant = new Excel_PropertiesContainer();

            // inizializzazione lista delle concentrazioni lette per il quadrante corrente 
            recognizedConcQuadrant.PropertiesDefinition = new List<Excel_PropertyWrapper>();

            // backup indici di riga e colonna corrente, la riga di backup è incrementata di 1
            int backRowIndex = _currentRowIndex;
            int backColIndex = _currentColIndex;

            // indice di colonna massimo da imporre per il quadrante corrente
            int maxColValueCurrentQuadrant = 0;

            // riconoscimenti disponibili
            bool titleRecognized = false;
            bool headerRecognized = false;
            bool filledConcentrations = false;

            // parametri da riempire per il corretto riconoscimento del quadrante 
            string titleMateriale = String.Empty;


            // riconoscimento title 
            if (_foglioExcelCorrente.Cells[_currentRowIndex, _currentColIndex].Value != null)
            {
                titleMateriale = _foglioExcelCorrente.Cells[_currentRowIndex, _currentColIndex].Value.ToString();
                titleRecognized = true;
                _currentRowIndex++;
            }
            // non è avvenuto il corretto riconoscimento per il title
            else
                titleRecognized = false;

            // set limite massimo di righe consentite fino al riconoscimento dell'header per il quadrante delle concentrazioni
            int limitOnHeaderRecognition = _currentRowIndex + _LIMITHEADERCONCENTRATIONRECOGNITION;

            headerRecognized = false;

            // incremento posizione row index
            _currentRowIndex++;

            // riconoscimento header
            while (_currentRowIndex <= limitOnHeaderRecognition)
            {
                if(RecognizeHeaderOnRow(Constants_Excel.PROPRIETAOBBLIGATORIE_FORMAT1_SHEET2.ToList(), Constants_Excel.PROPRIETAOPZIONALI_FORMAT1_SHEET2.ToList(), _currentColIndex, out maxColValueCurrentQuadrant))
                {
                    headerRecognized = true;
                    _currentRowIndex++;
                    break;
                }

                _currentRowIndex++;
            }

            filledConcentrations = false;

            // tutte le proprieta per gli elementi in lettura corrente 
            List<Excel_PropertyWrapper> allElementProperties = new List<Excel_PropertyWrapper>();


            // riconoscimento elementi concentrazioni
            while (_foglioExcelCorrente.Cells[_currentRowIndex, _currentColIndex].Value != null && _foglioExcelCorrente.Cells[_currentRowIndex, _currentColIndex].Merge == false)
            {
                // riempimento per i singoli elementi di concentrazione
                List<Excel_PropertyWrapper> readPropertiesCurrentElement;

                if (InsertNewRowPossibleValue(out readPropertiesCurrentElement))
                    allElementProperties.AddRange(readPropertiesCurrentElement);

                _currentRowIndex++;
            }

            // vedo se ho trovato la definizione di almeno un elemento con le sue proprietà
            if (allElementProperties.Count() > 0)
                filledConcentrations = true;
            
            // se tutti i valori da riconoscere hanno esito negativo significa che non sono in prossimità di un quadrante, ritorno al primo stato
            if(titleRecognized == false && headerRecognized == false && filledConcentrations == false)
            {
                _currentRowIndex = backRowIndex;
                _currentColIndex = backColIndex;

                return false;
            }


            // inserimento dei diversi messaggi di errore per quadrante, nel caso in cui non sia avvenuto anche solo uno dei riconoscimenti
            if (titleRecognized == false)
                _error_Messages_ReadExcel += String.Format("non sono riuscito a trovare il TITLE per il quadrante concentrazioni a partire dalla posizione ({0}, {1})", backRowIndex, backColIndex);

            if (headerRecognized == false)
                _error_Messages_ReadExcel += String.Format("non sono riuscito a trovare HEADER per il quadrante concentrazioni a partire dalla posizione ({0}, {1})", backRowIndex, backColIndex);

            if (filledConcentrations == false)
                _error_Messages_ReadExcel += String.Format("non sono riuscito a riconoscere informazioni di ELEMENTI per il quadrante concentrazioni dalla posizione ({0}, {1})", backRowIndex, backColIndex);

            // se ho riconosciuto tutti gli elementi vado a riempire il contenitore corrente per gli elementi di concentrazione
            if(titleRecognized == true && headerRecognized == true && filledConcentrations == true)
            {
                // attribuzione title
                recognizedConcQuadrant.NameInstance = titleMateriale;

                // attribuzione degli indici perimetro quadrante 
                recognizedConcQuadrant.StartingRowIndex = backRowIndex;
                recognizedConcQuadrant.EndingRowIndex = _currentRowIndex;
                recognizedConcQuadrant.StartColIndex = _currentColIndex;
                recognizedConcQuadrant.EndingColIndex = maxColValueCurrentQuadrant;

                // attribuzione possibili proprieta da leggere per gli elementi e le concentrazioni
                recognizedConcQuadrant.PropertiesDefinition = allElementProperties;

                // calcolo dell'indice corrente massimo di colonna (mi serve per stabilire quale sarà il prossimo indice di colonna per la lettura dei quadranti 
                // in via orizzontale
                if (maxColValueCurrentQuadrant >= _maxColIndexPreviousQuadrantsRecognition)
                    _maxColIndexPreviousQuadrantsRecognition = maxColValueCurrentQuadrant;

                return true;

            }

            _currentRowIndex = backRowIndex;
            _currentColIndex = backColIndex;

            return false;
        }


        /// <summary>
        /// Utilizzo dei 2 metodi successivi con l'accorgimento di avere come valore iniziale per la colonna di riconoscimento header proprio quella relativa al title 
        /// che è stato riconosciuto nello step precedente all'intero del metodo per il riconoscimento di tutti i quadranti
        /// </summary>
        /// <param name="headerMandatoryTitles"></param>
        /// <param name="headerOptionalTitles"></param>
        /// <param name="usedCol"></param>
        /// <param name="currentValueColMax"></param>
        /// <returns></returns>
        private static bool RecognizeHeaderOnRow(List<string> headerMandatoryTitles, List<string> headerOptionalTitles, int usedCol, out int currentValueColMax)
        {
            currentValueColMax = usedCol;

            // discriminazione su indice di colonna utilizzato per questa iterazione
            if (usedCol != _currentColIndex || _foglioExcelCorrente.Cells[_currentRowIndex, usedCol].Value == null)
                return false;

            return RecognizeHeaderOnRow(headerMandatoryTitles, headerOptionalTitles, out currentValueColMax);

        }



        /// <summary>
        /// Riferimento al metodo precedente con il calcolo dell'indice massimo di colonna per il riconoscimento delle proprieta correnti
        /// in questo modo posso calcolare l'indice massimo di colonna da cui riprendere la lettura per i quadranti orizzontali di concentrazione
        /// successivi
        /// </summary>
        /// <param name="headerMandatoryTitles"></param>
        /// <param name="headerOptionalTitles"></param>
        /// <param name="currentValueColMax"></param>
        /// <returns></returns>
        private static bool RecognizeHeaderOnRow(List<string> headerMandatoryTitles, List<string> headerOptionalTitles, out int currentValueColMax)
        {
            if(RecognizeHeaderOnRow(headerMandatoryTitles, headerOptionalTitles))
            {

                // calcolo dell'indice massimo relativo alle proprieta obbligatorie
                int maxColMandatoryProperties = _mandatoryTitlesColumnMapper.Select(x => x.Key).Max();
                int maxColOptionalProperties = _optionalTitlesColumnMapper.Select(x => x.Key).Max();

                if (maxColMandatoryProperties > maxColOptionalProperties)
                    currentValueColMax = maxColMandatoryProperties;
                else
                    currentValueColMax = maxColOptionalProperties;

                return true;
            }

            currentValueColMax = 0;


            return false;

        } 



        /// <summary>
        /// Mi dice se il riconoscimento degli header per gli elementi che vengono passati in input avviene correttamente per l'istanza corrente 
        /// durante questa fase c'è anche la compilazione dei messaggi di errori / warnings in uscita per il caso corrente 
        /// </summary>
        /// <param name="headerMandatoryTitles"></param>
        /// <param name="headerOptionalTitles"></param>
        /// <returns></returns>
        private static bool RecognizeHeaderOnRow(List<string> headerMandatoryTitles, List<string> headerOptionalTitles)
        {
            // inizializzazione del mapper di colonne per la lettura dei titoli di header correnti
            // header relativi alle proprieta opzionali e obbligatorie in lettura corrente 
            _mandatoryTitlesColumnMapper = new Dictionary<int, string>();
            _optionalTitlesColumnMapper = new Dictionary<int, string>();

            // indicazione del riconoscimento finale per gli elementi obbligatori / opzionali
            bool mandatoryRecognition = false;
            bool optionalRecognition = false;

            // iterazione per la colonna corrente a partire dall'indice di colonna correntemente in lettura
            // questo mi permette di non andare a toccare l'indice di colonna effettivo
            int colIndex = _currentColIndex;
            

            // iterazione fino ad eventualmente raggiungere l'indice finale di colonna 
            while(colIndex <= _foglioExcelCorrente.Dimension.End.Column)
            {
                if(_foglioExcelCorrente.Cells[_currentRowIndex, colIndex].Value != null)
                {
                    string currentValue = _foglioExcelCorrente.Cells[_currentRowIndex, colIndex].Value.ToString().ToUpper();

                    // verifica di contenimento per una proprieta obbligatoria
                    if (headerMandatoryTitles.Contains(currentValue) && !_mandatoryTitlesColumnMapper.ContainsValue(currentValue))
                        _mandatoryTitlesColumnMapper.Add(colIndex, currentValue);
                    // verifica di contenimento per la proprieta opzionale
                    else if (headerOptionalTitles.Contains(currentValue))
                        _optionalTitlesColumnMapper.Add(colIndex, currentValue);

                    // controllo di aver letto correttamente tutti gli headers -> proprieta obbligatorie
                    if (_mandatoryTitlesColumnMapper.Count() == headerMandatoryTitles.Count())
                        mandatoryRecognition = true;

                    if (_optionalTitlesColumnMapper.Count() == headerOptionalTitles.Count())
                        optionalRecognition = true;

                }


                // se ho letto tutte le proprieta disponibili allora esco dal ciclo corrente
                if (mandatoryRecognition && optionalRecognition)
                    break;

                colIndex++;
            }


            // se non ho letto correttamente tutte le proprieta vado a comporre il messaggio di errore / warnings
            // NB il riconoscimento avviene solamente se ho letto almeno una per queste proprietà
            if (!mandatoryRecognition && _mandatoryTitlesColumnMapper.Count() > 0)
                CompileErrorMessage(headerMandatoryTitles, _mandatoryTitlesColumnMapper);

            if (!optionalRecognition && _optionalTitlesColumnMapper.Count() > 0)
                CompileWarningMessage(headerOptionalTitles, _optionalTitlesColumnMapper);
            
            // l'header viene riconosciuto solamente per gli elementi obbligatori che vengono prelevati
            return mandatoryRecognition;
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="startingColIndex"></param>
        /// <returns></returns>
        private static bool RecognizeColumnsConcentrationsFormat2()
        {
            // lista dei mappers per le concentrazioni correnti
            _mapperElementsFormat2Sheet3 = new List<MapperElementFormat2>();

            // indice per la lettura dell'intestazione elemento
            int _rowElemento = _currentRowIndex;

            // indice per la lettura delle proprieta 
            int _rowProperties = _currentRowIndex + 1;

            // attributi per la lettura delle proprieta degli elementi
            string elementoCorrente = String.Empty;
            Dictionary<int, string> columnsMandatoryProperties = new Dictionary<int, string>();
            Dictionary<int, string> columnsOptionalProperties = new Dictionary<int, string>();
            

            // tutte le proprieta per l'elemento corrente 
            MapperElementFormat2 propertiesElementoCorrente;

            // non lettura corretta per anche solo un elemento di tutte le proprieta obbligatorie
            bool missedMandatoryProperties = false;

            while (_currentColIndex <= _foglioExcelCorrente.Dimension.End.Column)
            {

                // prima iterazione
                if(elementoCorrente == String.Empty)
                {
                    // elemento corrente 
                    if (_foglioExcelCorrente.Cells[_rowElemento, _currentColIndex].Value != null)
                        elementoCorrente = _foglioExcelCorrente.Cells[_rowElemento, _currentColIndex].Value.ToString();

                    // inizializzazione dizionari lettura colonne proprieta opzionali obbligatorie
                    columnsMandatoryProperties = new Dictionary<int, string>();
                    columnsOptionalProperties = new Dictionary<int, string>();
                   
                }
                // devo cambiare elemento prima di continuare con iterazione corrente 
                else if(_foglioExcelCorrente.Cells[_rowElemento, _currentColIndex].Value != null)
                {
                    propertiesElementoCorrente = new MapperElementFormat2();

                    propertiesElementoCorrente.NameElement = elementoCorrente;
                    propertiesElementoCorrente.CurrentRowIndex = _rowElemento;

                    // verifica lettura di tutte le proprieta obbligatorie per l'elemento corrente 
                    if (!CompileErrorMessage(Constants_Excel.PROPRIETAOBBLIGATORIE_ELEM_FORMAT2.ToList(), columnsMandatoryProperties))
                        missedMandatoryProperties = true;

                    // compilazione eventuale messaggio warning verifica lettura di tutte le proprieta opzionali
                    CompileWarningMessage(Constants_Excel.PROPRIETAOPZIONALI_ELEM_FORMAT2.ToList(), columnsOptionalProperties);

                    propertiesElementoCorrente.MandatoryProperties = columnsMandatoryProperties;
                    propertiesElementoCorrente.OptionalProperties = columnsOptionalProperties;

                    // inizializzazione per continuare iterazione corrente 
                    elementoCorrente = _foglioExcelCorrente.Cells[_rowElemento, _currentColIndex].Value.ToString();

                    columnsMandatoryProperties = new Dictionary<int, string>();
                    columnsOptionalProperties = new Dictionary<int, string>();
                }

                // lettura proprieta iterazione corrente 
                if(_foglioExcelCorrente.Cells[_rowProperties, _currentColIndex].Value != null)
                {
                    string proprietaLettura = _foglioExcelCorrente.Cells[_rowProperties, _currentColIndex].Value.ToString();

                    if (Constants_Excel.PROPRIETAOBBLIGATORIE_ELEM_FORMAT2.Contains(proprietaLettura.ToUpper()))
                        columnsMandatoryProperties.Add(_currentColIndex, proprietaLettura.ToUpper());

                    else if (Constants_Excel.PROPRIETAOPZIONALI_ELEM_FORMAT2.Contains(proprietaLettura.ToUpper()))
                        columnsOptionalProperties.Add(_currentColIndex, proprietaLettura.ToUpper());
                }

                _currentColIndex++;
            }


            // ritorno false se anche per uno solo degli elementi non ho letto correttamente tutte le definizioni per le proprieta obbligatorie
            if (missedMandatoryProperties)
                return false;

            return true;

        }



        /// <summary>
        /// Recupero dei valori per tutte le definizioni ricorrenti di riga per gli headers letti nello step precedente 
        /// l'unico motivo per il quale una riga non viene aggiunta all'interno del set è se tutte le proprieta sono individuate a null
        /// per la riga corrente 
        /// </summary>
        /// <param name="recognizedValues"></param>
        /// <returns></returns>
        private static bool InsertNewRowPossibleValue(out List<Excel_PropertyWrapper> recognizedValues)
        {
            // inizializzazione dell'oggetto con tutte le proprieta correnti
            recognizedValues = new List<Excel_PropertyWrapper>();
            
            // iterazione per le proprieta obbligatorie
            foreach(KeyValuePair<int, string> currentHeaderCol in _mandatoryTitlesColumnMapper)
            {
                if(_foglioExcelCorrente.Cells[_currentRowIndex, currentHeaderCol.Key].Value != null)
                {
                    recognizedValues.Add(
                        new Excel_PropertyWrapper(_currentRowIndex, currentHeaderCol.Key, currentHeaderCol.Value, false)
                        );
                }
            }

            // iterazione per le proprieta opzionali
            foreach(KeyValuePair<int, string> currentHeaderCol in _optionalTitlesColumnMapper)
            {
                if(_foglioExcelCorrente.Cells[_currentRowIndex, currentHeaderCol.Key].Value != null)
                {
                    recognizedValues.Add(
                        new Excel_PropertyWrapper(_currentRowIndex, currentHeaderCol.Key, currentHeaderCol.Value, true)
                        );
                }
            }
            
            if(recognizedValues.Count() == 0)
            {
                _warning_Messages_ReadExcel += "nessuna proprieta riconosciuta per la riga {0}" + _currentRowIndex;
                return false;
            }

            return true;
        }


        /// <summary>
        /// Permette di costruire i containers delle proprieta suddividi per indici di riga o di colonna in base all'iterazione foglio corrente 
        /// </summary>
        /// <param name="recognizedProperties"></param>
        /// <returns></returns>
        private static List<Excel_PropertiesContainer> BuildPropertiesContainerForRecognition_Format1Sheet1_Leghe(List<Excel_PropertyWrapper> recognizedProperties)
        {
            // ordimento della lista di partenza delle proprieta di riga in base all'indice di riga corrente 
            recognizedProperties = recognizedProperties.OrderBy(x => x.Row_Position).ToList();

            // indice di riga corrente in iterazione
            int rowIndexForProperty = 0;

            // definizione di tutti i contenitori definiti
            List<Excel_PropertiesContainer> allDefinedPropertiesContainer = new List<Excel_PropertiesContainer>();
            // definizione contenitore corrente 
            Excel_PropertiesContainer currentPropertyContainer = new Excel_PropertiesContainer();
            currentPropertyContainer.PropertiesDefinition = new List<Excel_PropertyWrapper>();

            // iterazione e inserimento nello stesso properties container in base alla definizione di proprieta di lega per il primo foglio e la tipologia di excel 1
            foreach(Excel_PropertyWrapper currentPropertyInstance in recognizedProperties)
            {
                // primo oggetto iterato
                if(currentPropertyInstance == recognizedProperties.First())
                {
                    // prima riga di iterazione
                    rowIndexForProperty = currentPropertyInstance.Row_Position;

                    currentPropertyContainer.PropertiesDefinition.Add(
                        currentPropertyInstance
                        );
                }


                // definizione di nuovo oggetto
                else if(rowIndexForProperty != currentPropertyInstance.Row_Position)
                {
                    // se il contenitore nella versione precedente contiene la definizione di proprieta lo aggiungo a tutti i contenitori definiti
                    if (currentPropertyContainer.PropertiesDefinition.Count() > 0)
                        allDefinedPropertiesContainer.Add(currentPropertyContainer);

                    currentPropertyContainer = new Excel_PropertiesContainer();
                    currentPropertyContainer.PropertiesDefinition = new List<Excel_PropertyWrapper>();

                    // aggiunta per la riga corrente (per questa tipologia è la stessa di inizio e fine 
                    currentPropertyContainer.StartingRowIndex = currentPropertyInstance.Row_Position;
                    currentPropertyContainer.EndingRowIndex = currentPropertyInstance.Row_Position;


                    // aggiunta della proprieta per la riga corrente 
                    currentPropertyContainer.PropertiesDefinition.Add(
                        currentPropertyInstance
                        );


                    // nuova riga iterazione 
                    rowIndexForProperty = currentPropertyInstance.Row_Position;
                }
                // aggiungo la proprieta di riga al contenitore già istanziato
                else
                    currentPropertyContainer.PropertiesDefinition.Add(
                        currentPropertyInstance
                        );

                // aggiungo anche l'ultimo elemento al container finale
                if(currentPropertyInstance == recognizedProperties.Last() && currentPropertyContainer.PropertiesDefinition.Count() > 0)
                    allDefinedPropertiesContainer.Add(currentPropertyContainer);


            }

            return allDefinedPropertiesContainer;
        }


        /// <summary>
        /// Adempimento di tutte le proprieta obbligatorie e opzionali per la lega sulla riga corrente 
        /// (foglio di tipo 2)
        /// </summary>
        /// <param name="currentRowIndex"></param>
        /// <returns></returns>
        private static Excel_PropertiesContainer BuildPropertiesContainerForRecognition_Format2Sheet_Leghe(int currentRowIndex)
        {

            #region ADEMPIMENTO DELLE PROPRIETA DI LEGA 

            // contenitore per le proprieta correnti di lega 
            Excel_PropertiesContainer currentPropertiesAlloy = new Excel_PropertiesContainer();
            currentPropertiesAlloy.PropertiesDefinition = new List<Excel_PropertyWrapper>();

            // riempimento delle proprieta obbligatorie
            foreach (KeyValuePair<int, string> currentMandatoryProperty in _mandatoryTitlesColumnMapper)
            {
                // creazione del contenitore per la proprieta obbligatoria attuale
                Excel_PropertyWrapper currentProperty = new Excel_PropertyWrapper(currentRowIndex,
                    currentMandatoryProperty.Key,
                    currentMandatoryProperty.Value,
                    false);

                // aggiunta della proprieta obbligatoria al contenitore
                currentPropertiesAlloy.PropertiesDefinition.Add(currentProperty);
            }

            // riempimento delle proprieta opzionali
            foreach(KeyValuePair<int, string> currentOptionalAlloy in _mandatoryTitlesColumnMapper)
            {
                // creazione del contenitore per la proprieta opzionale attuale
                Excel_PropertyWrapper currentProperty = new Excel_PropertyWrapper(currentRowIndex,
                    currentOptionalAlloy.Key,
                    currentOptionalAlloy.Value,
                    true);

                // aggiunta della proprieta opzionale al contenitore
                currentPropertiesAlloy.PropertiesDefinition.Add(currentProperty);
            }

            #endregion
            
            return currentPropertiesAlloy;
        }


        /// <summary>
        /// Recupero di tutti i possibili valori per le proprieta degli elementi per la riga (e quindi la lega) che viene passata in input
        /// per l'iterazione corrente 
        /// </summary>
        /// <param name="currentRowIndex"></param>
        /// <returns></returns>
        private static List<Excel_PropertiesContainer> BuildPropertiesContainerForRecognition_Format2Sheet_ConcentrationsOnRow(int currentRowIndex)
        {
            // inizializzazione del contenitore per tutte le proprieta di concentrazioni correntemente in lettura 
            List<Excel_PropertiesContainer> currentPropertiesConcentrations = new List<Excel_PropertiesContainer>();

            // selezione di tutte le proprieta relative alle concentrazioni per la riga corrente 
            List<MapperElementFormat2> propertiesObjConcentrationsCurrentRow = _mapperElementsFormat2Sheet3.Where(x => x.CurrentRowIndex == currentRowIndex).ToList();

            // iterazione e valorizzazione per l'oggetto corrente
            Excel_PropertiesContainer propertiesForCurrentAlloy = new Excel_PropertiesContainer();
            propertiesForCurrentAlloy.PropertiesDefinition = new List<Excel_PropertyWrapper>();

            foreach(MapperElementFormat2 propertiesElements in propertiesObjConcentrationsCurrentRow)
            {
                // proprieta obbligatorie
                foreach(KeyValuePair<int, string> mandatoryProperyInstance in propertiesElements.MandatoryProperties)
                {
                    Excel_PropertyWrapper currentProperty = new Excel_PropertyWrapper(currentRowIndex,
                        mandatoryProperyInstance.Key,
                        mandatoryProperyInstance.Value,
                        false);

                    // aggiunta della proprieta obbligatoria al set di tutte le proprieta per l'elemento corrente 
                    propertiesForCurrentAlloy.PropertiesDefinition.Add(currentProperty);
                }


                // proprieta opzionali
                foreach(KeyValuePair<int, string> optionalPropertyInstance in propertiesElements.OptionalProperties)
                {
                    Excel_PropertyWrapper currentProperty = new Excel_PropertyWrapper(
                        currentRowIndex,
                        optionalPropertyInstance.Key,
                        optionalPropertyInstance.Value,
                        true
                        );

                    // aggiunta della proprieta opzionale al set di tutte le proprieta per l'elemento corrente 
                    propertiesForCurrentAlloy.PropertiesDefinition.Add(currentProperty);
                }

                // aggiunta del set di tutte le proprieta individuate per l'elemento corrente all'interno dell'insieme finale
                currentPropertiesConcentrations.Add(propertiesForCurrentAlloy);
            }

            // ritorno del set finale per le proprieta 
            return currentPropertiesConcentrations;
        }




        /// <summary>
        /// Permette la compilazione dell'eventuale messaggio di errore da inserire per l'elemento correntemente in analisi 
        /// viene ritornato false se per il dizionario individuato non sono state lette correttamente tutte le proprieta in analisi
        /// </summary>
        /// <param name="mandatoryTitles"></param>
        /// <param name="propertiesToVerify"></param>
        /// <returns></returns>
        private static bool CompileErrorMessage(List<string> mandatoryTitles, Dictionary<int, string> propertiesToVerify)
        {
            // TODO : implementazione partendo dai metodi gia definiti per questo caso 
            return false;
        }


        /// <summary>
        /// Permette la compilazione dell'eventuale messaggio di warning da inserire per l'elemeneto correntemente in analisi
        /// in questo caso non è necessario un valore di ritorno in quanto le proprieta opzionali non sono discriminanti per una
        /// eventuale corretta lettura dell'oggetto finale
        /// </summary>
        /// <param name="optionalTitles"></param>
        /// <param name="propertiesToVerify"></param>
        private static void CompileWarningMessage(List<string> optionalTitles, Dictionary<int, string> propertiesToVerify)
        {
            // TODO : implementazione partendo dai metodi definiti per questo caso
        }

        #endregion
    }
}
