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
        /// Mappatura di una proprieta nulla letta per la riga corrente 
        /// </summary>
        private static Dictionary<int, string> _mapperNullPropertiesOnRows;
        

        /// <summary>
        /// Lista dei warnings eventualmente restituiti nel caso in cui durante il recupero dei valori qualche validazione non passa
        /// </summary>
        private static string _listaWarnings_LetturaFoglio = String.Empty;


        /// <summary>
        /// Lista degli errori eventualmente restituiti nel caso in cui durante il recupero dei valori qualche validazione non passa provocando errori
        /// per i quali non è possibile continuare con l'analisi 
        /// </summary>
        private static string _listaErrori_LetturaFoglio = String.Empty;


        /// <summary>
        /// Dizionario relativo alla lettura delle proprieta per le concentrazioni correnti per l'instanza di un certo elemento e per il materiale 
        /// corrente 
        /// </summary>
        private static Dictionary<int, string> _instanceConcentrationColMapper;
        
        #endregion


        #region RECUPERO INFORMAZIONI PER IL FORMATO 1 EXCEL

        #region RECUPERO INFORMAZIONI DI LEGA

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


            // compilazione messaggi per la riga corrente per la lega
            CompileErrorMessageForLeghe_Format1(readProperties);
            CompileWarningMessageForLeghe_Format1(readProperties);



            // le proprieta obbligatorie non sono state lette correttamente per il foglio corrente 
            if (readProperties.CounterMandatoryProperties < Constants_Excel.PROPRIETAOBBLIGATORIE_FORMAT1_SHEET1.Count())
                return false;

            return true;
        }


        /// <summary>
        /// Permette di compilare un messaggio di errore riferibile alla entry corrente di lega
        /// </summary>
        /// <param name="readProperties"></param>
        private static void CompileErrorMessageForLeghe_Format1(Excel_PropertyWrapper readProperties)
        {
            foreach(string mandatoryProperty in Constants_Excel.PROPRIETAOBBLIGATORIE_FORMAT1_SHEET1.ToList())
            {
                if (readProperties.GetMandatoryProperty(mandatoryProperty) == String.Empty)
                    _listaErrori_LetturaFoglio += String.Format(Excel_ErrorMessages.Formato1_Foglio1_Leghe.ERRORE_MANDATORYPROPERTYMANCANTE_LEGA, _currentRowIndex, mandatoryProperty);
            }
        }


        /// <summary>
        /// Permette di compilare un messaggio di warning riferibile alla entry corrente di lega
        /// </summary>
        /// <param name="readProperties"></param>
        private static void CompileWarningMessageForLeghe_Format1(Excel_PropertyWrapper readProperties)
        {
            foreach(string optionalProperty in Constants_Excel.PROPRIETAOPZIONALI_FORMAT1_SHEET1.ToList())
            {
                if (readProperties.GetOptionalProperty(optionalProperty) == String.Empty)
                    _listaWarnings_LetturaFoglio += String.Format(Excel_WarningMessages.Formato1_Foglio1_Leghe.WARNING_MANCANZAVALOREPERPROPRIETAOPZIONALE_LEGA, _currentRowIndex, optionalProperty);
            }
        }

        #endregion


        #region RECUPERO INFORMAZIONI PER LE CONCENTRAZIONI

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
            // inizializzazione delle 2 liste di errori warnings per l'iterazione sul foglio corrente 
            _listaWarnings_LetturaFoglio = String.Empty;
            _listaErrori_LetturaFoglio = String.Empty;

            // validazione e inserimento del foglio in lettura corrente 
            if (currentFoglioExcel == null)
                throw new Exception(ExceptionMessages.EXCEL_READERINFO_FOGLIONULLPERLETTURA);

            // validazione relativa alla lista dei quadranti dai quali andare a leggere i valori attuali delle concentrazioni
            if (emptyConcentrationsInfo.GetConcQuadrants_Type2 == null)
                throw new Exception(ExceptionMessages.EXCEL_READERINFO_QUADRANTICONCENTRAZIONINULLIPERLETTURA);
            if(emptyConcentrationsInfo.GetConcQuadrants_Type2.Count() == 0)
                throw new Exception(ExceptionMessages.EXCEL_READERINFO_QUADRANTICONCENTRAZIONINULLIPERLETTURA);

            _currentFoglioExcel = currentFoglioExcel;
            
            // inizializzazione istanza di primo foglio
            Excel_Format1_Sheet2_ConcQuadrant currentRowPropertiesFoglio = new Excel_Format1_Sheet2_ConcQuadrant();


            // dizionario per la mappatura di proprieta in lettura nulle da foglio excel corrente 
            _mapperNullPropertiesOnRows = new Dictionary<int, string>();


            // inzio iterazione su quadranti concentrazioni per lettura corrente
            foreach (Excel_Format1_Sheet2_ConcQuadrant currentQuadrant in emptyConcentrationsInfo.GetConcQuadrants_Type2)
            {
                // indica la validazione relativa alle informazioni rispetto a titolo e concentrazioni
                bool titleValidation = false;
                bool concValidation = false;

                currentQuadrant.ValidatedOnExcel = false;


                // tentativo di valorizzazione del nome nel caso in cui sia null o empty
                if ((currentQuadrant.NomeMateriale == null || currentQuadrant.NomeMateriale == String.Empty) &&
                    _currentFoglioExcel.Cells[currentQuadrant.StartingRow_Title, currentQuadrant.StartigCol] != null)
                {
                    currentQuadrant.NomeMateriale = _currentFoglioExcel.Cells[currentQuadrant.StartingRow_Title, currentQuadrant.StartigCol].Value.ToString();
                    titleValidation = true;
                }
                    
                else
                {
                    // segnalo nei messaggi di errore che non potro proseguire per la prossima analisi in quanto il titolo materiale è vuoto
                    _listaErrori_LetturaFoglio += String.Format(Excel_ErrorMessages.Formato1_Foglio2_Concentrazioni.ERRORE_NOMEMATERIALELETTURAQUADRANTEVUOTO, currentQuadrant.StartingRow_Title, currentQuadrant.StartigCol);

                    // invalidazione automatica per il quadrante di concentrazioni corrente
                    currentQuadrant.ValidatedOnExcel = false;
                }
                

                // posizionamento da lettura headers
                _currentRowIndex = currentQuadrant.StartingRow_Headers;
                
                // riempo per le proprieta in lettura corrente per il quadrante, con ttute le proprieta valorizzate per l'header e in base al valore inserito per indice di riga
                FillPropertyMapper(currentQuadrant.StartigCol, currentQuadrant.EndingCol);

                // posizionamento da lettura concentrazioni
                _currentRowIndex = currentQuadrant.StartingRow_Concentrations;

                // imposto = true e verifico che durante l'iterazione non succeda che anche solo una riga venga poi invalidata 
                concValidation = true;

                while (_currentRowIndex <= currentQuadrant.EndingRow_Concentrations)
                {


                    // istanza che verrà riempita con la lista delle concentrazioni qualora venga passato il controllo
                    Excel_PropertyWrapper currentReadConcentrations;
                    if (ReadConcentrationsForCurrentQuadrant(currentQuadrant, out currentReadConcentrations))
                    {
                        currentQuadrant.Concentrations.Add(currentReadConcentrations);
                        
                    }
                    // invalidazione automatica per la prima validazione e il quadrante di concentrazione corrente     
                    else
                        concValidation = false;

                    _currentRowIndex++;
                }


                // il quadrante passa la validazione se informazioni su title e concentrazioni sono correttamente attribuite
                if (concValidation && titleValidation)
                    currentQuadrant.ValidatedOnExcel = true;

                
            }

            filledConcentrationsInfo = emptyConcentrationsInfo;
            return EsitoRecuperoInformazioniFoglio.RecuperoCorretto;
        }

        
        /// <summary>
        /// Lettura delle concentrazioni dal quadrante delle concentrazioni corrente, 
        /// se non si riesce a valorizzare tutte le proprieta di quadrante viene restituito false e le readConcentrations rimangono 0
        /// altrimenti si aggiungera la definizione data a tutte le concentrazioni lette al quadrante corrente
        /// </summary>
        /// <param name="currentQuadrant"></param>
        /// <param name="readConcentrations"></param>
        /// <returns></returns>
        private static bool ReadConcentrationsForCurrentQuadrant(Excel_Format1_Sheet2_ConcQuadrant currentQuadrant, out Excel_PropertyWrapper readConcentrations)
        {
            // istanza in restituzione per tutti i valori letti per le concentrazioni
            readConcentrations = new Excel_PropertyWrapper(Constants_Excel.PROPRIETAOBBLIGATORIE_FORMAT1_SHEET2, Constants_Excel.PROPRIETAOPZIONALI_FORMAT1_SHEET2, TipologiaPropertiesFoglio.Format1_Foglio2_Concentrazioni);
            

            // inizio iteraizone su proprieta 
            foreach (KeyValuePair<int, string> currentPropertyHeader in PropertiesColMapper)
            {
                // ho trovato un valore per la proprieta corrente
                if (_currentFoglioExcel.Cells[_currentRowIndex, currentPropertyHeader.Key].Value != null)
                {
                    // inserisco unicamente se ritrovo la proprieta nelle definizioni date per proprieta opzionali / obbligatorie
                    if (Constants_Excel.PROPRIETAOBBLIGATORIE_FORMAT1_SHEET2.Contains(currentPropertyHeader.Value))
                        readConcentrations.InsertMandatoryValue(currentPropertyHeader.Value, _currentFoglioExcel.Cells[_currentRowIndex, currentPropertyHeader.Key].Value.ToString());
                    else if (Constants_Excel.PROPRIETAOPZIONALI_FORMAT1_SHEET2.Contains(currentPropertyHeader.Value))
                        readConcentrations.InsertOptionalValue(currentPropertyHeader.Value, _currentFoglioExcel.Cells[_currentRowIndex, currentPropertyHeader.Key].Value.ToString());
                }
                else
                {
                    // mappatura della proprieta nulla in lettura corrente 
                    _mapperNullPropertiesOnRows.Add(_currentRowIndex, currentPropertyHeader.Value);
                } 

            }

            // valorizzazione eventuali messaggi warnings errori per l'importazione delle informazioni di concentrazioni correnti
            CompileErrorMessages_ConcentrazioniFoglio1(readConcentrations);
            CompileWarningMessageForLeghe_Format1(readConcentrations);


            // le proprieta obbligatorie non sono state lette correttamente per il foglio corrente 
            if (readConcentrations.CounterMandatoryProperties < Constants_Excel.PROPRIETAOBBLIGATORIE_FORMAT1_SHEET2.Count())
                return false;
            

            return true;
        }


        /// <summary>
        /// Compilazione dei messaggi di errore individuati per la valorizzazione delle informazioni correnti di concentrazioni
        /// </summary>
        /// <param name="currentConcentrations"></param>
        private static void CompileErrorMessages_ConcentrazioniFoglio1(Excel_PropertyWrapper currentConcentrations)
        {
            // valorizzazione messaggi di errore per le proprieta excel
            foreach(string currentaMandatoryProperty in Constants_Excel.PROPRIETAOBBLIGATORIE_FORMAT1_SHEET2)
            {
                // se nel dizionario è coontenuta una proprieta che riporta lo stesso titolo di quella in analisi, significa che ho trovato una proprieta nulla obbigatoria
                Dictionary<int, string> _nullPropertyDefinitions = _mapperNullPropertiesOnRows.Where(x => x.Value == currentaMandatoryProperty).ToDictionary(x => x.Key, x => x.Value);

                foreach (KeyValuePair<int, string> currentNullValue in _nullPropertyDefinitions)
                    _listaErrori_LetturaFoglio += String.Format(Excel_ErrorMessages.Formato1_Foglio2_Concentrazioni.ERRORE_MANDATORYPROPERTYMANCANTE_CONCENTRAZIONI, currentNullValue.Key, currentNullValue.Value);

            }
        }


        


        /// <summary>
        /// Compilazione dei messaggi di warnings individuati per la valorizzazione delle informazioni correnti di concentrazioni
        /// </summary>
        /// <param name="currentConcentrations"></param>
        private static void CompilaWarningMessages_ConcentrazioniFoglio1(Excel_PropertyWrapper currentConcentrations)
        {
            // valorizzazione messaggi di warnings per le proprieta excel 
            foreach(string currentOptionalProperty in Constants_Excel.PROPRIETAOPZIONALI_FORMAT1_SHEET2)
            {
                // valorizzazione dell'eventuale dizionario contenente la definizione per i messaggi di warnings sulle proprieta opzionali
                Dictionary<int, string> _nullPropertyDefinitions = _mapperNullPropertiesOnRows.Where(x => x.Value == currentOptionalProperty).ToDictionary(x => x.Key, x => x.Value);

                foreach (KeyValuePair<int, string> currentNullValue in _nullPropertyDefinitions)
                    _listaWarnings_LetturaFoglio += String.Format(Excel_WarningMessages.Formato1_Foglio2_Concentrazioni.WARNING_MANCANZAVALOREPERPROPRIETAOPZIONALE_CONCENTRAZIONI, currentNullValue.Key, currentNullValue.Value);
            }
        }

        #endregion

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

            // inizializzazione delle 2 liste di errori warnings per l'iterazione sul foglio corrente 
            _listaWarnings_LetturaFoglio = String.Empty;
            _listaErrori_LetturaFoglio = String.Empty;


            // inizializzazione degli indici di riga per le proprieta eventualmente nulle riscontrate nella lettura
            _mapperNullPropertiesOnRows = new Dictionary<int, string>();


            // validazione e inserimento del foglio in lettura corrente 
            if (currentFoglioExcel == null)
                throw new Exception(ExceptionMessages.EXCEL_READERINFO_FOGLIONULLPERLETTURA);

            // foglio excel corrente 
            _currentFoglioExcel = currentFoglioExcel;



            // istanza del mappatore per le colonne di header su riga per questo e relativamente alle informazioni di lega 
            PropertiesColMapper = new Dictionary<int, string>();

            // istanza del mappatore per le colonne di header relative alla lettura per un determinato elemento
            _instanceConcentrationColMapper = new Dictionary<int, string>();

            // indice di riga per l'header principale (PRIMA RIGA)
            int principalRowHeader = emptyInfo.StartingRow_Leghe;

            // indice di riga per l'header secodnario (LETTURA HEADERS PER LE DIVERSE CONCENTRAZIONI)
            int secondaryRowHeader_Concentrations = principalRowHeader + 1;

            // posizionamento dell'indice riga globale alla lettura del primo header
            _currentRowIndex = principalRowHeader;

            // calcolo gli indici di inizio e fine lettura informazioni lega per il foglio corrente e riguardanti le colonne
            int startingColIndexLeghe = emptyInfo.StartingCol_Leghe;
            int endingColIndexLeghe = emptyInfo.ColonneConcentrazioni.Select(x => x.startingCol_Header).Min() - 1;

            // inizializzazione dei valori header per le proprieta generali di lega (viene inizializzato in questo punto una sola volta)
            InizializeMapperHeaderLeghe(startingColIndexLeghe, endingColIndexLeghe);

            
            // posizione di inizio lettura informazioni per il file excel corrente 
            // prendo in considerazione il salto della prima riga relativa all'indice 1 di header
            // prendo in considerazione il salto della seconda riga relativa all'indice 2 di header
            _currentRowIndex = emptyInfo.StartingRow_Leghe + 2;


            // inizio iterazione delle righe su cui leggere tutti i valori per i casi di lega in lettura 
            while(_currentRowIndex <= _currentFoglioExcel.Dimension.End.Row)
            {
                // istanza lettura proprieta di lega 
                Excel_PropertyWrapper readLegaProperties;
                bool hoLettoProprietaLegaCorrenti = ReadGeneralInfoLega(out readLegaProperties);

                // istanza lettura proprieta di concentrazioni
                Excel_Format2_ConcColumns readColumnsConcentrations;
                bool hoLet;
                

            }
            

            filledInfo = emptyInfo;
            return EsitoRecuperoInformazioniFoglio.RecuperoCorretto;



        }


        /// <summary>
        /// Lettura delle informazioni di lega generali per il file excel corrente per il secondo formato
        /// in input sono inseriti gli indici di inizio / fine lettura per le colonne sulle quali leggere queste proprieta
        /// l'indice di fine lettura per le proprieta coincide con l'indice di inizio lettura per le informazioni di concentrazioni -1
        /// </summary>
        /// <param name="currentInfoLega"></param>
        private static bool ReadGeneralInfoLega(out Excel_PropertyWrapper currentInfoLega)
        {
            // istanza proprieta lette per la lega correntemente in analisi
            currentInfoLega = new Excel_PropertyWrapper(Constants_Excel.PROPRIETAOBBLIGATORIE_FORMAT2_LEGHE, Constants_Excel.PROPRIETAOPZIONALI_FORMAT2_leghe, TipologiaPropertiesFoglio.Format2_Leghe);
            

            foreach(KeyValuePair<int, string> currentProperty in PropertiesColMapper)
            {
                // check di non null per il valore corrispondente alla proprieta in lettura (proprieta obbligatoria)
                if(Constants_Excel.PROPRIETAOBBLIGATORIE_FORMAT2_LEGHE.Contains(currentProperty.Value))
                {
                    if (_currentFoglioExcel.Cells[_currentRowIndex, currentProperty.Key].Value != null)
                        currentInfoLega.InsertMandatoryValue(currentProperty.Value, _currentFoglioExcel.Cells[_currentRowIndex, currentProperty.Key].Value.ToString());
                    // inserimento del caso corrente nelle proprieta non valorizzate per la lega corrente 
                    else
                        _mapperNullPropertiesOnRows.Add(_currentRowIndex, currentProperty.Value);
                }


                // check di non null per il valore corrispondente alla proprieta in lettura (proprieta opzionale)
                if (Constants_Excel.PROPRIETAOPZIONALI_FORMAT2_leghe.Contains(currentProperty.Value))
                {
                    if (_currentFoglioExcel.Cells[_currentRowIndex, currentProperty.Key].Value != null)
                        currentInfoLega.InsertOptionalValue(currentProperty.Value, _currentFoglioExcel.Cells[_currentRowIndex, currentProperty.Key].Value.ToString());
                    // inserimento del caso corrente nelle proprieta non valorizzate per la lega corrente (proprieta opzionale)
                    else
                        _mapperNullPropertiesOnRows.Add(_currentRowIndex, currentProperty.Value);
                }
                    
                    
            }

            // calcolo eventuali messaggi di warnings / errore per la lettura delle leghe
            CompileErrorMessages_ReadInfoLeghe_Format2(currentInfoLega);
            CompileWarningMessages_ReadingInfoLeghe_Foramt2(currentInfoLega);

            // restituisco false se non leggo tutte le proprieta obbligatorie
            if (currentInfoLega.CounterMandatoryProperties != Constants_Excel.PROPRIETAOBBLIGATORIE_FORMAT2_LEGHE.Count())
                return false;

            return true;
                
        }


        /// <summary>
        /// Permette di riempire oggetto proprieta lettura per un determinato elemento di concentrazioni, in input è passato il quadrante excel 
        /// riconosciuto in fase di validazione, in out put viene restituito questo quadrante riempito con le definizioni per le diverse proprieta di elemento in lettura 
        /// per il secondo formato e la lega corrente 
        /// La proprieta relativa al main header mi serve per la mappatura del nome da attribuire all'elemento corrente 
        /// La riga di header secondario mi serve per calolare gli indici di colonna per le diverse proprieta di header in lettura per l'elemento corrente 
        /// </summary>
        /// <param name="mainHeaderRow"></param>
        /// <param name="secondaryHeaderRow"></param>
        /// <param name="startingConcElem"></param>
        /// <param name="filledConcElem"></param>
        /// <returns></returns>
        private static bool readCurrentConcentrationsInstance(int mainHeaderRow, int secondaryHeaderRow, Excel_Format2_ConcColumns startingConcElem, out Excel_Format2_ConcColumns filledConcElem)
        {
            bool hoLettoNomeElemento = false;
            bool hoLettoConcentrationi = false;

            // lettura del nome per l'elemento corrente 
            if (_currentFoglioExcel.Cells[mainHeaderRow, startingConcElem.startingCol_Header].Value != null)
            {
                startingConcElem.NomeElemento = _currentFoglioExcel.Cells[mainHeaderRow, startingConcElem.startingCol_Header].Value.ToString();
                hoLettoNomeElemento = true;
            }
            // segnalazione del nome per l'elemento corrente e la lega in analisi corrente 
            else
                _listaErrori_LetturaFoglio += String.Format(Excel_ErrorMessages.Formato2_Foglio1_LegheConcentrazioni.ERRORE_MANCATALETTURANOMEELEMENTO, mainHeaderRow, startingConcElem.startingCol_Header);

            // lettura indici di colonne per le diverse proprieta sulle quali leggo per gli headers
            InizializeMapperHeaderConcentrations(secondaryHeaderRow, startingConcElem.startingCol_Header, startingConcElem.endingCol_Header);


            // inizio recupero proprieta correnti per elemento
            foreach(KeyValuePair<int, string> currentProperty in _instanceConcentrationColMapper)
            {
                if (_currentFoglioExcel.Cells[_currentRowIndex, currentProperty.Key].Value != null)
                {
                    // aggiunta della proprieta obbligatoria per l'elemento
                    if (Constants_Excel.PROPRIETAOBBLIGATORIE_ELEM_FORMAT2.Contains(currentProperty.Value))
                        startingConcElem.ReadProperties.InsertMandatoryValue(currentProperty.Value, _currentFoglioExcel.Cells[_currentRowIndex, currentProperty.Key].Value.ToString());
                    else if (Constants_Excel.PROPRIETAOPZIONALI_ELEM_FORMAT2.Contains(currentProperty.Value))
                        startingConcElem.ReadProperties.InsertOptionalValue(currentProperty.Value, _currentFoglioExcel.Cells[_currentRowIndex, currentProperty.Key].Value.ToString());
                }
                else
                {
                    // aggiunta nel dizionario delle proprieta lette come null per l'elemento corrente
                    _mapperNullPropertiesOnRows.Add(_currentRowIndex, currentProperty.Value);
                }
            }

            // indico di aver letto correttamente la concentrazione se ho letto tutte le proprieta obbligatorie per questa
            if (startingConcElem.ReadProperties.CounterMandatoryProperties == Constants_Excel.PROPRIETAOBBLIGATORIE_ELEM_FORMAT2.Count())
                hoLettoConcentrationi = true;


            // inserimento della messaggistica relativamente a warnings e errori trovati sul foglio in analisi corrente per le concentrazioni e la seconda tipologia di formato
            CompileErrorMessages_ReadInfoConcentrations_Format2(startingConcElem.ReadProperties);
            CompileWarningMessages_ReadInfoConcentrations_Foramt2(startingConcElem.ReadProperties);


            // istanza per la compilazione delle proprieta per la concentrazione corrente 
            filledConcElem = startingConcElem;

            // ritorno true solamente se ho letto sia il nome che le concentrazioni per l'elemento corrente 
            if (hoLettoConcentrationi && hoLettoNomeElemento)
                return true;

            return false;
        }


        /// <summary>
        /// Inizializzazione per il mapper delle proprieta in lettura corrente per gli headers relativi alle informazioni 
        /// di lega in lettura per la riga corrente 
        /// </summary>
        /// <param name="startingColIndex"></param>
        /// <param name="endingColIndex"></param>
        private static void InizializeMapperHeaderLeghe(int startingColIndex, int endingColIndex)
        {
            _currentColIndex = startingColIndex;

            while(_currentColIndex <= endingColIndex)
            {
                if(_currentFoglioExcel.Cells[_currentRowIndex, _currentColIndex].Value != null)
                {
                    // match definizione proprieta obbligatoria / opzionale
                    if (Constants_Excel.PROPRIETAOBBLIGATORIE_FORMAT2_LEGHE.Contains(_currentFoglioExcel.Cells[_currentRowIndex, _currentColIndex].Value.ToString().ToUpper())
                        || Constants_Excel.PROPRIETAOPZIONALI_FORMAT2_leghe.Contains(_currentFoglioExcel.Cells[_currentRowIndex, _currentColIndex].Value.ToString().ToUpper()))
                        PropertiesColMapper.Add(_currentColIndex, _currentFoglioExcel.Cells[_currentRowIndex, _currentColIndex].Value.ToString().ToUpper());
                }
            }

            _currentColIndex++;
        }


        /// <summary>
        /// Inizializzazione per il mapper delle proprieta in lettura corrente per gli headers relativi alle informazioni 
        /// di concentrazione in lettura per la riga corrente 
        /// </summary>
        /// <param name="startingRowHeader"></param>
        /// <param name="startingHeadConc"></param>
        /// <param name="endingHeadConc"></param>
        private static void InizializeMapperHeaderConcentrations(int startingRowSecondaryHeader, int startingHeadConc, int endingHeadConc)
        {
            // posizione primo indice di colonna 
            _currentColIndex = startingHeadConc;

            // inizializzazione wrapper headers per proprieta correnti per l'elemento
            _instanceConcentrationColMapper = new Dictionary<int, string>();

            while(_currentColIndex <= endingHeadConc)
            {
                // match definizioni proprieta obbligatorie opzionali per header di colonna concentrazioni
                if (Constants_Excel.PROPRIETAOBBLIGATORIE_ELEM_FORMAT2.Contains(_currentFoglioExcel.Cells[startingRowSecondaryHeader, _currentColIndex].Value.ToString().ToUpper())
                    || Constants_Excel.PROPRIETAOPZIONALI_ELEM_FORMAT2.Contains(_currentFoglioExcel.Cells[startingRowSecondaryHeader, _currentColIndex].Value.ToString().ToUpper()))
                    _instanceConcentrationColMapper.Add(_currentColIndex, _currentFoglioExcel.Cells[startingRowSecondaryHeader, _currentColIndex].Value.ToString().ToUpper());

                _currentColIndex++;
            }
        }


        /// <summary>
        /// Permette per la compilazione del messaggio di errore nella lettura delle informazioni per la lega corrente, questo nel caso in cui siano trovate 
        /// delle celle vuote in corrispondenza di tali proprieta (proprieta obbligatorie)
        /// </summary>
        /// <param name="readLegaProperties"></param>
        private static void CompileErrorMessages_ReadInfoLeghe_Format2(Excel_PropertyWrapper readLegaProperties)
        {
            // iterazione per le proprieta obbligatorie relative alle leghe per il secondo formato disponibile excel
            foreach(string currentMandatoryPropertyLega in Constants_Excel.PROPRIETAOBBLIGATORIE_FORMAT2_LEGHE)
            {
                // iterazione per valori nulli e proprieta obbligatoria corrente 
                foreach(KeyValuePair<int, string> valoreRigaNullo in _mapperNullPropertiesOnRows.Where(x => x.Value == currentMandatoryPropertyLega).ToDictionary(x => x.Key, x => x.Value))
                {
                    // recupero indice di colonna 
                    int colProprietaNulla = PropertiesColMapper.Where(x => x.Value == currentMandatoryPropertyLega).Select(x => x.Key).FirstOrDefault();

                    // inserimento del messaggio di errore per la proprieta corrente obbligatoria di cui è stata mancata la lettura
                    _listaErrori_LetturaFoglio += String.Format(Excel_ErrorMessages.Formato2_Foglio1_LegheConcentrazioni.ERRORE_MANCATALETTURAPROPRIETAOBBLIGATORIA_LEGA, valoreRigaNullo.Key, colProprietaNulla, currentMandatoryPropertyLega);
                }
            }
        }


        /// <summary>
        /// Permette per la compilazione del messaggio di warning per la lettuira delle informazioni per la lega corrente, questo nel caso in cui siano trovate 
        /// delle celle vuote in corrispondenza delle proprieta opzionali
        /// </summary>
        /// <param name="readLegaProperties"></param>
        private static void CompileWarningMessages_ReadingInfoLeghe_Foramt2(Excel_PropertyWrapper readLegaProperties)
        {
            // iterazione per la proprieta opzionale relativa alla lega per il secondo formato disponibile excel
            foreach(string currentOptionalPropertyLega in Constants_Excel.PROPRIETAOPZIONALI_FORMAT2_leghe)
            {
                foreach(KeyValuePair<int, string> valoreRigaNullo in _mapperNullPropertiesOnRows.Where(x => x.Value == currentOptionalPropertyLega).ToDictionary(x => x.Key, x => x.Value))
                {
                    // recupero indice di colonna
                    int colProprietaNulla = PropertiesColMapper.Where(x => x.Value == currentOptionalPropertyLega).Select(x => x.Key).FirstOrDefault();

                    // inserimento del messaggio di warning per la proprieta corrente opzionale di cui è stata mancata la lettura
                    _listaWarnings_LetturaFoglio += String.Format(Excel_WarningMessages.Formato2_Foglio1_LegheConcentrazioni.WARNING_MANCATALETTURAPROPRIETAOPZIONALE_LEGA, valoreRigaNullo.Key, colProprietaNulla, currentOptionalPropertyLega);
                }
            }
        }


        /// <summary>
        /// Permette la compilazione per i messaggi di errore relativamente alla mancata lettura delle informazioni obbligatorie per ogni elemento 
        /// e per le concentrazioni correnti
        /// </summary>
        /// <param name="readConcentrationsProperties"></param>
        private static void CompileErrorMessages_ReadInfoConcentrations_Format2(Excel_PropertyWrapper readConcentrationsProperties)
        {
            // iterazione per le proprieta obbligatorie relative alle concentrazioni per il secondo formato disponibile excel
            foreach (string currentMandatoryPropertyConcentrations in Constants_Excel.PROPRIETAOBBLIGATORIE_ELEM_FORMAT2)
            {
                // iterazione per valori nulli e proprieta obbligatoria corrente 
                foreach (KeyValuePair<int, string> valoreRigaNullo in _mapperNullPropertiesOnRows.Where(x => x.Value == currentMandatoryPropertyConcentrations).ToDictionary(x => x.Key, x => x.Value))
                {
                    // recupero indice di colonna 
                    int colProprietaNulla = _instanceConcentrationColMapper.Where(x => x.Value == currentMandatoryPropertyConcentrations).Select(x => x.Key).FirstOrDefault();

                    // inserimento del messaggio di errore per la proprieta corrente obbligatoria di cui è stata mancata la lettura
                    _listaErrori_LetturaFoglio += String.Format(Excel_ErrorMessages.Formato2_Foglio1_LegheConcentrazioni.ERRORE_MANCATALETTURAPROPRIETAOBBLIGATORIA_CONCENTRAZIONI, valoreRigaNullo.Key, colProprietaNulla, currentMandatoryPropertyConcentrations);
                }
            }
        }

        
        /// <summary>
        /// Permette la compilazione per i messaggi di warning relativamente alla mancata lettura delle informazioni opzionali per ogni elemento 
        /// e per le concentrazioni correnti
        /// </summary>
        /// <param name="readConcentrationsProperties"></param>
        private static void CompileWarningMessages_ReadInfoConcentrations_Foramt2(Excel_PropertyWrapper readConcentrationsProperties)
        {
            // iterazione per la proprieta opzionale relativa alla lega per il secondo formato disponibile excel
            foreach (string currentOptionalPropertyConcentrations in Constants_Excel.PROPRIETAOPZIONALI_ELEM_FORMAT2)
            {
                foreach (KeyValuePair<int, string> valoreRigaNullo in _mapperNullPropertiesOnRows.Where(x => x.Value == currentOptionalPropertyConcentrations).ToDictionary(x => x.Key, x => x.Value))
                {
                    // recupero indice di colonna
                    int colProprietaNulla = _instanceConcentrationColMapper.Where(x => x.Value == currentOptionalPropertyConcentrations).Select(x => x.Key).FirstOrDefault();

                    // inserimento del messaggio di warning per la proprieta corrente opzionale di cui è stata mancata la lettura
                    _listaWarnings_LetturaFoglio += String.Format(Excel_WarningMessages.Formato2_Foglio1_LegheConcentrazioni.WARNING_MANCATALETTURAPROPRIETAOPZIONALE_CONCENTRAZIONI, valoreRigaNullo.Key, colProprietaNulla, currentOptionalPropertyConcentrations);
                }
            }
        }

        #endregion

    }
}
