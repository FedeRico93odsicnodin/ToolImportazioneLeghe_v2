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
        /// Riferimento alle proprieta del foglio custom provenienti dalla validazione precedente e che sono ancora a empty
        /// rispetto al tentaitvo di fill fatto in questa fase
        /// </summary>
        private static Excel_AlloyInfo_Sheet _currentEmptyPropertiesSheetInstance;


        /// <summary>
        /// Istanza errori trovati su iterazione corrente 
        /// </summary>
        private static string _errorMessages_CurrentInstance = String.Empty;


        /// <summary>
        /// Istanza warnings trovati su iterazione corrente
        /// </summary>
        private static string _warningMessages_CurrentInstance = String.Empty;

        #endregion

        
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
        public static EsitoRecuperoInformazioniFoglio ReadLegheInfo(ExcelWorksheet currentFoglioExcel, Excel_AlloyInfo_Sheet emptyLegheInfo, out Excel_AlloyInfo_Sheet filledLegheInfo)
        {
            // inizializzazione foglio excel corrente 
            _currentFoglioExcel = currentFoglioExcel;

            // inizializzazione istanza foglio corrente 
            _currentEmptyPropertiesSheetInstance = emptyLegheInfo;

            // inizializzazione dei messaggi di errori warnings che si trovaranno con iterazione corrente 
            _errorMessages_CurrentInstance = String.Empty;
            _warningMessages_CurrentInstance = String.Empty;

            // indicazione di passaggio di almeno una validazione 
            bool atLeastAValidation = false;

            foreach(Excel_PropertiesContainer currentPropertiesLeghe in _currentEmptyPropertiesSheetInstance.AlloyPropertiesInstances)
            {
                // passaggio di almeno una validazione su foglio corrente 
                if (ReadPropertiesInfo(currentPropertiesLeghe))
                    atLeastAValidation = true;
            }


            // messaggi di errore / warnings per istanza corrente 
            _currentEmptyPropertiesSheetInstance.ErrorMessageSheet += _errorMessages_CurrentInstance;
            _currentEmptyPropertiesSheetInstance.WarningMessageSheet += _warningMessages_CurrentInstance;

            filledLegheInfo = _currentEmptyPropertiesSheetInstance;

            if (atLeastAValidation)
            {
                if (_warningMessages_CurrentInstance != String.Empty)
                    return EsitoRecuperoInformazioniFoglio.RecuperoConWarnings;

                return EsitoRecuperoInformazioniFoglio.RecuperoCorretto;
            }

            return EsitoRecuperoInformazioniFoglio.RecuperoConErrori;
        }

        
        /// <summary>
        /// Permette di recuperare tutte le informazioni di concentrazione all'interno dell'oggetto predisposto per contenere tutti i valori per le informazioni di concentrazioni
        /// in lettura corrente e per il foglio in analisi
        /// </summary>
        /// <param name="currentFoglioExcel"></param>
        /// <param name="emptyConcentrationsInfo"></param>
        /// <param name="filledConcentrationsInfo"></param>
        /// <param name="warnings_read_concentrations_list"></param>
        /// <param name="errors_read_concentrations_list"></param>
        /// <returns></returns>
        public static EsitoRecuperoInformazioniFoglio ReadConcentrationsInfo(ExcelWorksheet currentFoglioExcel, Excel_AlloyInfo_Sheet emptyConcentrationsInfo, out Excel_AlloyInfo_Sheet filledConcentrationsInfo)
        {
            // inizializzazione foglio excel corrente 
            _currentFoglioExcel = currentFoglioExcel;

            // inizializzazione istanza foglio corrente 
            _currentEmptyPropertiesSheetInstance = emptyConcentrationsInfo;

            // inizializzazione dei messaggi di errori warnings che si trovaranno con iterazione corrente 
            _errorMessages_CurrentInstance = String.Empty;
            _warningMessages_CurrentInstance = String.Empty;

            // indicazione di passaggio di almeno una validazione 
            bool atLeastAValidation = false;


            foreach (Excel_PropertiesContainer currentPropertiesConcentrations in _currentEmptyPropertiesSheetInstance.ConcentrationsPropertiesInstances)
            {
                // passaggio di almeno una validazione su foglio corrente 
                if (ReadPropertiesInfo(currentPropertiesConcentrations, true))
                    atLeastAValidation = true;
            }

            // messaggi di errore / warnings per istanza corrente 
            _currentEmptyPropertiesSheetInstance.ErrorMessageSheet += _errorMessages_CurrentInstance;
            _currentEmptyPropertiesSheetInstance.WarningMessageSheet += _warningMessages_CurrentInstance;

            filledConcentrationsInfo = _currentEmptyPropertiesSheetInstance;

            if (atLeastAValidation)
            {
                if (_warningMessages_CurrentInstance != String.Empty)
                    return EsitoRecuperoInformazioniFoglio.RecuperoConWarnings;

                return EsitoRecuperoInformazioniFoglio.RecuperoCorretto;
            }

            return EsitoRecuperoInformazioniFoglio.RecuperoConErrori;
        }



        /// <summary>
        /// Permette il recupero di tutte le informazioni per leghe e concentrazioni all'interno dell'oggetto predisposto per contenenere tutti i vlaori per le informazioni di
        /// leghe e concentrazioni in lettura corrente e per il foglio in analisi
        /// </summary>
        /// <param name="currentFoglioExcel"></param>
        /// <param name="emptyInfo"></param>
        /// <param name="filledInfo"></param>
        /// <param name="possibleReadErrors"></param>
        /// <param name="possibleReadWarnings"></param>
        /// <returns></returns>
        public static EsitoRecuperoInformazioniFoglio ReadInfoFormat2(ExcelWorksheet currentFoglioExcel, Excel_AlloyInfo_Sheet emptyInfo, out Excel_AlloyInfo_Sheet filledInfo)
        {
            // inizializzazione foglio excel corrente 
            _currentFoglioExcel = currentFoglioExcel;

            // inizializzazione istanza foglio corrente 
            _currentEmptyPropertiesSheetInstance = emptyInfo;

            // inizializzazione dei messaggi di errori warnings che si trovaranno con iterazione corrente 
            _errorMessages_CurrentInstance = String.Empty;
            _warningMessages_CurrentInstance = String.Empty;

            // indicazione di passaggio di almeno una validazione 
            bool atLeastAValidationOnLeghe = false;

            foreach (Excel_PropertiesContainer currentPropertiesLeghe in _currentEmptyPropertiesSheetInstance.AlloyPropertiesInstances)
            {
                // passaggio di almeno una validazione su foglio corrente 
                if (ReadPropertiesInfo(currentPropertiesLeghe))
                    atLeastAValidationOnLeghe = true;
            }

            // indicazione di passaggio di almeno una validazione 
            bool atLeastAValidationOnConcentrations = false;

            foreach (Excel_PropertiesContainer currentPropertiesConcentrations in _currentEmptyPropertiesSheetInstance.ConcentrationsPropertiesInstances)
            {
                // passaggio di almeno una validazione su foglio corrente 
                if (ReadPropertiesInfo(currentPropertiesConcentrations))
                    atLeastAValidationOnConcentrations = true;
            }


            // messaggi di errore / warnings per istanza corrente 
            _currentEmptyPropertiesSheetInstance.ErrorMessageSheet += _errorMessages_CurrentInstance;
            _currentEmptyPropertiesSheetInstance.WarningMessageSheet += _warningMessages_CurrentInstance;

            //
            bool validationOnLegheConcentrations = false;


            // ulteriore check fatto su elementi di lega e concentrazioni
            if (atLeastAValidationOnLeghe && atLeastAValidationOnConcentrations)
            {
                // se per ogni proprieta di lega non c'è validazione, allora non posso continuare 
                if(_currentEmptyPropertiesSheetInstance.AlloyPropertiesInstances.Where(x => x.ValidatedElem == true).Count() > 0)
                {
                    foreach (Excel_PropertiesContainer currentPropertiesLeghe in _currentEmptyPropertiesSheetInstance.AlloyPropertiesInstances)
                    {

                        if (_currentEmptyPropertiesSheetInstance.ConcentrationsPropertiesInstances.Where(x => x.StartingRowIndex == currentPropertiesLeghe.StartingRowIndex
                        ).Count() > 0 && currentPropertiesLeghe.ValidatedElem 
                            )
                        {
                            currentPropertiesLeghe.ValidatedAssociation = true;
                        }
                    }

                    // controllo se ho avuto una associazione almeno per un elemento di lega 
                    if (_currentEmptyPropertiesSheetInstance.AlloyPropertiesInstances.Where(x => x.ValidatedAssociation == true).Count() > 0)
                        validationOnLegheConcentrations = true;
                }
            }

            filledInfo = _currentEmptyPropertiesSheetInstance;

            // se ho passato anche l'ultima associazione allora posso ritornare senza errori
            if (validationOnLegheConcentrations)
            {
                if (_warningMessages_CurrentInstance != String.Empty)
                    return EsitoRecuperoInformazioniFoglio.RecuperoConWarnings;

                return EsitoRecuperoInformazioniFoglio.RecuperoCorretto;
            }

            return EsitoRecuperoInformazioniFoglio.RecuperoConErrori;

        }

        /// <summary>
        /// Permette la lettura di tutte le proprieta per l'istanza corrente con annesso il riconoscimento del title da attribuire al contenitore corrente 
        /// </summary>
        /// <param name="propertiesInstances"></param>
        /// <param name="TitleRecognition"></param>
        /// <returns></returns>
        private static bool ReadPropertiesInfo(Excel_PropertiesContainer propertiesInstances, bool TitleRecognition)
        {
            bool riconoscimentoTitle = false;
            bool riconiscimnetoProperties = false;


            // riconoscimento del title 
            if (propertiesInstances.NameInstance != null)
                riconoscimentoTitle = true;


            riconiscimnetoProperties = ReadPropertiesInfo(propertiesInstances);

            // inserimento della validazione per il singono container
            propertiesInstances.ValidatedElem = (riconoscimentoTitle || riconiscimnetoProperties);
            
            // entrambi i valori devono essere stati riconosciuti correttamente 
            return (riconoscimentoTitle || riconiscimnetoProperties);
        }


        /// <summary>
        /// Validazione per le proprietà passate in input
        /// queste proprieta possono essere riferite sia alle leghe che alle concentrazioni per il caso corrente 
        /// </summary>
        /// <param name="propertiesInstances"></param>
        /// <returns></returns>
        private static bool ReadPropertiesInfo(Excel_PropertiesContainer propertiesInstances)
        {
            // passaggio validazione 
            bool currentValidazione = true;

            // iterazione proprieta per la lega correntemente in analisi
            foreach (Excel_PropertyWrapper currentPropertyWrapper in propertiesInstances.PropertiesDefinition)
            {
                // ho trovato istanza per il valore della proprietà
                if (_currentFoglioExcel.Cells[currentPropertyWrapper.Row_Position, currentPropertyWrapper.Col_Position].Value != null)
                {
                    currentPropertyWrapper.StringValue = _currentFoglioExcel.Cells[currentPropertyWrapper.Row_Position, currentPropertyWrapper.Col_Position].Value.ToString();
                }
                // discriminazione per la proprieta opzionale / obbligatoria corrente 
                else
                {
                    if (currentPropertyWrapper.IsOptional)
                        _warningMessages_CurrentInstance += String.Format("riga {0}, colonna {1}: la proprietà opzionale non è stata valorizzata correttamente per la lega", currentPropertyWrapper.Row_Position, currentPropertyWrapper.Col_Position);
                    else
                    {
                        _errorMessages_CurrentInstance += String.Format("riga {0}, colonna {1}: la proprietà opzionale non è stata valorizzata correttamente per la lega", currentPropertyWrapper.Row_Position, currentPropertyWrapper.Col_Position);
                        currentValidazione = false;
                    }

                }
            }

            propertiesInstances.ValidatedElem = currentValidazione;

            return currentValidazione;
        }
        

    }
}
