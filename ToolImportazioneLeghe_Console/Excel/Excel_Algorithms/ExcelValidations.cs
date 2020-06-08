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
    /// Validazione inerente le diverse informazioni per l'excel corrente e che sono fatte nello stesso ambito excel 
    /// (senza il supporto nel recupero di eventuali valori da confrontare nella destinazione)
    /// </summary>
    public static class ExcelValidations
    {
        /// <summary>
        /// Messaggio di errori relativo alla validazione eseguita in questo step
        /// </summary>
        private static string _error_Message_Validation = String.Empty;


        /// <summary>
        /// Messaggio di warning relativo alla validazione eseguita in questo step
        /// </summary>
        private static string _warning_Message_Validation = String.Empty;


        /// <summary>
        /// Lista dei fogli di lega in analisi corrente 
        /// </summary>
        private static List<Excel_AlloyInfo_Sheet> _fogliLegheAnalisiCorrente;


        /// <summary>
        /// Lista dei fogli di concentrazione in analisi corrente 
        /// </summary>
        private static List<Excel_AlloyInfo_Sheet> _fogliConcentrazioniAnalisiCorrente;


        /// <summary>
        /// Nome del foglio di lega in analisi corrente e rispetto al quale si sta eseguendo 
        /// il match delle informazioni
        /// </summary>
        private static string _nomeFoglioLegheAnalisiCorrente = String.Empty;


        /// <summary>
        /// Lista dei fogli di concentrazioni in analisi temporanea sui quali riconoscere rispetto alle proprieta di lega 
        /// contenute nel foglio di leghe in iterazione corrente 
        /// </summary>
        private static List<Excel_AlloyInfo_Sheet> _fogliConcentrazioniTEMPAnalysis;


        /// <summary>
        /// Foglio di leghe sul quale si basa l'analisi corrente 
        /// </summary>
        private static Excel_AlloyInfo_Sheet _foglioLegheTEMPAnalysis;



        /// <summary>
        /// Validazione per il primo formato excel individuato
        /// </summary>
        public static class ValidateFormat_1
        {
            /// <summary>
            /// Permette di ordinare la lista dei fogli e delle informazioni lette attraverso il significato proprio dei nomi attribuiti.
            /// Per questi nomi viene trovata una parte comune e utilizzata per associare le relative informazioni di lega e concentrazioni
            /// STEPS:
            /// 1) ordinamento dei fogli in base alla nomenclatura trovata 
            /// 2) ordinamento e validazione delle informazioni in essi contenuti (validazione relativa ai diversi quadranti di cui sarà possibile la lettura in base 
            /// all'associazione fatta sul nome)
            /// 
            /// Come per gli altri casi precedenti il significato del valore restituito è il seguente 
            /// - recupero con errori: se non si riesce a trovare NESSUNA ASSOCIAZIONE DA INSERIRE
            /// - recupero con warnings: se alcune delle informazioni non possono essere inserite in base ai criteri esposti
            /// - recupero corretto: se tutte le informazioni sono state associate in maniera corretta sul foglio (quindi tutti i valori sono potenzialmente inseribili)
            /// </summary>
            /// <param name="alreadyReadSheets"></param>
            /// <param name="orderedByCategorySheets"></param>
            /// <param name="generalErrorMessage"></param>
            /// <param name="generalWarningMessage"></param>
            public static Constants_Excel.EsitoRecuperoInformazioniFoglio GroupSheetsByNameAssociation(
                List<Excel_AlloyInfo_Sheet> alreadyReadSheets, 
                out List<Excel_AlloyInfo_Sheet> orderedByCategorySheets, 
                out string generalErrorMessage, 
                out string generalWarningMessage)
            {
                // lista ordinata dei fogli 
                orderedByCategorySheets = new List<Excel_AlloyInfo_Sheet>();

                // inizializzazione possibili errori warnings per il caso corrente 
                _error_Message_Validation = String.Empty;
                _warning_Message_Validation = String.Empty;
                generalErrorMessage = String.Empty;
                generalWarningMessage = String.Empty;


                // nome per il foglio delle proprieta di lega in analisi corrente 
                _nomeFoglioLegheAnalisiCorrente = String.Empty;

                // counter per nessuna corrispondenza rispetto ai fogli di concentrazione a disposizione
                int counterMissedAlloyCorrenspondance = 0;


                // suddivisione delle 2 categorie di fogli in analisi
                if(alreadyReadSheets.Where(x => x.GetTipologiaFoglio == Constants_Excel.TipologiaFoglio_Format.FoglioLeghe).Count() > 0 &&
                    alreadyReadSheets.Where(x => x.GetTipologiaFoglio == Constants_Excel.TipologiaFoglio_Format.FoglioConcentrazioni).Count() > 0)
                {
                    _fogliLegheAnalisiCorrente = alreadyReadSheets.Where(x => x.GetTipologiaFoglio == Constants_Excel.TipologiaFoglio_Format.FoglioLeghe).ToList();
                    _fogliConcentrazioniAnalisiCorrente = alreadyReadSheets.Where(x => x.GetTipologiaFoglio == Constants_Excel.TipologiaFoglio_Format.FoglioConcentrazioni).ToList();

                }
                else
                {
                    // TODO : sollevamento eccezione mancanta di uno o altro set
                }

                // iterazione fogli correnti
                foreach (Excel_AlloyInfo_Sheet currentSheet in alreadyReadSheets)
                {
                    // verifica presenza fogli concentrazioni per il caso corrente 
                    if (VerifyFogliConcentrazioniPresence(currentSheet.GetSheetName))
                    {

                    }
                    else
                    {
                        counterMissedAlloyCorrenspondance++;
                    }
                        
                }


                orderedByCategorySheets = alreadyReadSheets;

                return Constants_Excel.EsitoRecuperoInformazioniFoglio.RecuperoConErrori;
            }


            /// <summary>
            /// In base alla validazione e alle associazioni precedenti ci si potrebbe trovare nella situazione in cui, nel caso in cui l'associazione sia fatta correttamente 
            /// o con dei warnings. potrebbe ancora essere possibile una qualche esclusione delle informazioni perché non validate correttamente sul foglio in considerazione
            /// (foglio degli elementi: ad esempio non trovo nessuna corrispondenza relativa agli elementi inseriti o non trovo nessuna corrispondenza per quanto riguarda la conversione dei valori)
            /// </summary>
            /// <param name="alreadyReadConcentrationsSheets"></param>
            /// <param name="validatedConcentrationsSheets"></param>
            /// <returns></returns>
            public static Constants_Excel.EsitoRecuperoInformazioniFoglio ValidateElementsInfo(List<Excel_AlloyInfo_Sheet> alreadyReadConcentrationsSheets, out List<Excel_AlloyInfo_Sheet> validatedConcentrationsSheets)
            {
                validatedConcentrationsSheets = alreadyReadConcentrationsSheets;

                return Constants_Excel.EsitoRecuperoInformazioniFoglio.RecuperoConErrori;
            }


            /// <summary>
            /// Permette di associare alcune delle informazioni relative alla futura popolazione e confronto con i valori già presenti a database per quanto riguarda la lettura delle informazioni contenute
            /// nell'excel in merito a valori già contenuti per le leghe (ad esempio la categoria e la base, queste devono corrispondere alla associazione che si farà con un valore presente a database o meno)
            /// </summary>
            /// <param name="alreadyReadAlloyPropertiesSheets"></param>
            /// <param name="preAssociatedValuesOnSheets"></param>
            /// <returns></returns>
            public static Constants_Excel.EsitoRecuperoInformazioniFoglio PreAssociationsAlloyProperties(List<Excel_AlloyInfo_Sheet> alreadyReadAlloyPropertiesSheets, out List<Excel_AlloyInfo_Sheet> preAssociatedValuesOnSheets)
            {
                preAssociatedValuesOnSheets = alreadyReadAlloyPropertiesSheets;

                return Constants_Excel.EsitoRecuperoInformazioniFoglio.RecuperoConErrori;
            }


            #region PROPRIETA PRIVATE

            /// <summary>
            /// Permette di stabilire come prima validazione se esiste o meno un certo numero di fogli di concentrazioni per il foglio di leghe 
            /// in analisi corrente 
            /// </summary>
            /// <param name="currentSheetAlloyName"></param>
            /// <returns></returns>
            private static bool VerifyFogliConcentrazioniPresence(string currentSheetAlloyName)
            {
                // inizializzazione delle 2 liste relative ai diversi fogli in analisi per le concentrazioni
                _fogliConcentrazioniTEMPAnalysis = new List<Excel_AlloyInfo_Sheet>();

                // variabile che mi dice di aver individuato almeno un foglio di concentrazioni valido per le leghe in analisi corrente 
                bool hoTrovatoAlmenoUnFoglioConcentrazioni = false;

                foreach (Excel_AlloyInfo_Sheet foglioConcentrazioni in _fogliConcentrazioniAnalisiCorrente)
                {
                    // verifica presenza di un foglio di concentrazioni che abbia corrispondenza per il foglio leghe corrente
                    if(ServiceLocator.GetUtilityFunctions.CheckDoubleNamesContainement(currentSheetAlloyName, foglioConcentrazioni.GetSheetName))
                    {
                        _fogliConcentrazioniTEMPAnalysis.Add(foglioConcentrazioni);
                        hoTrovatoAlmenoUnFoglioConcentrazioni = true;
                    }
                }


                return hoTrovatoAlmenoUnFoglioConcentrazioni;
            }

            #endregion
        }


        /// <summary>
        /// Validazione per il secondo formato excel individuato
        /// </summary>
        public static class ValidateFormat2
        {
            /// <summary>
            /// Permette la validazione delle informazioni di elemento per il foglio del secondo formato in lettura
            /// questo metodo puo essere tenuto separato dalla definizione di eventuali altri fogli in quanto le informazioni rimarrebbero comunque separate dai 2 casi
            /// vengono controllati rispettivamente i diversi elementi definiti e i valori numerici attribuiti alle diverse concentrazioni
            /// </summary>
            /// <param name="alreadyReadInformationSheet"></param>
            /// <param name="validatedOnElementsInfoSheet"></param>
            /// <returns></returns>
            public static Constants_Excel.EsitoRecuperoInformazioniFoglio ValidateElementsInfo(Excel_AlloyInfo_Sheet alreadyReadInformationSheet, out Excel_AlloyInfo_Sheet validatedOnElementsInfoSheet) 
            {
                validatedOnElementsInfoSheet = alreadyReadInformationSheet;

                return Constants_Excel.EsitoRecuperoInformazioniFoglio.RecuperoConErrori;
            }


            /// <summary>
            /// Permette la popolazione di alcune informazioni che caratterizzano la lega in considerazione, queste informazioni riguardano principalmente categoria e lega che si andranno ad utilizzare di riferimento
            /// all'interno del database 
            /// </summary>
            /// <param name="alreadyReadInformationSheet"></param>
            /// <param name="preAssociatedInfoOnSheet"></param>
            /// <returns></returns>
            public static Constants_Excel.EsitoRecuperoInformazioniFoglio PreAssociationsAlloyProperties(Excel_AlloyInfo_Sheet alreadyReadInformationSheet, out Excel_AlloyInfo_Sheet preAssociatedInfoOnSheet)
            {
                preAssociatedInfoOnSheet = alreadyReadInformationSheet;

                return Constants_Excel.EsitoRecuperoInformazioniFoglio.RecuperoConErrori;
            }
        }


        /// <summary>
        /// In questa classe saranno contenuti tutti gli algoritmi comuni alle 2 classi su lettura dei diversi formati in merito alle validazioni indispensabili per le informazioni
        /// (comunque di uguale natura) sui 2 formati
        /// </summary>
        static class CommonAlgorithmsOfValidation
        {

        }
    }
}
