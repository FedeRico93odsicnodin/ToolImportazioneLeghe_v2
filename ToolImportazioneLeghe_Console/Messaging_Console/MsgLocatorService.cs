using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Resources;
using System.Text;
using System.Threading.Tasks;
using ToolImportazioneLeghe_Console.Utils;

namespace ToolImportazioneLeghe_Console.Messaging_Console
{
    /// <summary>
    /// Questa classe mi serve per localizzare i diversi files di risorsa in base al contenuto del messaggio 
    /// e a chi ne fa uso 
    /// Questo servizio verrà sostituito completamente al console service e consentirà la localizzazione del messaggio 
    /// in base alla lingua oltre che il recupero e la scrittura dell'effettivo messaggio in un apparato console / view 
    /// </summary>
    public static class MsgLocatorService
    {

        #region GETTERS : INIZIALIZZAZIONI CLASSI ALL'OCCORRENZA

        private static CommonEMainService _commonMainService;


        /// <summary>
        /// Servizi resource di main
        /// </summary>
        public static CommonEMainService GetCommonMainService
        {
            get
            {
                if (_commonMainService == null)
                    _commonMainService = new CommonEMainService();

                return _commonMainService;
            }
                
        }

        #endregion



        #region CLASSE LOCAZIONE RISORSE

        /// <summary>
        /// Servizio relativo ai messaggi comuni e di main di apertura di una certa sorgente in lettura / scrittura 
        /// </summary>
        public class CommonEMainService
        {
            
            /// <summary>
            /// File di risorse dal quale andare a prendere le stringhe necessarie alla visualizzazione 
            /// dei messaggi
            /// </summary>
            private ResourceManager _locationRes;


            /// <summary>
            /// Permette di inizializzare il file di risorse utilizzato e l'attore che incorrera nei diversi messaggi presi
            /// da tale file di risorsa
            /// </summary>
            public CommonEMainService()
            {

                // impostazione della lingua in base a quanto inserito all'interno delle costanti
                switch (Constants.LinguaCorrenteTool)
                {
                    case Constants.LinguaSelezionata.ENG:
                        {
                            _locationRes = new ResourceManager("ToolImportazioneLeghe_Console.Messaging_Console.ResourceMsgString.CommonsEMain_ENG", Assembly.GetExecutingAssembly());
                            break;
                        }
                    default:
                        {
                            _locationRes = new ResourceManager("ToolImportazioneLeghe_Console.Messaging_Console.ResourceMsgString.CommonsEMain_ITA", Assembly.GetExecutingAssembly());
                            break;
                        }
                }


            }


            /// <summary>
            /// Segnalazione relativa al fatto che è impostata l'opzione sul tempo totale di procedura 
            /// questo tempo viene fatto partire all'interno della classe di MAIN
            /// </summary>
            public void MAIN_InizializzazioneTempoProcedura()
            {
                string currentMessage = String.Format(_locationRes.GetString("AvviamentoTempoProcedura"));
                ConsoleService.FormatMessageConsole(currentMessage, true);
            }
        }


        /// <summary>
        /// Servizio relativo allo step 1 di apertura di una certa sorgente in lettura / scrittura
        /// </summary>
        public class OpenSourceService
        {

        }

        
        /// <summary>
        /// Servizio relativo allo step 2 di riconoscimento delle informazioni su una certa sorgente in apertura scrittura
        /// </summary>
        public class RecognizeSourceService
        {

        }


        /// <summary>
        /// Servizio relativo allo step 3 di validazione delle informazioni necessarie su una certa sorgente destinazione
        /// </summary>
        public class ValidateInfoSourceService
        {
            /// <summary>
            /// Risorsa in utilizzo per il caso corrente 
            /// </summary>
            private string _sourceKind;


            /// <summary>
            /// Tipo di azione che si sta eseguendo su una particolare fonte dati
            /// </summary>
            private string _typeAction;


            /// <summary>
            /// File di risorse dal quale andare a prendere le stringhe necessarie alla visualizzazione 
            /// dei messaggi
            /// </summary>
            private ResourceManager _locationRes;


            /// <summary>
            /// Istanziamento delle proprieta relative alla formattazione standard in base a dove 
            /// leggo scrivo e al tipo di risorsa che sto utilizzando
            /// </summary>
            /// <param name="sourceKind"></param>
            /// <param name="typeAction"></param>
            public ValidateInfoSourceService(Constants.ResourceTypes sourceKind, Constants.ModalitaAperturaExcel typeAction)
            {
                this._sourceKind = sourceKind.ToString();
                this._typeAction = typeAction.ToString();

                // impostazione della lingua in base a quanto inserito all'interno delle costanti
                switch(Constants.LinguaCorrenteTool)
                {
                    case Constants.LinguaSelezionata.ENG:
                        {
                            _locationRes = new ResourceManager("ToolImportazioneLeghe_Console.Messaging_Console.ResourceMsgString.Validators_ENG", Assembly.GetExecutingAssembly());
                            break;
                        }
                    default:
                        {
                            _locationRes = new ResourceManager("ToolImportazioneLeghe_Console.Messaging_Console.ResourceMsgString.Validators_ITA", Assembly.GetExecutingAssembly());
                            break;
                        }
                }


            }


            /// <summary>
            /// Segnalazione di inizio della procedura di validazione delle informazioni per l'elemento correntemente in analisi
            /// </summary>
            public void InizioProceduraValidazioneInformazioni()
            {
                string currentMessage = String.Format(_locationRes.GetString("InizioProceduraValidazioneInformazioni"), _sourceKind, _typeAction);

            }

        }


        /// <summary>
        /// Servizio relativo allo step 4 di scrittura delle informazioni necessarie su una certa destinazione 
        /// </summary>
        public class WriteInfoSourceService
        {

        }


        #endregion
    }
}
