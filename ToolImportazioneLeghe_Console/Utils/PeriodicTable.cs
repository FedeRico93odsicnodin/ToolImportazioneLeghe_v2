using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ToolImportazioneLeghe_Console.Utils
{
    /// <summary>
    /// Classe elemento
    /// </summary>
    public class Element
    {
        public int AtomicNumber { get; set; }
        public string Symbol { get; set; }
        public string Name { get; set; }
        public decimal AtomicWeight { get; set; }
        public string NameIT { get; set; }
        //  public string GroupNumber { get; set; }
        //  public string GroupName { get; set; }
        //  public string Period { get; set; }
        //  public string Block { get; set; }
        //  public string CASRegistryID { get; set; }
        //  public string DiscoveryDate { get; set; }
        //  public string DiscovererName { get; set; }

        public Element() { }

        /// <summary>
        /// Proprieta valorizzate per i diversi elementi, comprendendo anche il nome da attribuire in italiano
        /// </summary>
        /// <param name="atomicNumber"></param>
        /// <param name="symbol"></param>
        /// <param name="name"></param>
        /// <param name="nameITA"></param>
        /// <param name="atomicWeight"></param>
        public Element(int atomicNumber, string nameITA, string symbol, string name, decimal atomicWeight)
        {
            AtomicNumber = atomicNumber;
            Symbol = symbol;
            Name = name;
            AtomicWeight = atomicWeight;
            
        }
    }


    /// <summary>
    /// Mappatura di tutti gli elementi della tavola periodica comprendendo numero - sigla - nome italiano / latino
    /// </summary>
    public static class PeriodicTable
    {
        /// <summary>
        /// Lista degli elementi correnti con le definizioni date per la classe precedente 
        /// </summary>
        private static List<Element> _listaElementi;


        /// <summary>
        /// Lista di tutti gli elementi per l'analisi
        /// </summary>
        public static List<Element> Elements
        {
            get
            {
                // inizializzazione della lista degli elementi
                if (_listaElementi == null)
                {
                    _listaElementi = new List<Element>();
                    _listaElementi.Add(new Element(1,   "Idrogeno",                 "H", "Hydrogen", 1.007825M));
                    _listaElementi.Add(new Element(2,    "Elio",                 "He", "Helium", 4.00260M));
                    _listaElementi.Add(new Element(3,    "Litio",                 "Li", "Lithium", 6.941M));
                    _listaElementi.Add(new Element(4,    "Berillio",                 "Be", "Beryllium", 9.01218M));
                    _listaElementi.Add(new Element(5,    "Boro",                 "B", "Boron", 10.81M));
                    _listaElementi.Add(new Element(6,    "Carbonio",                 "C", "Carbon", 12.011M));
                    _listaElementi.Add(new Element(7,    "Azoto",                 "N", "Nitrogen", 14.0067M));
                    _listaElementi.Add(new Element(8,    "Ossigeno",                 "O", "Oxygen", 15.999M));
                    _listaElementi.Add(new Element(9,    "Fluoro",                 "F", "Fluorine", 18.99840M));
                    _listaElementi.Add(new Element(10,   "Neon",                 "Ne", "Neon", 20.179M));
                    _listaElementi.Add(new Element(11,   "Sodio",                 "Na", "Sodium", 22.98977M));
                    _listaElementi.Add(new Element(12,   "Magnesio",                "Mg", "Magnesium", 24.305M));
                    _listaElementi.Add(new Element(13,   "Alluminio",                 "Al", "Aluminum", 26.98154M));
                    _listaElementi.Add(new Element(14,   "Silicio",                 "Si", "Silicon", 28.0855M));
                    _listaElementi.Add(new Element(15,   "Fosforo",                 "P", "Phosphorus", 0.0M));
                    _listaElementi.Add(new Element(16,   "Zolfo",                 "S", "Sulphur", 32.06M));
                    _listaElementi.Add(new Element(17,   "Cloro",                 "Cl", "Chlorine", 35.453M));
                    _listaElementi.Add(new Element(18,   "Argon",                 "Ar", "Argon", 39.948M));
                    _listaElementi.Add(new Element(19,   "Potassio",                 "K", "Potassium", 39.0983M));
                    _listaElementi.Add(new Element(20,   "Calcio",                 "Ca", "Calcium", 40.08M));
                    _listaElementi.Add(new Element(21,   "Scandio",                 "Sc", "Scandium", 44.9559M));
                    _listaElementi.Add(new Element(22,   "Titanio",                 "Ti", "Titanium", 47.90M));
                    _listaElementi.Add(new Element(23,   "Vanadio",                 "V", "Vanadium", 50.9414M));
                    _listaElementi.Add(new Element(24,   "Cromo",                 "Cr", "Chromium", 51.996M));
                    _listaElementi.Add(new Element(25,   "Manganese",                 "Mn", "Manganese", 54.9380M));
                    _listaElementi.Add(new Element(26,   "Ferro",                 "Fe", "Iron", 55.85M));
                    _listaElementi.Add(new Element(27,   "Cobalto",                 "Co", "Cobalt", 58.9332M));
                    _listaElementi.Add(new Element(28,   "Nichel",                 "Ni", "Nickel", 58.71M));
                    _listaElementi.Add(new Element(29,   "Rame",                 "Cu", "Copper", 63.546M));
                    _listaElementi.Add(new Element(30,   "Zinco",                 "Zn", "Zinc", 65.37M));
                    _listaElementi.Add(new Element(31,   "Gallio",             "Ga", "Gallium", 69.72M));
                    _listaElementi.Add(new Element(32,   "Germanio",                 "Ge", "Germanium", 72.59M));
                    _listaElementi.Add(new Element(33,   "Arsenico",                 "As", "Arsenic", 74.9216M));
                    _listaElementi.Add(new Element(34,   "Selenio",                 "Se", "Selenium", 78.96M));
                    _listaElementi.Add(new Element(35,   "Bromo",                 "Br", "Bromine", 79.904M));
                    _listaElementi.Add(new Element(36,   "Cripto",                 "Kr", "Krypton", 83.80M));
                    _listaElementi.Add(new Element(37,   "Rubidio",                 "Rb", "Rubidium", 85.4678M));
                    _listaElementi.Add(new Element(38,   "Stronzio",                 "Sr", "Strontium", 87.62M));
                    _listaElementi.Add(new Element(39,   "Yttrio",                 "Y", "Yttrium", 88.9059M));
                    _listaElementi.Add(new Element(40,   "Zirconio",                 "Zr", "Zirconium", 91.22M));
                    _listaElementi.Add(new Element(41,   "Niobio",                 "Nb", "Niobium", 92.91M));
                    _listaElementi.Add(new Element(42,   "Molibdeno",                 "Mo", "Molybdenum", 95.94M));
                    _listaElementi.Add(new Element(43,   "Tecnezio",                 "Tc", "Technetium", 99.0M));
                    _listaElementi.Add(new Element(44,   "Rutenio",                 "Ru", "Ruthenium", 101.1M));
                    _listaElementi.Add(new Element(45,   "Rodio",                "Rh", "Rhodium", 102.91M));
                    _listaElementi.Add(new Element(46,   "Palladio",                 "Pd", "Palladium", 106.42M));
                    _listaElementi.Add(new Element(47,   "Argento",                 "Ag", "Silver", 107.87M));
                    _listaElementi.Add(new Element(48,   "Cadmio",                 "Cd", "Cadmium", 112.4M));
                    _listaElementi.Add(new Element(49,   "Indio",                 "In", "Indium", 114.82M));
                    _listaElementi.Add(new Element(50,   "Stagno",                 "Sn", "Tin", 118.69M));
                    _listaElementi.Add(new Element(51,   "Antimonio",                 "Sb", "Antimony", 121.75M));
                    _listaElementi.Add(new Element(52,   "Tellurio",                 "Te", "Tellurium", 127.6M));
                    _listaElementi.Add(new Element(53,   "Iodio",                 "I", "Iodine", 126.9045M));
                    _listaElementi.Add(new Element(54,   "Xenon",                 "Xe", "Xenon", 131.29M));
                    _listaElementi.Add(new Element(55,   "Cesio",                 "Cs", "Cesium", 132.9054M));
                    _listaElementi.Add(new Element(56,   "Bario",                 "Ba", "Barium", 137.33M));
                    _listaElementi.Add(new Element(57,   "Lantanio",                 "La", "Lanthanum", 138.91M));
                    _listaElementi.Add(new Element(58,   "Cerio",                 "Ce", "Cerium", 140.12M));
                    _listaElementi.Add(new Element(59,   "Praseodimio",                 "Pr", "Praseodymium", 140.91M));
                    _listaElementi.Add(new Element(60,   "Neodimio",                 "Nd", "Neodymium", 0.0M));
                    _listaElementi.Add(new Element(61,   "Promezio",                 "Pm", "Promethium", 147.0M));
                    _listaElementi.Add(new Element(62,   "Samario",                 "Sm", "Samarium", 150.35M));
                    _listaElementi.Add(new Element(63,   "Europio",                 "Eu", "Europium", 167.26M));
                    _listaElementi.Add(new Element(64,   "Gadolinio",                 "Gd", "Gadolinium", 157.25M));
                    _listaElementi.Add(new Element(65,   "Terbio",                 "Tb", "Terbium", 158.925M));
                    _listaElementi.Add(new Element(66,   "Disprosio",                 "Dy", "Dysprosium", 162.50M));
                    _listaElementi.Add(new Element(67,   "Olmio",                 "Ho", "Holmium", 164.9M));
                    _listaElementi.Add(new Element(68,   "Erbio",                 "Er", "Erbium", 167.26M));
                    _listaElementi.Add(new Element(69,   "Tulio",                 "Tm", "Thulium", 168.93M));
                    _listaElementi.Add(new Element(70,   "Ytterbio",                 "Yb", "Ytterbium", 173.04M));
                    _listaElementi.Add(new Element(71,   "Lutezio",                 "Lu", "Lutetium", 174.97M));
                    _listaElementi.Add(new Element(72,   "Afnio",                 "Hf", "Hafnium", 178.49M));
                    _listaElementi.Add(new Element(73,   "Tantalo",                "Ta", "Tantalum", 180.95M));
                    _listaElementi.Add(new Element(74,   "Tungsteno",                 "W", "Tungsten", 183.85M));
                    _listaElementi.Add(new Element(75,   "Renio",                 "Re", "Rhenium", 186.23M));
                    _listaElementi.Add(new Element(76,   "Osmio",                 "Os", "Osmium", 190.2M));
                    _listaElementi.Add(new Element(77,   "Iridio",                 "Ir", "Iridium", 192.2M));
                    _listaElementi.Add(new Element(78,   "Platino",                 "Pt", "Platinum", 195.09M));
                    _listaElementi.Add(new Element(79,   "Oro",                 "Au", "Gold", 196.9655M));
                    _listaElementi.Add(new Element(80,   "Mercurio",                "Hg", "Mercury", 200.59M));
                    _listaElementi.Add(new Element(81,   "Tallio",                 "Tl", "Thallium", 204.383M));
                    _listaElementi.Add(new Element(82,   "Piombo",                 "Pb", "Lead", 207.2M));
                    _listaElementi.Add(new Element(83,   "Bismuto",                 "Bi", "Bismuth", 208.9804M));
                    _listaElementi.Add(new Element(84,   "Polonio",                 "Po", "Polonium", 210.0M));
                    _listaElementi.Add(new Element(85,   "Astato",                 "At", "Astatine", 210.0M));
                    _listaElementi.Add(new Element(86,   "Radon",                 "Rn", "Radon", 222.0M));
                    _listaElementi.Add(new Element(87,   "Francio",                 "Fr", "Francium", 233.0M));
                    _listaElementi.Add(new Element(88,   "Radio",                 "Ra", "Radium", 226.0254M));
                    _listaElementi.Add(new Element(89,   "Attinio",                 "Ac", "Actinium", 227.0M));
                    _listaElementi.Add(new Element(90,   "Torio",                 "Th", "Thorium", 232.04M));
                    _listaElementi.Add(new Element(91,   "Protoattini",                "Pa", "Protactinium", 231.0359M));
                    _listaElementi.Add(new Element(92,   "Uranio",                 "U", "Uranium", 238.03M));
                    _listaElementi.Add(new Element(93,   "Neptunio",                 "Np", "Neptunium", 237.0M));
                    _listaElementi.Add(new Element(94,   "Plutonio",                 "Pu", "Plutonium", 244.0M));
                    _listaElementi.Add(new Element(95,   "Americio",                 "Am", "Americium", 243.0M));
                    _listaElementi.Add(new Element(96,   "Curio",                 "Cm", "Curium", 247.0M));
                    _listaElementi.Add(new Element(97,   "Berchelio",                 "Bk", "Berkelium", 247.0M));
                    _listaElementi.Add(new Element(98,   "Californio",                 "Cf", "Californium", 251.0M));
                    _listaElementi.Add(new Element(99,   "Einsteinio",                 "Es", "Einsteinium", 254.0M));
                    _listaElementi.Add(new Element(100,  "Fermio",                "Fm", "Fermium", 257.0M));
                    _listaElementi.Add(new Element(101,  "Mendelevio",                 "Md", "Mendelevium", 258.0M));
                    _listaElementi.Add(new Element(102,  "Nobelio",                 "No", "Nobelium", 259.0M));
                    _listaElementi.Add(new Element(103,  "Laurenzio",                 "Lr", "Lawrencium", 262.0M));
                    _listaElementi.Add(new Element(104,  "Ruterfordio",                 "Rf", "Rutherfordium", 260.9M));
                    _listaElementi.Add(new Element(105,  "Dubnio",                 "Db", "Dubnium", 261.9M));
                    _listaElementi.Add(new Element(106,  "Seaborgio",                 "Sg", "Seaborgium", 262.94M));
                    _listaElementi.Add(new Element(107,  "Bohrio",                 "Bh", "Bohrium", 262.0M));
                    _listaElementi.Add(new Element(108,  "Hassio",                 "Hs", "Hassium", 264.8M));
                    _listaElementi.Add(new Element(109,  "Meitnerio",                 "Mt", "Meitnerium", 265.9M));
                    _listaElementi.Add(new Element(110,  "Darmstadtio",                 "Ds", "Darmstadtium", 261.9M));
                    _listaElementi.Add(new Element(112,  "Roentgenium",                 "Uub", "Ununbium", 276.8M));
                    _listaElementi.Add(new Element(114,  "Copernicium",                 "Uuq", "Ununquadium", 289.0M));
                    _listaElementi.Add(new Element(116,  "Nihonium",                 "Uuh", "Ununhexium", 0.0M));
                                                         //Flerovium
                                                         //Moscovium
                                                         //Livermorium
                                                         //Tennessine
                                                         //Oganesson
                      
                }

                return _listaElementi;
            }
        }


    }


    /// <summary>
    /// Mappatura di tutte le parole chiave relative al riconoscimento delle proprieta di lega 
    /// per la lega corrente 
    /// </summary>
    public static class Helpers_AlloyRecognition
    {
        /// <summary>
        /// Lista di tutti gli identificativi per la base ferro
        /// </summary>
        private static List<string> _ironMarkers = new List<string>()
        {
            "Leghe ferrose",
            "Acciai inossidabili e resistenti al calore",
            "Acciai colati",

        };


        /// <summary>
        /// Lista di tutti gli identificativi per la base Oro
        /// </summary>
        public static List<string> _goldMarkers = new List<string>()
        {
            "Metalli preziosi e loro leghe (Ag, Au, Pt, Pd)"
        };


        /// <summary>
        /// Ritorno l'ipotetico elemento a partire dalla descrizione mappata all'interno di questo metodo e letta sul foglio excel corrente 
        /// questo metodo deve essere utilizzato nella validazione solo se la proprieta in questione relativa alla lega non è stata riconosciuta correttamente 
        /// per mezzo di un elemento 
        /// </summary>
        /// <param name="typeDescription"></param>
        /// <returns></returns>
        public static Element GetElementFromTypeDescription(string typeDescription)
        {
            #region RICONOSCIMENTO MATERIALE FERROSO 

            foreach(string ironMarker in _ironMarkers) 
                if(typeDescription.Contains(ironMarker)) { return PeriodicTable.Elements.Where(x => x.AtomicNumber == 26).FirstOrDefault(); }

            #endregion


            #region RICONOSCIMENTO LEGA ORO

            foreach(string goldMarker in _goldMarkers)
                if(typeDescription.Contains(goldMarker)) { return PeriodicTable.Elements.Where(x => x.AtomicNumber == 79).FirstOrDefault(); }

            #endregion

            // nel caso non sia stato riconosciuto nessun tipo di elemento per la lega ritorno null
            return null;

        }
    }
}
