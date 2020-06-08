using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ToolImportazioneLeghe_Console.Messaging_Console;

namespace ToolImportazioneLeghe_Console.Utils
{
    /// <summary>
    /// Funzioni di utilità globale per il tool corrente 
    /// </summary>
    public class UtilityFunctions
    {
        /// <summary>
        /// Permette di ottenre il nome a partire dal path per il file corrente 
        /// </summary>
        /// <param name="currentFilePath"></param>
        /// <returns></returns>
        public string GetFileName(string currentFilePath)
        {
            string targetFile = String.Empty;
            

            try
            {
                if (currentFilePath == null)
                    throw new Exception(ExceptionMessages.UTILFUNCTIONS_PATHNULL);
                if (currentFilePath == String.Empty)
                    throw new Exception(ExceptionMessages.UTILFUNCTIONS_PATHNULL);

                targetFile = currentFilePath.Substring(currentFilePath.LastIndexOf("\\") + 1);

                if (targetFile == String.Empty)
                    throw new Exception(ExceptionMessages.UTILFUNCTIONS_NOMEVUOTO);

            }
            catch(Exception e)
            {
                throw new Exception(String.Format(ExceptionMessages.UTILFUNCTIONS_ERRORENELLALETTURANOMEFILE, e.Message));
            }
            
            return targetFile;
        }


        /// <summary>
        /// Permette di stabilire il contenimento del primo nome passato in input nel secondo o viceversa 
        /// </summary>
        /// <param name="name_1"></param>
        /// <param name="name_2"></param>
        /// <returns></returns>
        public bool CheckDoubleNamesContainement(string name_1, string name_2)
        {
            if (name_1.Contains(name_2))
                return true;

            if (name_2.Contains(name_1))
                return true;


            return false;
        }


        /// <summary>
        /// Permette di capile se il path passato è valido e nel caso viene validato creando le giuste directory
        /// e poi il file 
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="esistenza"></param>
        public void BuildFilePath(string filePath, out bool esistenza)
        {
            try
            {
                esistenza = false;


                // inizializzazione del log 
                string[] logFilePathParts = filePath.Split('\\');

                string pathRecostrunction = String.Empty;

                string validatedPath = String.Empty;

                foreach (string pathPart in logFilePathParts)
                {

                    if (pathPart == logFilePathParts.First())
                    {
                        pathRecostrunction = pathPart;

                        if (!Directory.Exists(pathRecostrunction))
                            throw new Exception(String.Format(ExceptionMessages.LOGGING_NONHOTROVATOPARTEDELPATH, filePath));
                    }
                    else if (pathPart != logFilePathParts.Last())
                    {
                        pathRecostrunction += "\\" + pathPart;

                        if (!Directory.Exists(pathRecostrunction))
                        {
                            Directory.CreateDirectory(pathRecostrunction);
                            ConsoleService.ConsoleLogging.RicreazioneFolder(pathPart, ServiceLocator.GetUtilityFunctions.GetFileName(filePath));
                        }

                    }
                    else
                    {
                        validatedPath = pathRecostrunction += "\\" + pathPart;
                        if (File.Exists(validatedPath))
                        {
                            
                            esistenza = true;
                        }
                            
                        else
                        {
                            File.Create(validatedPath);
                            ConsoleService.ConsoleLogging.RicreazioneLogFile(pathPart);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                throw new Exception(String.Format(ExceptionMessages.LOGGING_VERIFICADIUNCERTOERRORE, e.Message));
            }
        }
    }
}
