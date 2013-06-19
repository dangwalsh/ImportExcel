using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Autodesk.Revit.Attributes;
using Autodesk.Revit.UI;

namespace ImportExcel
{
    [Transaction(TransactionMode.Manual)]
    class EACommand : IExternalCommand
    {
        /// <summary>
        /// Command entry point
        /// </summary>
        /// <param name="commandData"></param>
        /// <param name="message"></param>
        /// <param name="elements"></param>
        /// <returns></returns>
        public Result Execute(ExternalCommandData commandData, ref string message, Autodesk.Revit.DB.ElementSet elements)
        {
            try
            {
                
                EADocumentData eadd = new EADocumentData(elements, commandData);
                EASharedParamData easp = new EASharedParamData("Rooms");
                //easp.SetNewParameterToTypeRoom("Display");
                MainWindow mw = new MainWindow();
                mw.ShowDialog();                
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.Read();
                return Result.Failed;
            }
        }
    }


    /// <summary>
    /// Custom file read exception
    /// </summary>
    public class FileReadFailedException : Exception
    {
        public FileReadFailedException()
        {
        }

        public FileReadFailedException(string message)
            : base(message)
        {
        }

        public FileReadFailedException(string message, Exception inner)
            : base(message, inner)
        {
        }
    }


    /// <summary>
    /// Custom file read exception
    /// </summary>
    public class CountFailedException : Exception
    {
        public CountFailedException()
        {
        }

        public CountFailedException(string message)
            : base(message)
        {
        }

        public CountFailedException(string message, Exception inner)
            : base(message, inner)
        {
        }
    }

    /// <summary>
    /// Custom parameter name exception
    /// </summary>
    public class InvalidNameException : Exception
    {
        public InvalidNameException()
        {
        }

        public InvalidNameException(string message)
            : base(message)
        {
        }

        public InvalidNameException(string message, Exception inner)
            : base(message, inner)
        {
        }
    }
}
