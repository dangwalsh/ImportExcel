﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using Autodesk.Revit.ApplicationServices;

namespace ImportExcel
{
    class EADocument
    {
        static private ElementSet _elements;
        static private ExternalCommandData _commandData;

        /// <summary>
        /// Accessor for the Revit Document object
        /// </summary>
        static public Document Doc
        {
            get { return _commandData.Application.ActiveUIDocument.Document; }
        }

        /// <summary>
        /// Accessor for the Revit Application object
        /// </summary>
        static public Application App
        {
            get { return _commandData.Application.Application; }
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="e"></param>
        /// <param name="d"></param>
        public EADocument(ElementSet e, ExternalCommandData d)
        {
            _elements = e;
            _commandData = d;
        }
    }
}
