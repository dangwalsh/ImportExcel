using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Autodesk.Revit.DB;
using Autodesk.Revit.ApplicationServices;

namespace ImportExcel
{
    class EAParameterFactory
    {
        public EAParameterFactory(string group, string name) // group name should be rooms
        {
            EAParameterBinding paramBind = new EAParameterBinding(group, name);
        }
    }
}
