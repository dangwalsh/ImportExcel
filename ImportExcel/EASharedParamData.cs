using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.ApplicationServices;


namespace ImportExcel
{
    class EASharedParamData
    {
        private const string _sharedFilePath = @"L:\4_Revit\SharedPARAM.txt";
        private bool _doesExist = false;
        private string _groupName = "";
        private DefinitionFile _sharedFile;
        private static List<string> _sharedParamNames = new List<string>();

        /// <summary>
        /// Accessor for list of shared parameters
        /// </summary>
        static public List<string> SharedParamNames
        {
            get { return _sharedParamNames; }
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="groupName"></param>
        public EASharedParamData(string groupName)
        {
            _groupName = groupName;
            if (LoadSharedParameterFromFile(out _doesExist) && _doesExist)
            {
                GetSharedParameterList();
            }
        }

        /// <summary>
        /// Method to load shared parameter file into memory
        /// </summary>
        /// <param name="exist"></param>
        /// <returns></returns>
        private bool LoadSharedParameterFromFile(out bool exist)
        {
            exist = true;
            if (!File.Exists(_sharedFilePath))
            {
                exist = false;
                return true;
            }
            EADocumentData.Doc.Application.SharedParametersFilename = _sharedFilePath;
            try
            {
                _sharedFile = EADocumentData.Doc.Application.OpenSharedParameterFile();
            }
            catch (System.Exception e)
            {
                TaskDialog.Show("Error", e.Message);
                return false;
            }
            return true;
        }

        /// <summary>
        /// Method to iterate through shared parameters by group
        /// </summary>
        /// <returns></returns>
        public bool GetSharedParameterList()
        {
            if (File.Exists(_sharedFilePath) && null == _sharedFile)
            {
                TaskDialog.Show("Error", "SharedPARAM.txt has an invalid format.");
                return false;
            }

            DefinitionGroup group = _sharedFile.Groups.get_Item(_groupName);
            if (null == group) return false;

            foreach (Definition definition in group.Definitions)
            {
                _sharedParamNames.Add(definition.Name); 
            }
            return true;
        }

        /// <summary>
        /// Method that finds the existing shared parameter and binds a new project parameter
        /// </summary>
        /// <param name="name"></param>
        /// <param name="cats"></param>
        /// <param name="group"></param>
        /// <param name="inst"></param>
        static public void CreateProjParamFromExistSharedParam(string name, CategorySet cats, BuiltInParameterGroup group, bool inst)
        {
            try
            {
                DefinitionFile defFile = EADocumentData.App.OpenSharedParameterFile();
                if (defFile == null) throw new FileReadFailedException("No SharedParameter File!");

                var v = (from DefinitionGroup dg in defFile.Groups
                         from ExternalDefinition d in dg.Definitions
                         where d.Name == name
                         select d);
                if (v == null || v.Count() < 1) throw new InvalidNameException("Invalid Name Input!");

                ExternalDefinition def = v.First();

                Autodesk.Revit.DB.Binding binding = EADocumentData.App.Create.NewTypeBinding(cats);
                if (inst) binding = EADocumentData.App.Create.NewInstanceBinding(cats);

                BindingMap map = (new UIApplication(EADocumentData.App)).ActiveUIDocument.Document.ParameterBindings;
                map.Insert(def, binding, group);
            }
            catch (InvalidNameException ex)
            {
                TaskDialog.Show("Error", ex.Message);
            }
            catch (FileReadFailedException ex)
            {
                TaskDialog.Show("Error", ex.Message);
            }
        }

        /*
        /// <summary>
        /// Method that creates a new shared parameter and binds a new project parameter
        /// </summary>
        /// <param name="defGroup"></param>
        /// <param name="name"></param>
        /// <param name="type"></param>
        /// <param name="visible"></param>
        /// <param name="cats"></param>
        /// <param name="paramGroup"></param>
        /// <param name="inst"></param>
        static public void CreateProjParamFromNewSharedParam(string defGroup, string name, ParameterType type, bool visible, CategorySet cats, BuiltInParameterGroup paramGroup, bool inst)
        {
            DefinitionFile defFile = EADocumentData.App.OpenSharedParameterFile();
            if (defFile == null) throw new Exception("No SharedParameter File!");

            ExternalDefinition def = EADocumentData.App.OpenSharedParameterFile().Groups.Create(defGroup).Definitions.Create(name, type, visible) as ExternalDefinition;

            Autodesk.Revit.DB.Binding binding = EADocumentData.App.Create.NewTypeBinding(cats);
            if (inst) binding = EADocumentData.App.Create.NewInstanceBinding(cats);

            BindingMap map = (new UIApplication(EADocumentData.App)).ActiveUIDocument.Document.ParameterBindings;
            map.Insert(def, binding, paramGroup);
        }
        */
    }
}
