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
        //string _paramName = "";
        private DefinitionFile _sharedFile;
        private static List<string> _sharedParamNames = new List<string>();
        //FamilyManager _manager = null;
        //Document _doc;
        //Application _app;

        static public List<string> SharedParamNames
        {
            get { return _sharedParamNames; }
        }

        public EASharedParamData(string groupName)
        {
            //_doc = EADocumentData.Doc;
            //_app = _doc.Application;
            //_manager = _doc.FamilyManager;
            _groupName = groupName;

            if (LoadSharedParameterFromFile(out _doesExist) && _doesExist)
            {
                //bool result = AddSharedParameterToFamily();
                GetSharedParameterList();
            }
        }

        //public EAsharedParamData(string groupName, string paramName)
        //{
        //    _doc = EADocumentData.Doc;
        //    _app = _doc.Application;
        //    _manager = _doc.FamilyManager;
        //    _groupName = groupName;
        //    _paramName = paramName;

        //    if (LoadSharedParameterFromFile(out _doesExist) && _doesExist)
        //    {
        //        bool result = AddSharedParameterToFamily();
        //    }
        //}

        //public EAParameterBinding(Document doc, string groupName, string paramName)
        //{
        //    _doc = doc;
        //    _app = doc.Application;
        //    _manager = doc.FamilyManager;
        //    _groupName = groupName;
        //    _paramName = paramName;

        //    if (LoadSharedParameterFromFile(out _doesExist) && _doesExist)
        //    {
        //        bool result = AddSharedParameter();
        //    }
        //}

        public bool SetNewParameterToTypeRoom(string paramName)
        {
            // Create a new group in the shared parameters file
            DefinitionGroup group = _sharedFile.Groups.get_Item(_groupName);
            if (null == group) return false;
            ExternalDefinition def = group.Definitions.get_Item(paramName) as ExternalDefinition;
            if (null == def) return false;
            // Create a type definition
            Definition myDefinition_CompanyName = group.Definitions.Create(paramName, ParameterType.Text);
            // Create a category set and insert category of wall to it
            CategorySet myCategories = EADocumentData.App.Create.NewCategorySet();
            // Use BuiltInCategory to get category of wall
            Category myCategory = EADocumentData.Doc.Settings.Categories.get_Item(BuiltInCategory.OST_Rooms);
            myCategories.Insert(myCategory);
            //Create an object of TypeBinding according to the Categories
            TypeBinding typeBinding = EADocumentData.App.Create.NewTypeBinding(myCategories);
            // Get the BingdingMap of current document.
            BindingMap bindingMap = EADocumentData.Doc.ParameterBindings;
            // Bind the definitions to the document
            bool typeBindOK = bindingMap.Insert(myDefinition_CompanyName, typeBinding, BuiltInParameterGroup.PG_TEXT);

            return typeBindOK;
        }

        private bool LoadSharedParameterFromFile(out bool exist)
        {
            exist = true;

            // verify that the file exists before continuing
            if (!File.Exists(_sharedFilePath))
            {
                exist = false;
                return true;
            }

            // if the file exists make it the current shared parameter file
            EADocumentData.Doc.Application.SharedParametersFilename = _sharedFilePath;

            try
            {
                // access the shared parameter file
                _sharedFile = EADocumentData.Doc.Application.OpenSharedParameterFile();
            }
            catch (System.Exception e)
            {
                TaskDialog.Show("Error", e.Message);
                return false;
            }

            return true;
        }

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

        static public void RawCreateProjectParameterFromExistingSharedParameter(string name, CategorySet cats, BuiltInParameterGroup group, bool inst)
        {
            DefinitionFile defFile = EADocumentData.App.OpenSharedParameterFile();
            if (defFile == null) throw new Exception("No SharedParameter File!");

            var v = (from DefinitionGroup dg in defFile.Groups
                     from ExternalDefinition d in dg.Definitions
                     where d.Name == name
                     select d);
            if (v == null || v.Count() < 1) throw new Exception("Invalid Name Input!");

            ExternalDefinition def = v.First();

            Autodesk.Revit.DB.Binding binding = EADocumentData.App.Create.NewTypeBinding(cats);
            if (inst) binding = EADocumentData.App.Create.NewInstanceBinding(cats);

            BindingMap map = (new UIApplication(EADocumentData.App)).ActiveUIDocument.Document.ParameterBindings;
            map.Insert(def, binding, group);
        }

        static public void RawCreateProjectParameterFromNewSharedParameter(string defGroup, string name, ParameterType type, bool visible, CategorySet cats, BuiltInParameterGroup paramGroup, bool inst)
        {
            DefinitionFile defFile = EADocumentData.App.OpenSharedParameterFile();
            if (defFile == null) throw new Exception("No SharedParameter File!");

            ExternalDefinition def = EADocumentData.App.OpenSharedParameterFile().Groups.Create(defGroup).Definitions.Create(name, type, visible) as ExternalDefinition;

            Autodesk.Revit.DB.Binding binding = EADocumentData.App.Create.NewTypeBinding(cats);
            if (inst) binding = EADocumentData.App.Create.NewInstanceBinding(cats);

            BindingMap map = (new UIApplication(EADocumentData.App)).ActiveUIDocument.Document.ParameterBindings;
            map.Insert(def, binding, paramGroup);
        }

        //private bool AddSharedParameterToFamily()
        //{
        //    // check to make sure that the file still exists and is valid
        //    if (File.Exists(_sharedFilePath) && null == _sharedFile)
        //    {
        //        TaskDialog.Show("Error", "SharedPARAM.txt has an invalid format.");
        //        return false;
        //    }

        //    // get the Areas group from shared param file
        //    DefinitionGroup group = _sharedFile.Groups.get_Item(_groupName);
        //    if (null == group) return false;

        //    // get the definition named Program Area
        //    ExternalDefinition def = group.Definitions.get_Item(_paramName) as ExternalDefinition;
        //    if (null == def) return false;

        //    // check whether the parameter already exists in the document
        //    FamilyParameter param = _manager.get_Parameter(def.Name);
        //    if (null != param) return false;   
                
        //    try
        //    {
        //        // add that parameter to the current family document
        //        _manager.AddParameter(def, def.ParameterGroup, true);
        //    }
        //    catch (System.Exception e)
        //    {
        //        TaskDialog.Show("Error", e.Message);
        //        return false;
        //    }

        //    return true;
        //}
    }
}
