using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;

using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.ApplicationServices;

namespace ImportExcel
{
    class EAFactory
    {
        /// <summary>
        /// Method that creates new rooms
        /// </summary>
        /// <param name="p"></param>
        /// <returns></returns>
        static public Room CreateRoom(Phase p)
        {
            Room room = EADocumentData.Doc.Create.NewRoom(p);
            return room;
        }

        /// <summary>
        /// Method that creates a phase object equal to the last phase in the active document
        /// </summary>
        /// <returns></returns>
        static public Phase CreatePhase()
        {
            PhaseArray phases = EADocumentData.Doc.Phases;
            int n = phases.Size;
            return phases.get_Item(n - 1);          
        }

        /// <summary>
        /// Method that creates transaction calls room creation and commits to DB
        /// </summary>
        public void CreateRooms()
        {
            try
            {
                Phase phase = CreatePhase();
                Transaction t = new Transaction(EADocumentData.Doc, "Create Rooms");
                if (t.Start() == TransactionStatus.Started)
                {
                    foreach (DataRow row in EAFileData.RoomTable.Rows)
                    {
                        if (row.RowState != DataRowState.Deleted)
                        {
                            short s = 0;
                            if (Int16.TryParse(row[EAFileData.MapDict["Count"]].ToString(), out s))
                            {
                                //EAFileData.MapDict.Remove("Count");
                                for (int i = 0; i < s; ++i)
                                {
                                    Room room = CreateRoom(phase);
                                    room.Name = row[EAFileData.MapDict["Name"]].ToString();
                                    //EAFileData.MapDict.Remove("Name");
                                    ParamFactory(room as Element, row);
                                }
                            }
                            //else
                            //{
                            //    //throw new CountFailedException("Failed to access room count");
                            //}
                        }
                    }
                    t.Commit();
                }
            }
            //catch (CountFailedException ex)
            //{
            //    TaskDialog.Show("Count Failure", ex.Message);
            //}
            catch (Exception ex)
            {
                TaskDialog.Show("Error", ex.Message);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="elem"></param>
        void ParamFactory(Element elem, DataRow row)
        {
            foreach (KeyValuePair<string, string> usrParam in EAFileData.MapDict)
            {
                if (usrParam.Key == "Count" || usrParam.Key == "Name")
                    continue;

                bool match = false;
                ParameterSetIterator parSetItor = elem.Parameters.ForwardIterator();
                while (parSetItor.MoveNext())
                {
                    Parameter elemParam = parSetItor.Current as Parameter;

                    if (usrParam.Key == elemParam.Definition.Name)  // if it already exists we need to assign it
                    {
                        match = true;
                        // need to test storage type  NOTE: may not need to do this...
                        switch (elemParam.StorageType)
                        {
                            case StorageType.ElementId:
                                // don't ever want to assign elementId
                                break;
                            case StorageType.Double:
                                elemParam.SetValueString(row[usrParam.Value].ToString());
                                break;
                            case StorageType.String:
                                elemParam.Set(row[usrParam.Value].ToString());
                                break;
                            case StorageType.Integer:
                                if (elemParam.Definition.ParameterType != ParameterType.YesNo)
                                    elemParam.SetValueString(row[usrParam.Value].ToString());
                                break;
                            default:
                                break;
                        }
                    }
                }
                    
                if (!match)  // if it does NOT exist we need to create it
                {
                    Category roomCat = EADocumentData.Doc.Settings.Categories.get_Item(BuiltInCategory.OST_Rooms);
                    CategorySet catSet = new CategorySet();
                    catSet.Insert(roomCat);
                    EASharedParamData.RawCreateProjectParameterFromExistingSharedParameter(usrParam.Key, 
                                                                                           catSet, 
                                                                                           BuiltInParameterGroup.PG_IDENTITY_DATA, 
                                                                                           true
                                                                                           );
                    // now that we've made it, we need to set it
                    ParameterSetIterator parSetRevItor = elem.Parameters.ReverseIterator();
                    while (parSetRevItor.MoveNext())
                    {
                        Parameter elemParam = parSetRevItor.Current as Parameter;
                        if (usrParam.Key == elemParam.Definition.Name)  // if it already exists we need to assign it
                        {
                            switch (elemParam.StorageType)
                            {
                                case StorageType.ElementId:
                                    // don't ever want to assign elementId
                                    break;
                                case StorageType.Double:
                                    elemParam.SetValueString(row[usrParam.Value].ToString());
                                    break;
                                case StorageType.String:
                                    elemParam.Set(row[usrParam.Value].ToString());
                                    break;
                                case StorageType.Integer:
                                    if (elemParam.Definition.ParameterType != ParameterType.YesNo)
                                        elemParam.SetValueString(row[usrParam.Value].ToString());
                                    break;
                                default:
                                    break;
                            }
                            break;
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Destructor
        /// </summary>
        ~EAFactory()
        {
            EAFileData.RoomTable = null;
        }
    }
}
