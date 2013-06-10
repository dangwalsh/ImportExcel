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
    class EARoomFactory
    {
        /// <summary>
        /// Method that creates new rooms
        /// </summary>
        /// <param name="p"></param>
        /// <returns></returns>
        static public Room CreateRoom(Phase p)
        {
            Room room = EADocument.Doc.Create.NewRoom(p);
            return room;
        }

        /// <summary>
        /// Method that creates a phase object equal to the last phase in the active document
        /// </summary>
        /// <returns></returns>
        static public Phase CreatePhase()
        {
            PhaseArray phases = EADocument.Doc.Phases;
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
                Transaction t = new Transaction(EADocument.Doc, "Create Rooms");
                if (t.Start() == TransactionStatus.Started)
                {
                    foreach (DataRow row in EAData.Table.Rows)
                    {
                        if (row.RowState != DataRowState.Deleted)
                        {
                            short s = 0;
                            if (Int16.TryParse(row[EAData.ParamDict["Count"]].ToString(), out s))
                            {
                                for (int i = 0; i < s; ++i)
                                {
                                    Room room = CreateRoom(phase);
                                    room.Name = row[EAData.ParamDict["Name"]].ToString();
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
            catch (CountFailedException ex)
            {
                TaskDialog.Show("Count Failure", ex.Message);
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", ex.Message);
            }
        }

        /// <summary>
        /// Destructor
        /// </summary>
        ~EARoomFactory()
        {
            EAData.Table = null;
        }
    }
}
