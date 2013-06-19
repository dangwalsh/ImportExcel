using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Data;

using Autodesk.Revit.UI;

namespace ImportExcel
{
    class EAFileData
    {
        private string _path;
        private string _name;
        static private DataTable _roomTable;
        static private DataTable _paramTable;
        static private Dictionary<string, string> _mapDict;

        /// <summary>
        /// Data member for file path
        /// </summary>
        public string Path
        {
            get { return _path; }
            set { _path = value; }
        }

        /// <summary>
        /// Data member for table name
        /// </summary>
        public string Name
        {
            get { return _name; }
        }

        /// <summary>
        /// Data member for room DataTable
        /// </summary>
        static public DataTable RoomTable
        {
            get { return _roomTable; }
            set { _roomTable = value; }
        }

        /// <summary>
        /// Data memeber for parameter DataTable
        /// </summary>
        static public DataTable ParamTable
        {
            get { return _paramTable; }
            set { _paramTable = value; }
        }

        /// <summary>
        /// Data member for mapping table columns to Revit parameters
        /// </summary>
        static public Dictionary<string, string> MapDict
        {
            get { return _mapDict; }
            set { _mapDict = value; }
        }
        
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="p"></param>
        public EAFileData(string p)
        {
            _path = p;
            this.Parse();
        }

        /// <summary>
        /// Method that reads the text file and builds datatables for datagrids
        /// </summary>
        public void Parse()
        {
            try
            {
                StreamReader stream = File.OpenText(this.Path);

                string line;
                string[] a;
                char delim = '\t';
                char[] c = { '.', '"', '\'' };

                while (null != (line = stream.ReadLine()))
                {
                    a = line.Split(delim).Select<string, string>(s => s.Trim(c)).ToArray();

                    if (a[0] == "") 
                        continue;

                    if (null == _name)
                    {
                        _name = a[0];
                        continue;
                    }

                    if (null == _roomTable)
                    {
                        _roomTable = new DataTable();
                        _paramTable = new DataTable();
                        _mapDict = new Dictionary<string,string>();

                        DataColumn col = new DataColumn();
                        col.DataType = typeof(string);
                        col.ColumnName = "Param";
                        _paramTable.Columns.Add(col);
                        _paramTable.BeginLoadData();
                     
                        foreach (string column_name in a)
                        {
                            DataColumn column = new DataColumn();
                            column.DataType = typeof(string);
                            column.ColumnName = column_name;
                            _roomTable.Columns.Add(column);

                            string[] da = { column_name };
                            DataRow d = _paramTable.LoadDataRow(da, true);
                        }

                        _paramTable.EndLoadData();
                        _roomTable.BeginLoadData();
                        continue;
                    }                    
                    DataRow dr = _roomTable.LoadDataRow(a, true);
                }
                _roomTable.EndLoadData();
                stream.Close();
            }

            catch (Exception ex)
            {
                TaskDialog.Show("Error", ex.Message);
            }
        }
    }
}
