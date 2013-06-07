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
    class EAData
    {
        private string _path;
        private string _name;
        static private DataTable _table;
        static private DataTable _paramTable;
        static private Dictionary<string, string> _paramDict;

        /// <summary>
        /// 
        /// </summary>
        public string Path
        {
            get { return _path; }
            set { _path = value; }
        }

        /// <summary>
        /// 
        /// </summary>
        public string Name
        {
            get { return _name; }
        }

        /// <summary>
        /// 
        /// </summary>
        static public DataTable Table
        {
            get { return _table; }
            set { _table = value; }
        }

        /// <summary>
        /// 
        /// </summary>
        static public DataTable ParamTable
        {
            get { return _paramTable; }
            set { _paramTable = value; }
        }

        static public Dictionary<string, string> ParamDict
        {
            get { return _paramDict; }
            set { _paramDict = value; }
        }
        
        /// <summary>
        /// 
        /// </summary>
        /// <param name="p"></param>
        public EAData(string p)
        {
            _path = p;
            this.Parse();
        }

        /// <summary>
        /// 
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

                    if (null == _table)
                    {
                        _table = new DataTable();
                        _paramTable = new DataTable();
                        _paramDict = new Dictionary<string,string>();

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
                            _table.Columns.Add(column);

                            string[] da = { column_name };
                            DataRow d = _paramTable.LoadDataRow(da, true);
                        }
                        _paramTable.EndLoadData();
                        _table.BeginLoadData();
                        continue;
                    }                    

                    DataRow dr = _table.LoadDataRow(a, true);
                }
                _table.EndLoadData();
            }

            catch (Exception ex)
            {
                TaskDialog.Show("Error", ex.Message);
            }
        }
    }
}
