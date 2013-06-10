using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Data;

using Autodesk.Revit.DB;

namespace ImportExcel
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private EAData _data;
        public MainWindow()
        {
            InitializeComponent();
        }

        /// <summary>
        /// File open event handler
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnFile_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();

                dlg.Title = "Select Program Data File";
                dlg.CheckFileExists = true;
                dlg.CheckPathExists = true;
                dlg.Filter = ".txt Files (*.txt)|*.txt";

                if (dlg.ShowDialog() == true)
                {
                    _data = new EAData(dlg.FileName);
                    this.groupBox.Header = _data.Name;
                    this.programGrid.ItemsSource = EAData.Table.AsDataView();
                    this.paramGrid.ItemsSource = EAData.ParamTable.AsDataView();
                }
                else
                {
                    throw new FileReadFailedException("File read failed!");
                }
            }
            catch (FileReadFailedException ex)
            {
                Console.WriteLine(ex.Message);
                Console.Read();
            }
        }

        /// <summary>
        /// Import button event handler
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnRoom_Click(object sender, RoutedEventArgs e)
        {
            EARoomFactory factory = new EARoomFactory();
            factory.CreateRooms();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ComboBox temp = (ComboBox)sender;
            string p = temp.SelectedItem.ToString();
            DataRowView rowView = paramGrid.SelectedItem as DataRowView;
            if (rowView != null)
            {
                DataRow row = rowView.Row;
                string s = row["Param"].ToString();
                EAData.ParamDict.Add(p, row.ItemArray[0].ToString());
            }
        }
    }

    public class ParamList : List<string>
    {
        /// <summary>
        /// Constructor
        /// </summary>
        public ParamList()
        {
            this.Add("Count");
            GetRooms();
        }

        /// <summary>
        /// Collect room instances from Revit Document
        /// </summary>
        private void GetRooms()
        {
            FilteredElementCollector collector = new FilteredElementCollector(EADocument.Doc);
            List<Element> rooms = new List<Element>();
            rooms = collector.OfCategory(BuiltInCategory.OST_Rooms).ToList();

            foreach (Element e in rooms)
            {
                ParamItor(e);
                break;
            }
        }

        /// <summary>
        /// Iterate parameters and add to List
        /// </summary>
        /// <param name="elem"></param>
        private void ParamItor(Element elem)
        {
            foreach (Parameter param in elem.Parameters)
            {
                this.Add(param.Definition.Name);
            }
        }
    }
}
