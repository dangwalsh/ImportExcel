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
        private EAFileData _data;
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
                    _data = new EAFileData(dlg.FileName);
                    this.groupBox.Header = _data.Name;
                    this.programGrid.ItemsSource = EAFileData.RoomTable.AsDataView();
                    this.paramGrid.ItemsSource = EAFileData.ParamTable.AsDataView();
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
            EAFactory factory = new EAFactory();
            factory.CreateRooms();
            Autodesk.Revit.UI.TaskDialog.Show("Success", "Successfully imported room data!");
            this.Close();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

            // TODO: Must add ability to remove Dictionary entry if user deselects parameter mapping

            ComboBox temp = (ComboBox)sender;
            string p = temp.SelectedItem.ToString();
            DataRowView rowView = paramGrid.SelectedItem as DataRowView;
            if (rowView != null)
            {
                DataRow row = rowView.Row;
                string s = row["Param"].ToString();
                EAFileData.MapDict.Add(p, row.ItemArray[0].ToString());
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
            Transaction t = new Transaction(EADocumentData.Doc, "Make Temp Room");
            if (t.Start() == TransactionStatus.Started)
            {
                EAFactory.CreateRoom(EAFactory.CreatePhase());
                GetRooms();
            }
            t.RollBack();

            foreach (string name in EASharedParamData.SharedParamNames)
            {
                this.Add(name);
            }

            this.Sort();
        }

        /// <summary>
        /// Collect room instances from Revit Document
        /// </summary>
        private void GetRooms()
        {
            FilteredElementCollector collector = new FilteredElementCollector(EADocumentData.Doc);
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
