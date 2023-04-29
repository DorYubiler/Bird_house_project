using System;
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
using Microsoft.Office.Interop.Excel;


namespace BirdsHouse_Project
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        public MainWindow()
        {
            InitializeComponent();
            
        }
        string userName, userPass; 

        private void Submit_btn_Click(object sender, RoutedEventArgs e)
        {
            userName = uName.Text;
            userPass = uPass.Password;


            string filePath = "";
            // ...

            // Create an instance of Excel application
            var excelApp = new System.Windows.Application();

            // Open the Excel workbook
            Workbook workbook = excelApp.Workbooks.Open(filePath);

            // Select the first worksheet in the workbook
            Worksheet worksheet = (Worksheet)workbook.Worksheets[1];

            // Write "hi" into cell A1
            worksheet.Cells[1, 1] = "hi";

            // Save the changes to the Excel file
            workbook.Save();

            // Close the workbook and Excel application
            workbook.Close();
            excelApp.Quit();


        }
    }
}
