using ClosedXML.Excel;
using Microsoft.Win32;
using System.Windows;

namespace ExcelConvert
{
    public partial class MainWindow : Window
    {

        public string mainFilePath;
        public string instructionFilePath;


        public MainWindow()
        {
            InitializeComponent();

        }

        public void Input_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog mainFileOpenDialog = new OpenFileDialog()
            {
                Filter = "Excel Files|*.xlsx;*.xls",
                Title = "Select the main Excel file"
            };

            if (mainFileOpenDialog.ShowDialog() == true)
            {
                mainFilePath = mainFileOpenDialog.FileName;
            }
        }




        public void Instruction_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog instructionFileOpenDialog = new OpenFileDialog
            {
                Filter = "Excel Files|*.xlsx;*.xls",
                Title = "Select the instruction Excel file"
            };
            if (instructionFileOpenDialog.ShowDialog() == true)
            {
                instructionFilePath = instructionFileOpenDialog.FileName;
            }
        }


        public void Convert_Click(object sender, RoutedEventArgs e)
        {
            //if (mainFilePath == null || instructionFilePath == null)
            //{
            //    MessageBox.Show("input and Instruction is mandatory!!!");
            //    return;
            //}
            //OpenFileDialog mainFileOpenDialog = new OpenFileDialog
            //{
            //    Filter = "Excel Files|*.xlsx;*.xls",
            //    Title = "Select the main Excel file"
            //};

            //if (mainFileOpenDialog.ShowDialog() == true)
            //{
            //string mainFilePath = mainFileOpenDialog.FileName;

            //OpenFileDialog instructionFileOpenDialog = new OpenFileDialog
            //{
            //    Filter = "Excel Files|*.xlsx;*.xls",
            //    Title = "Select the instruction Excel file"
            //};

            //if (instructionFileOpenDialog.ShowDialog() == true)
            //{
            //string instructionFilePath = instructionFileOpenDialog.FileName;

            //using (var mainWorkbook = new XLWorkbook(mainFilePath))
            //{
            //    var mainWorksheet = mainWorkbook.Worksheet(1);
            //}

            using (var instructionWorkbook = new XLWorkbook(instructionFilePath))
            using (var mainWorkbook = new XLWorkbook(mainFilePath))
            using (var newWorkbook = new XLWorkbook())
            {

                var instructionWorksheet = instructionWorkbook.Worksheet(1);
                var newWorksheet = newWorkbook.Worksheets.Add("Sheet1");

                foreach (IXLRow instructionRow in instructionWorksheet.RowsUsed())
                {
                    if (instructionRow.RowNumber() == 1) continue;
                    if (instructionRow.Cell(1).IsEmpty()) break;
                    //input
                    var sheetno = (int)instructionRow.Cell(2).GetDouble();
                    var firstRow = (int)instructionRow.Cell(3).GetDouble();
                    var firstColumn = (int)instructionRow.Cell(4).GetDouble();
                    var nextRow = (int)instructionRow.Cell(5).GetDouble();
                    var nextColumn = (int)instructionRow.Cell(6).GetDouble();
                    var lastRow = (int)instructionRow.Cell(7).GetDouble();
                    var lastColumn = (int)instructionRow.Cell(8).GetDouble();

                    //output
                    var firstRowO = (int)instructionRow.Cell(9).GetDouble();
                    var firstColumnO = (int)instructionRow.Cell(10).GetDouble();
                    var nextRowO = (int)instructionRow.Cell(11).GetDouble();
                    var nextColumnO = (int)instructionRow.Cell(12).GetDouble();

                    var mainWorksheet = mainWorkbook.Worksheet(sheetno);


                    var ir = firstRow;
                    var ic= firstColumn;
                    var or = firstRowO;
                    var oc = firstColumnO;
                    while(true)
                    {
                        if (ir > lastRow) break;
                        if (ic > lastColumn) break;

                        newWorksheet.Cell(or, oc).Value = mainWorksheet.Cell(ir, ic).Value;

                        ir += nextRow;
                        ic += nextColumn;

                        or += nextRowO;
                        oc += nextColumnO;
                    }


                }

                // save to new file
                string newFilePath = @"C:\Users\user1\Documents\Excel Convert\New_Excel.xlsx";
                newWorkbook.SaveAs(newFilePath);
            }


            MessageBox.Show("Conversion successful! New file saved.");
        }
    }

    //}
    //}  
}




