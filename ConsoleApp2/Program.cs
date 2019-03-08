using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Office = Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using System.Drawing;
using System.Windows.Forms;

namespace ConsoleApp2
{
    public class Program : System.Windows.Forms.Form
    {
        static void Main(string[] args)
        {
            CzytajDane();
        }
        private static string m_xlFileName = @"C:\Users\mielon\Desktop\Autor1.xlsx";

        private static Excel.Application m_xlApp; // obiekt aplikacji

        private static Excel.Range m_projectRange; // zakres danych do wczytania

        private static Excel.Workbook m_xlWorkbook; // dokument

        private static Excel.Worksheet m_xlWorksheet; // arkusz

        private static System.Object m_xx = System.Type.Missing;

        

        public static void CzytajDane()

        {

            m_xlApp = new Excel.Application();

            m_xlApp.DisplayAlerts = false;


            m_xlWorkbook = m_xlApp.Workbooks.Open(m_xlFileName);


            m_xlWorksheet = (Excel.Worksheet)m_xlWorkbook.Worksheets[1]; // 1 wskazuje na pierwszy arkusz


            string startCell = "A1"; // zakres danych do wczytania

            string endCell = "C3";

            m_projectRange = m_xlWorksheet.get_Range(startCell, endCell);


            Array projectCells = (Array)m_projectRange.Cells.Value2;

            DataGridView dataGridView = new DataGridView();
            
            dataGridView.ColumnCount = m_projectRange.Columns.Count; // dataGridView jest obiektem DataGridView

            dataGridView.RowCount = m_projectRange.Rows.Count;


            for (int i = 0; i < dataGridView.ColumnCount; i++)

            {

                for (int j = 0; j < dataGridView.RowCount; j++)

                {

                    dataGridView[i, j].Value = projectCells.GetValue(j + 1, i + 1);

                }

            }
            

            m_xlApp.Quit();

        }
    }



}
