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
using OfficeOpenXml;
using System.IO;
using LINQLESSON2.Model;
using System.Data;

namespace LINQLESSON2
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }
        #region exmapl
        //.Concat(массив) и .Union(массив);
        static void Example1()
        {
            int[] seq1 = { 1, 2, 3 };
            int[] seq2 = { 4, 5, 6 };

            //возвращает элементы 1 и 2 последовательности
            IEnumerable<int> concat = seq1.Concat(seq2);
            //тоже самое, но убирает повторение 
            IEnumerable<int> union = seq1.Union(seq2);
        }
        //.Intersect(массив) и .Except(массив);
        static void Example2()
        {
            int[] seq1 = { 1, 2, 3 };
            int[] seq2 = { 4, 5, 6 };

            //Общие для двух последовательностей
            IEnumerable<int> intersect = seq1.Intersect(seq2);

            //из 1 первой, которых нет во второй
            IEnumerable<int> except = seq1.Except(seq2);


            //NOT in, NOT EXITSTS - SQL
        }
        //массив.Cast<int>() и массив.OfType<int>();
        static void Example3()
        {
            int[] seq1 = { 1, 2, 3 };
            string[] seq2 = { "4", "5", "6" };

            // делает ошибки
            IEnumerable<int> seq1Int = seq2.Cast<int>();

            //пропускает ошибки
            IEnumerable<int> seq1IntNotErr = seq2.OfType<int>();
            //локальная
            var asEn = seq1IntNotErr.AsEnumerable();
            // не локальная
            var asEnQ = seq1IntNotErr.AsQueryable();

        }
        #endregion

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            ExcelPackage package = new ExcelPackage();
            string pathForSaving = "\\Template";
            FileInfo template = new FileInfo("myTemplate.xlsx");

            using (package = new ExcelPackage(template, true))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[1];

                Helper helper = new Helper();
                DataSet data = helper.GetData();
                int row = 3;

                List<Area> areas = new List<Area>();
                List<Pavilion> pavilions = new List<Pavilion>();
                foreach (DataRow dataRow in data.Tables["Area"].Rows)
                {
                    Area area = new Area();
                    area.AreaId = Int32.Parse(dataRow["AreaID"].ToString());
                    area.Name = dataRow["Name"].ToString();
                    area.FullName = dataRow["FullName"].ToString();
                    area.PavilionId = Int32.Parse(dataRow["PavilionId"].ToString());
                    areas.Add(area);
                }

                foreach (DataRow dataRow in data.Tables["dic_Pavilion"].Rows)
                {
                    Pavilion pav = new Pavilion();
                    pav.PavilionId = Int32.Parse(dataRow["PavilionID"].ToString());
                    pav.Name = dataRow["Name"].ToString();
                   
                    pavilions.Add(pav);
                }


                var query = areas.Join(pavilions,
                    a => a.PavilionId,
                    p => p.PavilionId,
                    (a, p) => new
                    {
                        a.Name,
                        a.FullName,
                        PavilionName = p.Name

                    });

                foreach(var dataRow in query)
                {
                    worksheet.Cells[row, 2].Value = dataRow.Name;
                    worksheet.Cells[row, 3].Value = dataRow.FullName;
                    worksheet.Cells[row, 7].Value = dataRow.PavilionName;
                    row++;
                }

                //foreach(DataRow Datarow in data.Tables["Area"].Rows)
                //{
                //    worksheet.Cells[row, 2].Value = Datarow["Name"];
                //    worksheet.Cells[row, 3].Value = Datarow["FullName"];
                //    //worksheet.Cells[row, 4].Value = Datarow["PavilionName"];
                //    row++;
                //}
                package.SaveAs(new System.IO.FileInfo("demoOut.xlsx"));
            }
        }

        public class Area
        {
            public int AreaId { get; set; }
            public string Name { get; set; }
            public string FullName { get; set; }
            public int PavilionId { get; set; }
        }

        public  class Pavilion
        {
            public int PavilionId { get; set; }
            public string Name { get; set; }

            
        }
    }
}
