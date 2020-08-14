using NPOI;
using NPOI.SS.UserModel;
using System;
using System.IO;
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

namespace jianjianwang
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        Dictionary<double, List<string>> dataMap = new Dictionary<double, List<string>>();
        Dictionary<double, List<Wind>> dataSortMap = new Dictionary<double, List<Wind>>();
        private class Wind

        {
            public string weight;
            public int count;

            public Wind(string weight, int count)
            {
                this.weight = weight;
                this.count = count;
            }
        }
        public MainWindow()
        {
            InitializeComponent();
        }
        /// <summary>
        /// 读取excel
        /// 读取数据 
        /// </summary>
        /// <param name="filePath">文件路径</param>
        private void ReadFile(String filePath)
        {
            Console.WriteLine(filePath);
            IWorkbook workbook = WorkbookFactory.Create(filePath);
            ISheet sheet = workbook.GetSheetAt(1);//获取第2个工作薄
            IRow row = sheet.GetRow(0);//获取第一行

            //定位 年份列、风向列
            int yearColumnIndex = 0;
            int windColumnIndex = 0;
            for (int i = 0; i < row.Cells.Count; i++)
            {
                if ("年份".Equals(row.Cells[i].StringCellValue))
                {
                    yearColumnIndex = i;
                }
                if ("风向".Equals(row.Cells[i].StringCellValue))
                {
                    windColumnIndex = i;
                }
            }

            int rowCount = sheet.PhysicalNumberOfRows;
            // 以行遍历,从第二行开始
            for (int j = 1; j < rowCount; j++)
            {
                row = sheet.GetRow(j);
                double year = row.Cells[yearColumnIndex].NumericCellValue;

                List<string> listDataByYear;
                if (!dataMap.ContainsKey(year))
                {
                    listDataByYear = new List<string>();
                    dataMap.Add(year, listDataByYear);
                }
                else
                {
                    listDataByYear = dataMap[year];
                }

                listDataByYear.Add(row.Cells[windColumnIndex].StringCellValue);


            }
        }
        /// <summary>
        /// 整理数据
        /// </summary>
        private void SortData()
        {
            foreach (var data in dataMap)
            {
                List<string> windList = data.Value;
                var groupList = windList.GroupBy(x => x).Select(group => new Wind(group.Key, group.Count())).ToList();
                Console.WriteLine(groupList);
                dataSortMap.Add(data.Key, groupList);
            }

        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            // 在WPF中， OpenFileDialog位于Microsoft.Win32名称空间
            Microsoft.Win32.OpenFileDialog dialog = new Microsoft.Win32.OpenFileDialog();

            if (dialog.ShowDialog() == true)
            {
                Go(dialog.FileName);

            }


        }

        private void Window_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string filePath = ((Array)e.Data.GetData(DataFormats.FileDrop)).GetValue(0).ToString();
                Go(filePath);
            }
        }

        private void Go(string filePath)
        {
            ReadFile(filePath);
            SortData();
        }
    }
}
