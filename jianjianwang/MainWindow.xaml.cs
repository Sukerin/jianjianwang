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
using System.Collections.ObjectModel;
using Newtonsoft.Json;

namespace jianjianwang
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        Dictionary<double, List<string>> dataMap = new Dictionary<double, List<string>>();
     
        Dictionary<double, List<Wind>> dataSortMap = new Dictionary<double, List<Wind>>();

        Dictionary<string, List<Wind>> finalSortMap = new Dictionary<string, List<Wind>>();

        List<string> defaultDirectionList = new List<string>{ "N","NNE","NE","ENE","E","ESE","SE","SSE","S","SSW","SW","WSW","W","WNW","NW","NNW" };


        
        private class Wind

        {
            public double year { get; set; }
            public string direction { get; set; }
            public int count { get; set; }

            public Wind(string direction, int count, double year)
            {
                this.direction = direction;
                this.count = count;
            }
            public Wind(string direction, int count)
            {
                this.direction = direction;
                this.count = count;
            }
        }
        public MainWindow()
        {
            Resources["Winds"] = dataSortMap;
            InitializeComponent();
            webBrowser.Navigate(new Uri(Directory.GetCurrentDirectory() + "\\charts1.html"));
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
                string direction= row.Cells[windColumnIndex].StringCellValue;
                List<Wind> windList;
                if (defaultDirectionList.Contains(direction))
                {
                    if (!finalSortMap.ContainsKey(direction))
                    {
                        windList = new List<Wind>();
                        finalSortMap.Add(direction, windList);
                    }
                    else
                    {
                        windList = finalSortMap[direction];
                    }
                    Wind wind = new Wind(direction, 0, year);
                    windList.Add(wind);
                }

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


            foreach (var data in finalSortMap)
            {

                List<Wind> windList = data.Value;

                var groupDic = windList.GroupBy(x => x.year).Select(group => new Wind(data.Key,group.Count(),group.Key)).ToList();

                foreach (var dd in defaultDirectionList)
                {
                    if (!finalSortMap.ContainsKey(dd))
                    {
                        finalSortMap.Add(dd, new List<Wind>());
                    }

                }
            }


            foreach (var data in dataMap)
            {

                List<string> windList = data.Value;

                var groupDic = windList.GroupBy(x => x).ToDictionary(group => group.Key, group => new Wind(group.Key, group.Count()));
                groupDic.Remove("/");
                groupDic.Remove("C");
                foreach (var dd in defaultDirectionList)
                {
                    if (!groupDic.ContainsKey(dd))
                    {
                        groupDic.Add(dd, new Wind(dd, 0));
                    }

                }

                groupDic=groupDic.OrderBy(p => p.Key).ToDictionary(p => p.Key, o => o.Value);
                var groupList = new List<Wind>();
                foreach(var item in groupDic)
                {
                    groupList.Add(item.Value);
                }
                dataSortMap.Add(data.Key, groupList);

            }
            

      

      

        }
        /// <summary>
        /// 渲染
        /// </summary>
        void Render()
        {
            
            //渲染表格
            int count = 0;
            foreach (List<Wind> lst in dataSortMap.Values)
            {
                if (lst.Count > count)
                {

                    for (int i = count; i < lst.Count; i++)
                    {
                        DataGridTextColumn column = new DataGridTextColumn();
                        column.Header = lst[i].direction;
                        column.Binding = new Binding(string.Format("Value[{0}].count", i));
                        dataGrid.Columns.Add(column);

                    }
                    count = lst.Count;
                }
            }

            //渲染玫瑰图
            var yearJson =JsonConvert.SerializeObject(dataSortMap.Keys);
            List<List<int>> dataCountlist = new List<List<int>>();
            List<int> countlist;
            int maxCount=0;
            foreach (var value in dataSortMap.Values)
            {
                countlist = value.Select(x => x.count).ToList();
                if(maxCount< countlist.Max())
                {
                    maxCount = countlist.Max();
                }
                
                dataCountlist.Add(countlist);
            }
            //计算几位数
            int n = (int)Math.Ceiling(Math.Log10(maxCount));
            double v1 = Math.Floor(maxCount / Math.Pow(10, n - 2));
            double maxAxis = (v1 + 1) * Math.Pow(10, n - 2);

            var result=JsonConvert.SerializeObject(dataSortMap);
            
            webBrowser.InvokeScript("initOption", JsonConvert.SerializeObject(dataCountlist), maxAxis);
            
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
            dataMap.Clear();
            dataSortMap.Clear();
            ReadFile(filePath);
            SortData();
            Render();

        }
    }
}
