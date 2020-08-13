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
            foreach(var cell in row.Cells)
            {
                if (cell.StringCellValue.Equals("年份"))
                {

                }
                if (cell.StringCellValue.Equals("风向"))
                {

                }
            }
           
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            // 在WPF中， OpenFileDialog位于Microsoft.Win32名称空间
            Microsoft.Win32.OpenFileDialog dialog = new Microsoft.Win32.OpenFileDialog();

            if (dialog.ShowDialog() == true)
            {
                ReadFile(dialog.FileName);
            }
     
     
        }

    }
}
