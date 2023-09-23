using ExportListToExcel.Model;
using Microsoft.Win32;
using OfficeOpenXml.Style;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using System.IO;

namespace ExportListToExcel.ViewModel
{
    public class MainViewModel:BaseViewModel
    {
        private int _TotalData;
        public int TotalData { get=>_TotalData; set { _TotalData = value;OnPropertyChanged(); } }
        private ObservableCollection<ProfileActor> _Profiles;
        public ObservableCollection<ProfileActor> Profiles { get=>_Profiles; set { _Profiles = value;OnPropertyChanged(); } }
        public MainViewModel() 
        { 
            AddDataCommand=new RelayCommand<object>((p) => TotalData <= 0 ? false : true, (p) => { AddData(p); });
            ExportExcelCommand=new RelayCommand<object>((p)=>TotalData<=0?false : true, (p) => { ExportExcel(p); });
        }

        private void ExportExcel(object p)
        {
            string filePath = "";
            // tạo SaveFileDialog để lưu file excel
            SaveFileDialog dialog = new SaveFileDialog();

            // chỉ lọc ra các file có định dạng Excel
            dialog.Filter = "Excel | *.xlsx | Excel 2003 | *.xls";

            // Nếu mở file và chọn nơi lưu file thành công sẽ lưu đường dẫn lại dùng
            if (dialog.ShowDialog() == true)
            {
                filePath = dialog.FileName;
            }

            // nếu đường dẫn null hoặc rỗng thì báo không hợp lệ và return hàm
            if (string.IsNullOrEmpty(filePath))
            {
                MessageBox.Show("Đường dẫn báo cáo không hợp lệ");
                return;
            }

            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (ExcelPackage x = new ExcelPackage())
                {
                    // đặt tên người tạo file
                    x.Workbook.Properties.Author ="Đức";

                    // đặt tiêu đề cho file
                    x.Workbook.Properties.Title = "Add data";

                    //Tạo một sheet để làm việc trên đó
                    x.Workbook.Worksheets.Add("Add data sheet");

                    // lấy sheet vừa add ra để thao tác
                    ExcelWorksheet ws = x.Workbook.Worksheets[0];

                    // đặt tên cho sheet
                    ws.Name = "Add data sheet";
                    // fontsize mặc định cho cả sheet
                    ws.Cells.Style.Font.Size = 11;
                    // font family mặc định cho cả sheet
                    ws.Cells.Style.Font.Name = "Calibri";

                    // Tạo danh sách các column header
                    string[] arrColumnHeader = {
                                                "Họ tên",
                                                "Năm sinh"
                };

                    // lấy ra số lượng cột cần dùng dựa vào số lượng header
                    var countColHeader = arrColumnHeader.Count();

                    // merge các column lại từ column 1 đến số column header
                    // gán giá trị cho cell vừa merge là Thống kê thông tni User Kteam
                    ws.Cells[1, 1].Value = "Add data";
                    ws.Cells[1, 1, 1, countColHeader].Merge = true;
                    // in đậm
                    ws.Cells[1, 1, 1, countColHeader].Style.Font.Bold = true;
                    // căn giữa
                    ws.Cells[1, 1, 1, countColHeader].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    int colIndex = 1;
                    int rowIndex = 2;

                    //tạo các header từ column header đã tạo từ bên trên
                    foreach (var item in arrColumnHeader)
                    {
                        var cell = ws.Cells[rowIndex, colIndex];

                        //set màu thành gray
                        var fill = cell.Style.Fill;
                        fill.PatternType = ExcelFillStyle.Solid;
                        //fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);

                        //căn chỉnh các border
                        //var border = cell.Style.Border;
                        //border.Bottom.Style =
                        //    border.Top.Style =
                        //    border.Left.Style =
                        //    border.Right.Style = ExcelBorderStyle.Thin;

                        //gán giá trị
                        cell.Value = item;

                        colIndex++;
                    }
                    
                    // với mỗi item trong danh sách sẽ ghi trên 1 dòng
                    foreach (var item in Profiles)
                    {
                        // bắt đầu ghi từ cột 1. Excel bắt đầu từ 1 không phải từ 0
                        colIndex = 1;

                        // rowIndex tương ứng từng dòng dữ liệu
                        rowIndex++;

                        //gán giá trị cho từng cell                      
                        ws.Cells[rowIndex, colIndex++].Value = item.Id;

                        // lưu ý phải .ToShortDateString để dữ liệu khi in ra Excel là ngày như ta vẫn thấy.Nếu không sẽ ra tổng số :v
                        ws.Cells[rowIndex, colIndex++].Value = item.Name;

                    }

                    //Lưu file lại
                    Byte[] bin = x.GetAsByteArray();
                    File.WriteAllBytes(filePath, bin);
                }
                MessageBox.Show("Xuất excel thành công!");
            }
            catch (Exception EE)
            {
                MessageBox.Show("Có lỗi khi lưu file!");
            }
        }

        public ICommand AddDataCommand { get; set; }
        public ICommand ExportExcelCommand { get; set; }
        private void AddData(object p)
        {
            StartTask(() => {
                int IdInt = 0;
                for (int i = 0; i < TotalData; i++)
                {
                    if (Profiles == null)
                    {
                        Profiles = new ObservableCollection<ProfileActor>();
                    }
                    Application.Current.Dispatcher.InvokeAsync(new Action(() =>
                    {
                        var Profile = new ProfileActor() { Id = IdInt, Name = "Vương Vũ Tiệp" };
                        Profiles.Add(Profile);
                        foreach (ProfileActor actor in Profiles)
                        {
                            actor.Id = ++IdInt;
                        }
                    }));
                }
            }, null, null);
        }
    }
}
