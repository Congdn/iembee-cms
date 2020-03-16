using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.IO;
using System.IO.Compression;
using System.Web;
using System.Linq;
using System.Web.Mvc;
using System.Web.UI.WebControls;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Threading.Tasks;

namespace iembee.Controllers
{
    [Authorize]
    public class HomeController : Controller
    {
        [OutputCache(Duration = 10, VaryByParam = "*")]
        public ActionResult Index()
        {
            return View();
        }
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<ActionResult> Index(Data model)
        {
            if (ModelState.IsValid)
            {
                var filePath = "";
                //Save upload file
                var filename = Path.GetFileName(model.fileInput.FileName);
                filePath = Path.Combine(Server.MapPath("~/App_Data/uploads"), filename);
                model.fileInput.SaveAs(filePath);

                //Read data
                var datas = new List<Dictionary<string, string>>();

                Excel.Application app = new Excel.Application();
                Excel.Workbook wb = app.Workbooks.Open(filePath);
                Excel.Worksheet ws = wb.Sheets[1];
                Excel.Range range = ws.UsedRange;

                var rows = range.Rows.Count;
                var cols = range.Columns.Count;
                for (int i = 2; i <= rows; i++)
                {
                    var rowData = new Dictionary<string, string>();
                    rowData.Add("Id", range.Cells[i, 1].Value.ToString());
                    rowData.Add("TenMatHang", range.Cells[i, 2].Value.ToString());
                    rowData.Add("DonViTinh", range.Cells[i, 3].Value.ToString());
                    rowData.Add("GiaMua", range.Cells[i, 4].Value.ToString());
                    rowData.Add("GiaBan", range.Cells[i, 5].Value.ToString());

                    datas.Add(rowData);
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
                Marshal.ReleaseComObject(ws);
                wb.Close(false);
                Marshal.FinalReleaseComObject(wb);
                app.Quit();
                Marshal.FinalReleaseComObject(app);
                //

                var savePath = Path.Combine(Server.MapPath("~/App_Data/file_of_mouth_" + DateTime.Now.Month), model.tenkh + "_" + DateTime.Now.ToString("dd-MM-yy"));
                if (!Directory.Exists(savePath))
                {
                    Directory.CreateDirectory(savePath);
                }
                else
                {
                    Directory.Delete(savePath, true);
                    Directory.CreateDirectory(savePath);
                }

                //Export data
                //await Task.WhenAll(
                //    Task.Run(() => ExportFile(model.hangtoida, model.hangtoithieu, 12, model.tongnhap3, model.tongxuat3, datas, savePath, model.tenkh, model.diachi, model.dienthoai)),
                //    Task.Run(() => ExportFile(model.hangtoida, model.hangtoithieu, 1, model.tongnhap3, model.tongxuat3, datas, savePath, model.tenkh, model.diachi, model.dienthoai)),
                //    Task.Run(() => ExportFile(model.hangtoida, model.hangtoithieu, 2, model.tongnhap3, model.tongxuat3, datas, savePath, model.tenkh, model.diachi, model.dienthoai))
                //);
                //await Task.Run(() =>
                //{
                //    ExportFile(model.hangtoida, model.hangtoithieu, 12, model.tongnhap3, model.tongxuat3, datas, savePath, model.tenkh, model.diachi, model.dienthoai);
                //    ExportFile(model.hangtoida, model.hangtoithieu, 1, model.tongnhap3, model.tongxuat3, datas, savePath, model.tenkh, model.diachi, model.dienthoai);
                //    ExportFile(model.hangtoida, model.hangtoithieu, 2, model.tongnhap3, model.tongxuat3, datas, savePath, model.tenkh, model.diachi, model.dienthoai);
                //});
                Parallel.Invoke(
                    () => ExportFile(model.hangtoida, model.hangtoithieu, 12, model.tongnhap3, model.tongxuat3, datas, savePath, model.tenkh, model.diachi, model.dienthoai),
                    () => ExportFile(model.hangtoida, model.hangtoithieu, 12, model.tongnhap3, model.tongxuat3, datas, savePath, model.tenkh, model.diachi, model.dienthoai),
                    () => ExportFile(model.hangtoida, model.hangtoithieu, 12, model.tongnhap3, model.tongxuat3, datas, savePath, model.tenkh, model.diachi, model.dienthoai)
                );
                //Return file rar
                var zipPath = Path.Combine(Server.MapPath("~/App_Data/file_of_mouth_" + DateTime.Now.Month), model.tenkh + DateTime.Now.Month + ".zip");
                if (System.IO.File.Exists(zipPath))
                {
                    System.IO.File.Delete(zipPath);
                }
                ZipFile.CreateFromDirectory(savePath, zipPath);

                ModelState.AddModelError("", "Hoàn thành");
                return new FilePathResult(zipPath, "application/zip");
            }

            return View(model);
        }

        [HttpGet]
        public ActionResult GetFile()
        {
            var filePath = Directory.EnumerateFiles(Server.MapPath("~/App_Data/file_of_mouth_" + DateTime.Now.Month)).Where(file => file.ToLower().EndsWith(".zip")).FirstOrDefault();
            //Directory.GetFiles(Server.MapPath("~/App_Data/file_of_mouth_" + DateTime.Now.Month), SearchOption.TopDirectoryOnly);
            if (!System.IO.File.Exists(filePath))
            {
                var model = new Data();
                ModelState.AddModelError("", "Chưa export file cho tháng này");
                return View("Index", model);
            }
            return new FilePathResult(filePath, "application/zip");
        }

        public void ExportFile(double SoLuongToiDa, double SoLuongToiThieu, int Month, decimal BuyTotal, decimal SaleTotal, List<Dictionary<string, string>> lswRes, string Path, string companyName = "", string Address = "", string Phone = "")
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook wb = excelApp.Workbooks.Add(Type.Missing);
            Excel._Worksheet ws = null,
                ws2 = null;
            excelApp.Windows.Application.ActiveWindow.DisplayGridlines = true;
            var cols = lswRes[0].Count;
            var rows = lswRes.Count;


            try
            {
                #region WorkSheet Bán hàng
                //header
                ws = wb.Worksheets[1];
                ws.Name = "Bán hàng";

                ws.Range[ws.Cells[1, 1], ws.Cells[1, cols + 2]].Merge();
                ws.Range[ws.Cells[2, 1], ws.Cells[2, cols + 2]].Merge();
                ws.Range[ws.Cells[3, 1], ws.Cells[3, cols + 2]].Merge();
                ws.Range[ws.Cells[4, 1], ws.Cells[4, cols + 2]].Merge();
                ws.Range[ws.Cells[5, 1], ws.Cells[5, cols + 2]].Merge();
                //Tên công ty
                ws.Cells[1, 1].Value = companyName.ToUpper();
                ws.Cells[1, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                ws.Cells[1, 1].Font.Size = 20;
                //Địa chỉ
                ws.Cells[2, 1].Value = Address;
                ws.Cells[2, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                ws.Cells[2, 1].Font.Size = 14;
                //Điện thoại
                ws.Cells[3, 1].Value = "'" + Phone;
                ws.Cells[3, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                ws.Rows[3].Style.NumberFormat = "General";
                ws.Cells[3, 1].Font.Size = 12;
                //Tiêu đề
                if (Month == 12)
                    ws.Cells[4, 1].Value = "Bảng tổng hợp doanh số bán hàng tháng " + Month + "/" + DateTime.Now.AddYears(-1).Year;
                else
                    ws.Cells[4, 1].Value = "Bảng tổng hợp doanh số bán hàng tháng " + Month + "/" + DateTime.Now.Year;
                ws.Cells[4, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                ws.Cells[4, 1].Font.Size = 12;
                //Kẻ
                ws.Cells[6, 1].Borders.Weight = Excel.XlBorderWeight.xlThin;
                //title
                for (int i = 1; i <= cols; i++)
                {
                    if (i != 4) ws.Cells[6, i + 1] = lswRes[0].Keys.ToList().ElementAt(i - 1);
                    if (i == 4) ws.Cells[6, i + 1] = "SL";
                    ws.Cells[6, i + 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    ws.Cells[6, i + 1].Font.Bold = true;
                    ws.Cells[6, i + 1].Borders.Weight = Excel.XlBorderWeight.xlThin;
                }
                ws.Cells[6, cols + 2] = "Doanh thu";
                ws.Cells[6, cols + 2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                ws.Cells[6, cols + 2].Font.Bold = true;
                ws.Cells[6, cols + 2].Borders.Weight = Excel.XlBorderWeight.xlThin;
                //
                //Item
                Random random = new Random();
                Random OrderRandom = new Random();
                DateTime date;
                date = new DateTime(DateTime.Now.Year, Month, 1);
                decimal excelSaleTotal = 0;
                int autoNum = 1;
                //Vẽ bảng theo ngày
                while (date.Month == Month)
                {
                    Excel.Range rangeAdd = ws.UsedRange;
                    var rowsAdd = rangeAdd.Rows.Count;
                    //Thêm date
                    ws.Cells[rowsAdd + 1, 1] = "'" + date.ToString("dd/MM/yyyy");
                    ws.Cells[rowsAdd + 1, 1].Borders.Weight = Excel.XlBorderWeight.xlThin;
                    date = date.AddDays(1);
                    //
                    var orderQuatity = OrderRandom.Next(Convert.ToInt32(SoLuongToiThieu), Convert.ToInt32(SoLuongToiDa));
                    //Random số lượng hàng mỗi ngày
                    for (int i = 1; i <= orderQuatity; i++)
                    {
                        var rowRandom = random.Next(2, lswRes.Count - 1);
                        var item = lswRes.ElementAt(rowRandom);
                        ws.Cells[rowsAdd + i, 2] = autoNum;
                        autoNum++;
                        ws.Cells[rowsAdd + i, 2].Borders.Weight = Excel.XlBorderWeight.xlThin;
                        //vẽ từng hàng
                        for (int j = 2; j <= cols; j++)
                        {
                            ws.Cells[rowsAdd + i, j + 1] = item[item.Keys.ToList().ElementAt(j - 1)];
                            ws.Cells[rowsAdd + i, j + 1].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (j == 5)
                            {
                                excelSaleTotal += Convert.ToDecimal(item[item.Keys.ToList().ElementAt(j - 1)]);
                            }
                        }
                    }
                }
                //Số lượng trung bình từ tổng nhập và tổng giá bán
                double quantity = Math.Round(Convert.ToDouble(SaleTotal / excelSaleTotal));
                double soluong_thua = 0;
                //clear excelSaleTotal
                excelSaleTotal = 0;
                Excel.Range range = ws.UsedRange;
                var row2s = range.Rows.Count;
                //Add cột số lượng
                for (int i = 7; i <= row2s; i++)
                {
                    double soLuong = 0;
                    if (quantity <= 1)
                        soLuong = quantity + random.Next(0, 1);
                    else if (quantity <= 2)
                        soLuong = quantity + random.Next(-1, 2);
                    else
                        soLuong = quantity + random.Next(-2, 3);
                    if (soLuong <= 0)
                    {
                        soluong_thua = Math.Abs(soLuong);
                        ws.Cells[i, 5] = 1;
                        ws.Cells[i, 7] = Convert.ToDecimal(range.Cells[i, 6].Value.ToString());
                    }
                    else
                    {
                        if (soluong_thua > 0)
                        {
                            double rdSoluongthua = random.Next(0, (int)soluong_thua);
                            soLuong -= rdSoluongthua;
                            soluong_thua -= rdSoluongthua;
                        }

                        ws.Cells[i, 5] = soLuong;
                        ws.Cells[i, 7] = (decimal)soLuong * Convert.ToDecimal(range.Cells[i, 6].Value.ToString());
                    }
                    excelSaleTotal += Convert.ToDecimal(range.Cells[i, 7].Value.ToString());

                    ws.Cells[i, 7].Borders.Weight = Excel.XlBorderWeight.xlThin;
                    ws.Cells[i, 2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    ws.Cells[i, 4].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    ws.Cells[i, 5].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                }
                //
                ws.Cells[row2s + 1, 3] = "Tổng";
                ws.Cells[row2s + 1, 7] = excelSaleTotal.ToString();
                ws.Cells[row2s + 1, 3].Borders.Weight = Excel.XlBorderWeight.xlThin;
                ws.Cells[row2s + 1, 3].Font.Bold = true;
                ws.Cells[row2s + 1, 7].Borders.Weight = Excel.XlBorderWeight.xlThin;
                ws.Cells[row2s + 1, 7].Font.Bold = true;
                #endregion

                #region WorkSheet Nhập hàng
                //header
                ws2 = wb.Sheets.Add(Type.Missing, Type.Missing, 1, Type.Missing)
                        as Excel.Worksheet;
                ws2.Name = "Nhập hàng";

                ws2.Range[ws2.Cells[1, 1], ws2.Cells[1, cols + 2]].Merge();
                ws2.Range[ws2.Cells[2, 1], ws2.Cells[2, cols + 2]].Merge();
                ws2.Range[ws2.Cells[3, 1], ws2.Cells[3, cols + 2]].Merge();
                ws2.Range[ws2.Cells[4, 1], ws2.Cells[4, cols + 2]].Merge();
                ws2.Range[ws2.Cells[5, 1], ws2.Cells[5, cols + 2]].Merge();
                //Tên công ty
                ws2.Cells[1, 1].Value = companyName.ToUpper();
                ws2.Cells[1, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                ws2.Cells[1, 1].Font.Size = 20;
                //Địa chỉ
                ws2.Cells[2, 1].Value = Address;
                ws2.Cells[2, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                ws2.Cells[2, 1].Font.Size = 14;
                //Điện thoại
                ws2.Cells[3, 1].Value = "'" + Phone;
                ws2.Cells[3, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                ws2.Rows[3].Style.NumberFormat = "General";
                ws2.Cells[3, 1].Font.Size = 12;
                //Tiêu đề
                if (Month == 12)
                    ws2.Cells[4, 1].Value = "Bảng tổng hợp doanh số nhập hàng tháng " + Month + "/" + DateTime.Now.AddYears(-1).Year;
                else
                    ws2.Cells[4, 1].Value = "Bảng tổng hợp doanh số nhập hàng tháng " + Month + "/" + DateTime.Now.Year;
                ws2.Cells[4, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                ws2.Cells[4, 1].Font.Size = 12;
                //Kẻ
                ws2.Cells[6, 1].Borders.Weight = Excel.XlBorderWeight.xlThin;
                //title
                for (int i = 1; i <= cols; i++)
                {
                    if (i < 4) ws2.Cells[6, i + 1] = lswRes[0].Keys.ElementAt(i - 1);
                    if (i == 4) ws2.Cells[6, i + 1] = "SL";
                    if (i == 5) ws2.Cells[6, i + 1] = lswRes[0].Keys.ElementAt(i - 2);
                    ws2.Cells[6, i + 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    ws2.Cells[6, i + 1].Font.Bold = true;
                    ws2.Cells[6, i + 1].Borders.Weight = Excel.XlBorderWeight.xlThin;
                }
                ws2.Cells[6, cols + 2] = "Giá vốn";
                ws2.Cells[6, cols + 2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                ws2.Cells[6, cols + 2].Font.Bold = true;
                ws2.Cells[6, cols + 2].Borders.Weight = Excel.XlBorderWeight.xlThin;
                //
                //Item
                DateTime date2;
                if (Month == 12) date2 = new DateTime(2018, Month, 1);
                else date2 = new DateTime(2019, Month, 1);
                decimal excelBuyTotal = 0;
                autoNum = 1;
                //Vẽ bảng theo ngày
                while (date2.Month == Month)
                {
                    Excel.Range rangeAdd = ws2.UsedRange;
                    var rowsAdd = rangeAdd.Rows.Count;
                    //Thêm date
                    ws2.Cells[rowsAdd + 1, 1] = "'" + date2.ToString("dd/MM/yyyy");
                    ws2.Cells[rowsAdd + 1, 1].Borders.Weight = Excel.XlBorderWeight.xlThin;
                    date2 = date2.AddDays(1);
                    //
                    var orderQuatity = OrderRandom.Next(Convert.ToInt32(SoLuongToiThieu), Convert.ToInt32(SoLuongToiDa));
                    //Random số lượng hàng mỗi ngày
                    for (int i = 1; i <= orderQuatity; i++)
                    {
                        var rowRandom = random.Next(2, lswRes.Count - 1);
                        var item = lswRes.ElementAt(rowRandom);
                        ws2.Cells[rowsAdd + i, 2] = autoNum;
                        autoNum++;
                        ws2.Cells[rowsAdd + i, 2].Borders.Weight = Excel.XlBorderWeight.xlThin;
                        //vẽ từng hàng
                        for (int j = 2; j <= cols; j++)
                        {
                            if (j < 4)
                            {
                                ws2.Cells[rowsAdd + i, j + 1] = item[item.Keys.ToList().ElementAt(j - 1)];
                                ws2.Cells[rowsAdd + i, j + 1].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                            if (j == 5)
                            {
                                if (int.TryParse(item[item.Keys.ToList().ElementAt(j - 2)], out int giaban))
                                {
                                    excelBuyTotal += Convert.ToDecimal(giaban);
                                }
                                else
                                {
                                    excelBuyTotal += 1;
                                }
                                ws2.Cells[rowsAdd + i, j + 1] = item[item.Keys.ToList().ElementAt(j - 2)];
                                ws2.Cells[rowsAdd + i, j + 1].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }
                    }
                }
                //Số lượng trung bình từ tổng nhập và tổng giá bán
                quantity = Math.Round(Convert.ToDouble(BuyTotal / excelBuyTotal));
                //clear excelBuyTotal
                excelBuyTotal = 0;
                soluong_thua = 0;
                Excel.Range range2 = ws2.UsedRange;
                row2s = range2.Rows.Count;
                //Add cột số lượng
                for (int i = 7; i <= row2s; i++)
                {
                    double soLuong = 0;
                    if (quantity <= 1)
                        soLuong = quantity + random.Next(0, 1);
                    else if (quantity <= 2)
                        soLuong = quantity + random.Next(-1, 2);
                    else
                        soLuong = quantity + random.Next(-2, 3);

                    if (soLuong < 0)
                    {
                        soluong_thua = Math.Abs(soLuong);
                        ws2.Cells[i, 5] = 1;
                        ws2.Cells[i, 7] = Convert.ToDecimal(range2.Cells[i, 6].Value.ToString());
                        //ws2.Cells[i, 7] = 0;
                    }
                    else
                    {
                        if(soluong_thua > 0)
                        {
                            double rdSoluongthua = random.Next(0, (int)soluong_thua);
                            soLuong -= rdSoluongthua;
                            soluong_thua -= rdSoluongthua;
                        }
                        ws2.Cells[i, 5] = soLuong;
                        ws2.Cells[i, 7] = (decimal)soLuong * Convert.ToDecimal(range2.Cells[i, 6].Value.ToString());
                    }
                    excelBuyTotal += Convert.ToDecimal(range2.Cells[i, 7].Value.ToString());

                    ws2.Cells[i, 7].Borders.Weight = Excel.XlBorderWeight.xlThin;
                    ws2.Cells[i, 5].Borders.Weight = Excel.XlBorderWeight.xlThin;
                    ws2.Cells[i, 2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    ws2.Cells[i, 4].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    ws2.Cells[i, 5].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                }
                //
                ws2.Cells[row2s + 1, 3] = "Tổng";
                ws2.Cells[row2s + 1, 7] = excelBuyTotal.ToString();
                ws2.Cells[row2s + 1, 3].Borders.Weight = Excel.XlBorderWeight.xlThin;
                ws2.Cells[row2s + 1, 3].Font.Bold = true;
                ws2.Cells[row2s + 1, 7].Borders.Weight = Excel.XlBorderWeight.xlThin;
                ws2.Cells[row2s + 1, 7].Font.Bold = true;
                #endregion

                //Một số thông số định dạng worksheet
                #region FormatStyle
                //
                ws.Columns[1].Style.Font.Size = 12;
                ws.Columns[1].ColumnWidth = 10;
                ws.Columns[1].Style.NumberFormat = "dd/mm/yyyy";

                ws2.Columns[1].Style.Font.Size = 12;
                ws2.Columns[1].ColumnWidth = 10;
                ws2.Columns[1].Style.NumberFormat = "dd/mm/yyyy";

                ws.Columns[2].Style.Font.Size = 12;
                ws.Columns[2].ColumnWidth = 6;
                ws.Columns[2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                ws2.Columns[2].Style.Font.Size = 12;
                ws2.Columns[2].ColumnWidth = 6;
                ws2.Columns[2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                ws.Columns[3].Style.Font.Size = 12;
                ws.Columns[3].ColumnWidth = 35;

                ws2.Columns[3].Style.Font.Size = 12;
                ws2.Columns[3].ColumnWidth = 35;

                ws.Columns[4].Style.Font.Size = 12;
                ws.Columns[4].ColumnWidth = 4;
                ws.Columns[4].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                ws2.Columns[4].Style.Font.Size = 12;
                ws2.Columns[4].ColumnWidth = 4;
                ws2.Columns[4].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                ws.Columns[5].Style.Font.Size = 12;
                ws.Columns[5].ColumnWidth = 4;
                ws.Columns[5].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                ws2.Columns[5].Style.Font.Size = 12;
                ws2.Columns[5].ColumnWidth = 4;
                ws2.Columns[5].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                ws.Columns[6].Style.Font.Size = 12;
                ws.Columns[6].ColumnWidth = 11;
                ws.Columns[6].Style.NumberFormat = "#,##0";

                ws2.Columns[6].Style.Font.Size = 12;
                ws2.Columns[6].ColumnWidth = 11;
                ws2.Columns[6].Style.NumberFormat = "#,##0";

                ws.Columns[7].Style.Font.Size = 12;
                ws.Columns[7].ColumnWidth = 13;
                ws.Columns[7].Style.NumberFormat = "#,##0";

                ws2.Columns[7].Style.Font.Size = 12;
                ws2.Columns[7].ColumnWidth = 13;
                ws2.Columns[7].Style.NumberFormat = "#,##0";

                ws.Rows.Font.Name = "Times New Roman";
                ws.PageSetup.PaperSize = Excel.XlPaperSize.xlPaperA4;
                ws.PageSetup.TopMargin = 27;
                ws.PageSetup.RightMargin = 27;
                ws.PageSetup.BottomMargin = 27;
                ws.PageSetup.LeftMargin = 56;

                ws2.Rows.Font.Name = "Times New Roman";
                ws2.PageSetup.PaperSize = Excel.XlPaperSize.xlPaperA4;
                ws2.PageSetup.TopMargin = 27;
                ws2.PageSetup.RightMargin = 27;
                ws2.PageSetup.BottomMargin = 27;
                ws2.PageSetup.LeftMargin = 56;
                #endregion
                wb.SaveAs(Path + @"\Output_" + Month + ".xlsx");
            }
            catch (Exception ex)
            {
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
                Marshal.ReleaseComObject(ws);
                if (ws2 != null) Marshal.ReleaseComObject(ws2);
                wb.Close(false);
                Marshal.FinalReleaseComObject(wb);
                excelApp.Quit();
                Marshal.FinalReleaseComObject(excelApp);
            }
        }

    }
    public class Data
    {
        [Required(ErrorMessage = "File does not exist")]
        public HttpPostedFileBase fileInput { get; set; }
        [Required]
        public double hangtoithieu { get; set; }
        [Required]
        public double hangtoida { get; set; }
        [Required]
        public string tenkh { get; set; }
        public string diachi { get; set; }
        public string dienthoai { get; set; }
        [Required]
        public decimal tongnhap3 { get; set; }
        [Required]
        public decimal tongnhap2 { get; set; }
        [Required]
        public decimal tongnhap1 { get; set; }
        [Required]
        public decimal tongxuat3 { get; set; }
        [Required]
        public decimal tongxuat2 { get; set; }
        [Required]
        public decimal tongxuat1 { get; set; }
    }
}