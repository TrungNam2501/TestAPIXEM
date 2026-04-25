using OfficeOpenXml;
using ServiceBBAPI.SQLCNN;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Web;

namespace ServiceBBAPI.Models
{
    public class Get_Excel_BB
    {
        public static string Create_Excel(string partno, string effdat, string pathFile, string loaikeo, string slipno, string barcode, string ca, string daylimt, string Soluong, string mesid, string Machno, string pday, string indat, string classs, string intime, string pallet, string OEM, string someSX)
        {
            try
            {
                FileInfo newFile = new FileInfo(pathFile);
                // Neu file da ton tai thi xoa di
                if (newFile.Exists)
                {
                    newFile.Delete(); // ensures we create a new workbook
                    newFile = new FileInfo(pathFile);
                }
                //string maso = "";

                ////-------------------------------

                //string sqlBarcode = "SELECT * FROM prdebe where factory='V'  and prodat='" + pday + "' and partno='" + partno + "' and mesid ='" + mesid + "' order by intime";
                //DataTable dtBar = SQL10_33.ExecuteQuery(sqlBarcode);
                ////lấy barcode theo từng loại keo
                //int a = 1;
                //int b = 2;
                //if (dtBar.Rows.Count > 0)
                //{
                //    for (int i = 0; i < dtBar.Rows.Count; i++)
                //    {
                //        if (i > 0)
                //        {
                //            a += 2;
                //            b = a + 1;
                //        }
                //    }

                //    string sqlweight = "select Plan_Num from [mfnsShareDB].[dbo].[IF_RtPlan2Mixing] where  Plan_Id ='"+mesid+"'"; 
                //    DataTable dtweight = SQLCNN8MAY.ExecuteQuery(sqlweight,Machno);
                //    int weight = Convert.ToInt32(dtweight.Rows[0][0].ToString());
                //    try
                //    {
                //        if (weight == 1)
                //        {
                //            maso = a.ToString();
                //            //string sqlfinish = "UPDATE KEORE SET comp = 'Y' where  recipe_name = '" + partno + "' and mesid = '" + mesid + "' and machno = '" + Machno + "' ";
                //            //DataTable dtfinish = sql.ExecuteQuery(sqlfinish);
                //        }
                //        else
                //        {
                //            maso = a.ToString() + "-" + b.ToString();
                //        }
                //    }
                //    catch (Exception)
                //    {
                //        return "Fail";
                //    }
                //}
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (ExcelPackage package = new ExcelPackage(newFile))
                {
                    //Check loại tem biểu thị
                    string tenloaithe = "";
                    string kichthuoc = "";
                    string KVS = "";
                    string tentheloaikeoTQ = "";
                    string Ca = "";

                    if (loaikeo == "RC")
                    {
                        tenloaithe = "Thẻ biểu thị keo tinh luyện";
                        tentheloaikeoTQ = "精煉膠標示卡";
                        KVS = "KVS3J1C001.9  Rev.5";
                        kichthuoc = "(180mm×130mm×0.08mm)";
                    }

                    if (loaikeo == "RD" || loaikeo == "RR")
                    {
                        tenloaithe = "Thẻ biểu thị keo xúc tiến";
                        tentheloaikeoTQ = "加促膠標示卡";
                        KVS = "KVS3J1C001.8  Rev.5";
                        kichthuoc = "(180mm×130mm×0.08mm)";
                    }

                    if (loaikeo == "RB")
                    {
                        tenloaithe = "Thẻ biểu thị keo cán luyên";
                        tentheloaikeoTQ = "混煉膠標示卡";
                        KVS = "KVS3J1C001.7  Rev.5";
                        kichthuoc = "(180mm×130mm×0.08mm)";
                    }
                    string partnoEDGE = "";
                    if (partno.Length > 7)
                    {
                        partnoEDGE = partno.Substring(7);
                    }


                    
                    if (ca == "1")
                    {
                        Ca = "早 (Sáng)";
                    }
                    if (ca == "2")
                    {
                        Ca = "夜 (Đêm)";
                    }

                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("1");

                    /*if (loaikeo == "RC") //2021-04-22 Bỏ tinh luyện
                     {
                         worksheet.Cells["C4"].Value = "精煉";
                         worksheet.Cells["C5"].Value = "Tinh luyện";
                     }*/

                    worksheet.Cells["C2"].Value = " 建大橡膠（越南）有限公司";
                    worksheet.Cells["C3"].Value = "Công ty Cao su Kenda(Việt Nam)";

                    if (OEM == "OEM")
                    {
                        worksheet.Cells["M4"].Value = "QC-OEM";
                    }


                    worksheet.Cells["E4"].Value = tentheloaikeoTQ;
                    worksheet.Cells["C5"].Value = "Machine BB37" + Machno; //2021-04-22
                    worksheet.Cells["E5"].Value = tenloaithe;
                    worksheet.Cells["N5"].Value = pallet;
                    worksheet.Cells["C6"].Value = "日限  Thời hạn sử dụng:  " + daylimt + "日" + daylimt + "Ngày";
                    worksheet.Cells["C7"].Value = "禁止雨淋，油污，置地，及粉水未乾";
                    worksheet.Cells["C8"].Value = "Cấm ướt mưa, dính dầu, để lên đất, bột nước chưa khô";
                    worksheet.Cells["C9"].Value = "生產日期 Ngày Tháng Sản Xuất";
                    worksheet.Cells["C10"].Value = indat + " " + intime;
                    worksheet.Cells["C11"].Value = "批號";
                    worksheet.Cells["C12"].Value = "Số lô";
                    worksheet.Cells["E11"].Value = slipno;
                    worksheet.Cells["I11"].Value = "重量";
                    worksheet.Cells["I12"].Value = "Trọng lượng";
                    worksheet.Cells["L11"].Style.WrapText = true;
                    worksheet.Cells["L11"].Value = Soluong + "kg";
                    worksheet.Cells["C13"].Value = "有效日";
                    worksheet.Cells["C14"].Value = "Ngày hiệu lực";
                    worksheet.Cells["E13"].Value = effdat + " " + intime;
                    worksheet.Cells["I13"].Value = "班別";
                    worksheet.Cells["I14"].Value = "Ca";
                    worksheet.Cells["L13"].Style.WrapText = true;
                    worksheet.Cells["L13"].Value = Ca;
                    worksheet.Cells["C15"].Value = "規格";
                    worksheet.Cells["C16"].Value = "Quy Cách";
                    worksheet.Cells["D15"].Style.WrapText = true;

                    if (partnoEDGE.Trim() == "")
                        worksheet.Cells["D15"].Value = partno.Substring(0).Trim();
                    else
                        worksheet.Cells["D15"].Value = partno.Substring(0, 7).Trim() + partnoEDGE.Trim();

                    worksheet.Cells["G15"].Value = "編號順序";
                    worksheet.Cells["G17"].Value = "Thứ tự mã số";
                    worksheet.Cells["I15"].Style.WrapText = true;
                    worksheet.Cells["I15"].Value = someSX;
                    worksheet.Cells["K15"].Value = "判 定";
                    worksheet.Cells["K17"].Value = "Phán định";
                    worksheet.Cells["M15"].Style.WrapText = true;
                    worksheet.Cells["M15"].Value = "";
                    worksheet.Cells["C17"].Value = "委託";
                    worksheet.Cells["C18"].Value = "Ủy thác";
                    worksheet.Cells["C20"].Value = KVS;
                    worksheet.Cells["K20"].Value = "*" + barcode + "*";
                    worksheet.Cells["C21"].Value = kichthuoc;
                    worksheet.Cells["K21"].Value = "*" + barcode + "*";

                    // Formatting style all
                    using (var range = worksheet.Cells["B2:P21"])
                    {
                        range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        range.Style.Font.Name = "Arial";
                        range.Style.Font.Size = 13;
                        range.Style.Fill.BackgroundColor.SetColor(Color.White);
                        range.Style.Font.Color.SetColor(Color.Black);

                        // Setting background color dark blue
                        //range.Style.Fill.BackgroundColor.SetColor(Color.DarkBlue);
                        // Setting font color
                        //range.Style.Font.Color.SetColor(Color.White);
                    }


                    //Formatting style for cell
                    worksheet.Cells["M4"].Style.Font.Bold = true;
                    worksheet.Cells["E4"].Style.Font.Bold = true;
                    worksheet.Cells["E5"].Style.Font.Bold = true;  //barcode size
                    worksheet.Cells["K20"].Style.Font.Name = "Code39AzaleaWide2"; //font barcode code39   
                    worksheet.Cells["K20"].Style.Font.Size = 28;
                    worksheet.Cells["C16"].Style.Font.Size = 12;
                    worksheet.Cells["M4"].Style.Font.Size = 16;
                    //Formatting style for cell
                    worksheet.Cells["c4"].Style.Font.Bold = true;
                    //worksheet.Cells["c5"].Style.Font.Bold = true;
                    worksheet.Cells["C5"].Style.Font.Size = 11; //2021-04-22
                    //worksheet.Cells["D15"].Style.Font.Size = 10;
                    //merge cells

                    worksheet.Cells["C2:O2"].Merge = true;
                    worksheet.Cells["C3:O3"].Merge = true;
                    worksheet.Cells["E4:L4"].Merge = true;
                    worksheet.Cells["M4:O4"].Merge = true; //2024-01-22
                    worksheet.Cells["C5:D5"].Merge = true; //2021-04-22
                    worksheet.Cells["E5:L5"].Merge = true;
                    worksheet.Cells["N5:O5"].Merge = true;
                    worksheet.Cells["C6:O6"].Merge = true;
                    worksheet.Cells["C7:O7"].Merge = true;
                    worksheet.Cells["C8:O8"].Merge = true;
                    worksheet.Cells["C9:O9"].Merge = true;
                    worksheet.Cells["C10:O10"].Merge = true;
                    worksheet.Cells["C11:D11"].Merge = true;
                    worksheet.Cells["C12:D12"].Merge = true;
                    worksheet.Cells["E11:H12"].Merge = true;
                    worksheet.Cells["I11:K11"].Merge = true;
                    worksheet.Cells["I12:K12"].Merge = true;
                    worksheet.Cells["L11:O12"].Merge = true;
                    worksheet.Cells["C13:D13"].Merge = true;
                    worksheet.Cells["C14:D14"].Merge = true;
                    worksheet.Cells["E13:H14"].Merge = true;
                    worksheet.Cells["I13:K13"].Merge = true;
                    worksheet.Cells["I14:K14"].Merge = true;
                    worksheet.Cells["L13:O14"].Merge = true;
                    worksheet.Cells["D15:F16"].Merge = true;
                    worksheet.Cells["G15:H16"].Merge = true;
                    worksheet.Cells["G17:H18"].Merge = true;
                    worksheet.Cells["I15:J18"].Merge = true;
                    worksheet.Cells["K15:L16"].Merge = true;
                    worksheet.Cells["K17:L18"].Merge = true;
                    worksheet.Cells["M15:O18"].Merge = true;
                    worksheet.Cells["D17:F18"].Merge = true;
                    worksheet.Cells["C20:G20"].Merge = true;
                    worksheet.Cells["K20:O20"].Merge = true;
                    worksheet.Cells["C21:G21"].Merge = true;
                    worksheet.Cells["K21:O21"].Merge = true;

                    //Formatting Width,Height
                    worksheet.Row(1).Height = 5;
                    worksheet.Row(2).Height = 20.5;
                    worksheet.Row(7).Height = 22.5;
                    worksheet.Row(8).Height = 22.5;
                    worksheet.Row(19).Height = 5;
                    worksheet.Row(20).Height = 30;
                    worksheet.Row(22).Height = 5.5;

                    worksheet.Column(1).Width = 2;
                    worksheet.Column(2).Width = 2;
                    worksheet.Column(3).Width = 10;
                    worksheet.Column(4).Width = 5;
                    worksheet.Column(5).Width = 5;
                    worksheet.Column(6).Width = 5;
                    worksheet.Column(7).Width = 8;
                    worksheet.Column(8).Width = 8;
                    worksheet.Column(9).Width = 5;
                    worksheet.Column(10).Width = 5;
                    worksheet.Column(11).Width = 5;
                    worksheet.Column(12).Width = 7;
                    worksheet.Column(13).Width = 5;
                    worksheet.Column(14).Width = 5;
                    worksheet.Column(15).Width = 5;
                    worksheet.Column(16).Width = 2;
                    worksheet.Column(17).Width = 2;

                    //Formatting Align cells
                    worksheet.Cells["A1:P22"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    worksheet.Cells["A1:P22"].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;


                    worksheet.Cells["D15"].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                    worksheet.Cells["D15"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    worksheet.Cells["C7"].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Bottom;
                    worksheet.Cells["C8"].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Top;
                    worksheet.Cells["G15"].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Bottom;
                    worksheet.Cells["K15"].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Bottom;
                    worksheet.Cells["G17"].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Top;
                    worksheet.Cells["K17"].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Top;
                    worksheet.Cells["E14"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    worksheet.Cells["C2"].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Bottom;

                    // Setting  borde
                    worksheet.Cells["C6:O18"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);

                    var cell = worksheet.Cells["C6:O18"];
                    var border = cell.Style.Border;
                    border.Top.Style = border.Left.Style = border.Right.Style = border.Bottom.Style =
                        OfficeOpenXml.Style.ExcelBorderStyle.Thin;

                    cell = worksheet.Cells["C7:O7"];
                    border = cell.Style.Border;
                    border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.None;
                    cell = worksheet.Cells["C8:O8"];
                    border = cell.Style.Border;
                    border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.None;

                    cell = worksheet.Cells["C11:K11"];
                    border = cell.Style.Border;
                    border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.None;
                    cell = worksheet.Cells["C12:K12"];
                    border = cell.Style.Border;
                    border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.None;

                    cell = worksheet.Cells["C13:D13"];
                    border = cell.Style.Border;
                    border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.None;
                    cell = worksheet.Cells["C14:D14"];
                    border = cell.Style.Border;
                    border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.None;

                    cell = worksheet.Cells["I13:K13"];
                    border = cell.Style.Border;
                    border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.None;
                    cell = worksheet.Cells["I14:K14"];
                    border = cell.Style.Border;
                    border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.None;

                    cell = worksheet.Cells["C15"];
                    border = cell.Style.Border;
                    border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.None;
                    cell = worksheet.Cells["C16"];
                    border = cell.Style.Border;
                    border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.None;

                    cell = worksheet.Cells["C17"];
                    border = cell.Style.Border;
                    border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.None;
                    cell = worksheet.Cells["C18"];
                    border = cell.Style.Border;
                    border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.None;

                    cell = worksheet.Cells["G16:H16"];
                    border = cell.Style.Border;
                    border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.None;
                    cell = worksheet.Cells["G17:H17"];
                    border = cell.Style.Border;
                    border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.None;

                    cell = worksheet.Cells["K16:L16"];
                    border = cell.Style.Border;
                    border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.None;
                    cell = worksheet.Cells["K17:L17"];
                    border = cell.Style.Border;
                    border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.None;

                    cell = worksheet.Cells["B2:P22"];
                    cell.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Medium);

                    // Setting Margin
                    worksheet.PrinterSettings.TopMargin = (decimal)0.2;
                    worksheet.PrinterSettings.LeftMargin = (decimal)0.8;
                    worksheet.PrinterSettings.RightMargin = (decimal)0.5;
                    worksheet.PrinterSettings.PaperSize = ePaperSize.A5;
                    worksheet.PrinterSettings.Orientation = eOrientation.Landscape;

                    package.Save();
                    //return newFile.FullName;
                    return "";
                }

            }
            catch { return "Lỗi file Excel!"; }
        }

    }
}