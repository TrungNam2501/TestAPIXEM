using ServiceBBAPI.Models;
using ServiceBBAPI.SQLCNN;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.NetworkInformation;
using System.Numerics;
using System.Web;
using System.Web.Caching;
using System.Web.Services;
using System.Windows.Media.Media3D;

namespace ServiceBBAPI
{
    /// <summary>
    /// Summary description for ServiceBBintem
    /// </summary>
    [WebService(Namespace = "http://tempuri.org/")]
    [WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
    [System.ComponentModel.ToolboxItem(false)]
    // To allow this Web Service to be called from script, using ASP.NET AJAX, uncomment the following line. 
    // [System.Web.Script.Services.ScriptService]
    public class ServiceBBintem : System.Web.Services.WebService
    {
        private Ping ping = new Ping();
        private bool b;
        [WebMethod]
        public string HelloWorld()
        {
            return "Hello World";
        }
        [WebMethod]
        public string[] Login(string empno)
        {
            string[] Result = new string[4];
            DataTable dt = new DataTable();
            string sql;
            sql = "Select name,depno, empno from peremp where empno = '" + empno + "' and subno = '4'";
            dt = SQL10_34.ExecuteQuery(sql);
            if (dt.Rows.Count == 0)
            {
                Result[0] = "Số thẻ không tồn tại";
                return Result;
            }
            Result[0] = "";
            Result[1] = dt.Rows[0][0].ToString().Trim();
            Result[2] = dt.Rows[0][1].ToString().Trim();
            Result[3] = dt.Rows[0][2].ToString().Trim();
            return Result;
        }
        [WebMethod]
        public string[] may()
        {
            string[] Result = new string[8];
          
            Result[0] = "01";
            Result[1] = "02";
            Result[2] = "03";
            Result[3] = "04";
            Result[4] = "05";
            Result[5] = "06";
            Result[6] = "07";
            Result[7] = "08";
            return Result;
        }
        [WebMethod]
        public string Get_shift()
        {
            //--------------------------------
            //  lấy ca làm việc theo thời gian
            //--------------------------------

            string shift;
            DateTime dNow = DateTime.Now;

          
            DateTime dFrom1 = DateTime.Now;
            TimeSpan tsFrom1 = new TimeSpan(7, 30, 0); 
            dFrom1 = dFrom1.Date + tsFrom1;

            DateTime dTo1 = DateTime.Now;
         
            TimeSpan tsTo1 = new TimeSpan(19, 00, 0); 
        
            dTo1 = dTo1.Date + tsTo1;
          

            if (dNow >= dFrom1 && dNow <= dTo1)
                shift = "1";
            else
                shift = "2";
            return shift;
        }
        private string setPday(string shift)
        {
            //------------------------------------------
            //nếu là ca 2 mà thời gian insert từ 0h sáng hôm sau -> 6h20 sáng thì pday phải là ngày hiện tại - 1
            //------------------------------------------

            string sPday = DateTime.Now.ToString("yyyyMMdd");

            DateTime dFrom = DateTime.Now;  //tao biến datetime thời gian bắt đầu từ 0h sáng hôm sau
            TimeSpan tsFrom = new TimeSpan(0, 0, 0);
            dFrom = dFrom.Date + tsFrom;

            DateTime dTo = DateTime.Now; //tao biến datetime thời gian đến 6h20 sáng hôm sau
            TimeSpan tsTo = new TimeSpan(7, 30, 0); // 2021-05-19
            dTo = dTo.Date + tsTo;

            DateTime dNow = DateTime.Now;
         

            if (shift == "2") //Nếu là ca 2
            {
                if (dNow >= dFrom && dNow <= dTo)
                {
                    DateTime Yesterday = DateTime.Now.AddDays(-1);   //ngày hiện tại - 1
                    sPday = Yesterday.ToString("yyyyMMdd");
                }
            }
            return sPday;
        }
        private string setPdayBB(string shift)
        {
            //------------------------------------------
            //nếu là ca 2 mà thời gian insert từ 0h sáng hôm sau -> 6h20 sáng thì pday phải là ngày hiện tại - 1
            //------------------------------------------

            string sPday = DateTime.Now.ToString("yyyyMMdd");

            DateTime dFrom = DateTime.Now;  //tao biến datetime thời gian bắt đầu từ 0h sáng hôm sau
            TimeSpan tsFrom = new TimeSpan(0, 0, 0);
            dFrom = dFrom.Date + tsFrom;

            DateTime dTo = DateTime.Now; //tao biến datetime thời gian đến 6h20 sáng hôm sau
            TimeSpan tsTo = new TimeSpan(7, 30, 0); // 2021-05-19
            dTo = dTo.Date + tsTo;

            DateTime dNow = DateTime.Now;


            if (shift == "1")
            {
                if (dNow >= dFrom && dNow <= dTo)
                {
                    DateTime Yesterday = DateTime.Now.AddDays(-1);
                    sPday = Yesterday.ToString("yyyyMMdd");
                }
            }
            return sPday;
        }
        [WebMethod]
        public string Get_shiftMayBB()
        {
            //--------------------------------
            //  lấy ca làm việc theo thời gian
            //--------------------------------

            string shift;
            DateTime dNow = DateTime.Now;


            DateTime dFrom1 = DateTime.Now;
            TimeSpan tsFrom1 = new TimeSpan(7, 30, 0);
            dFrom1 = dFrom1.Date + tsFrom1;

            DateTime dTo1 = DateTime.Now;

            TimeSpan tsTo1 = new TimeSpan(19, 00, 0);

            dTo1 = dTo1.Date + tsTo1;


            if (dNow >= dFrom1 && dNow <= dTo1)
                shift = "0";
            else
                shift = "1";
            return shift;
        }
        [WebMethod]
        public DataSet Get_Mesid_list(string Machno) 
        {
            string shift_id = Get_shiftMayBB();
           
            string pday = setPdayBB(shift_id);
            DateTime myPday = DateTime.ParseExact(pday, "yyyyMMdd", CultureInfo.InvariantCulture);
            pday = myPday.ToString("yyyy-MM-dd");

            // thêm  join vào bảng idgrouplot để biết mes nào đang chạy
            string sqlCheckMes = "select a.Plan_Id, a.Recipe_Code FROM [mfnsShareDB].[dbo].[IF_RtPlan2Mixing] a, [mfns].[dbo].[Ppt_GroupLot] b where a.Shift_Id ='" + shift_id + "' and a.P_Date ='" + pday + "' and a.Plan_Id not like 'V%' and a.Plan_Id = b.MesPlanID and b.End_datetime is not null";


            DataTable dtMes = SQLCNN8MAY.ExecuteQuery(sqlCheckMes,Machno);
            DataSet ds = new DataSet();
            ds.Tables.Add(dtMes);
            return ds;
        }
        [WebMethod]
        public string[] Printer()
        {
            DataTable dtPrinter = SQL10_33.ExecuteQuery("SELECT [TenMay],[MaMay] FROM [BB].[dbo].[Printer_BB] where IP='BB'");


            string[] print = new string[dtPrinter.Rows.Count];

           
            for (int i = 0; i < dtPrinter.Rows.Count; i++)
            {
                print[i] = dtPrinter.Rows[i][1].ToString().Trim();
            }

            return print;


        }

        [WebMethod]
        public string Finish(string tenkeo, string mesid, string machno) 
        {
            string result = "";
            //string sqlfinish = "UPDATE KEOREKV1 SET comp = 'Y' where subno='4' and factory='E' and  recipe_name = '" + tenkeo + "' and mesid = '" + mesid + "' and machno = '" + machno + "' ";

            //bool b1 = sql.SQLExecuteNonQuery(sqlfinish);
            //if (!b1)
            //    result = "ERROR";
            //else
            //    result = "";
            return result;
        }

        [WebMethod]
        public DataSet Get_Barcode(string mesid, string machno)
        {
            string sqlBar = "Select barcode,weight from prdebe where subno='4' and factory='V' and mesid = '" + mesid + "' and machno ='" + "V-BB37" + machno + "'";
            DataTable dtBar = SQL10_33.ExecuteQuery(sqlBar);
            DataSet ds = new DataSet();
            ds.Tables.Add(dtBar);
            return ds;
        }

        [WebMethod]
        public DataSet Get_Mesid_list_Inlai(string Machno) 
        {
            string shift_id = Get_shiftMayBB();
            
            string pday = setPdayBB(shift_id);
          
            DateTime myPday = DateTime.ParseExact(pday, "yyyyMMdd", CultureInfo.InvariantCulture);
            pday = myPday.ToString("yyyyMMdd");
            string pday1 = myPday.ToString("yyyy-MM-dd");
            // thêm  join vào bảng idgrouplot để biết mes nào đang chạy
            string sqlCheckMes = "select a.Plan_Id, a.Recipe_Code FROM [mfnsShareDB].[dbo].[IF_RtPlan2Mixing] a, [mfns].[dbo].[Ppt_GroupLot] b where a.Shift_Id ='" + shift_id + "' and a.P_Date ='" + pday1 + "'  and a.Plan_Id = b.MesPlanID and b.End_datetime is not null";
            DataTable dtMes = SQLCNN8MAY.ExecuteQuery(sqlCheckMes,Machno);
            DataSet ds = new DataSet();
            ds.Tables.Add(dtMes);
            return ds;
        }
        [WebMethod(Description = "Lấy IP máy PDA đang gọi WebService")]
        public string GetClientIp()
        {
            string ip = Context.Request.ServerVariables["HTTP_X_FORWARDED_FOR"];
            if (string.IsNullOrEmpty(ip))
            {
                ip = Context.Request.ServerVariables["REMOTE_ADDR"];
            }
            return ip;
        }

        [WebMethod]
        public string Print_BB(string tenkeo, string Machno, string printername, string ca, string usrno, string candao, string Soluong, string mesid, string pallet, string mac_pda,string OEM, string depno)
        {
            #region Check MAC PDA 
            mac_pda = mac_pda.Trim();
            string Mac_date = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            try
            {
                string sql_selectMac = "SELECT * FROM [QL PDA].[dbo].[KV2 PDA] where MacAddress='" + mac_pda + "' ";
                DataTable dt_slcMac = SQL10_33.ExecuteQuery(sql_selectMac);
                if (dt_slcMac.Rows.Count > 0)
                {
                    string sql_updateMac = "Update [QL PDA].[dbo].[KV2 PDA] set Intime='" + Mac_date + "' where MacAddress = '" + mac_pda + "'";
                    bool chk_updateMac = SQL10_33.ExecuteNonQuery(sql_updateMac);
                }
                else
                {
                    string IP = GetClientIp();
                    string sql_selectIP = "SELECT * FROM [QL PDA].[dbo].[KV2 PDA] where IP='" + IP.Trim() + "' ";
                    string sql_updateMac = "Update [QL PDA].[dbo].[KV2 PDA] set Intime='" + Mac_date + "', MacAddress ='"+ mac_pda + "' where IP = '" + IP.Trim() + "'";
                    bool chk_updateMac = SQL10_33.ExecuteNonQuery(sql_updateMac);

                }
             
            }
            catch { }
            #endregion

            string filename = "_" + printername + ".xlsx";
            string pathfolder = System.AppDomain.CurrentDomain.BaseDirectory + @"Data\";
            string pathfile = System.AppDomain.CurrentDomain.BaseDirectory + @"Data\" + filename; //tạo file và folder cho file Excel
            if (!Directory.Exists(pathfolder)) // Nếu chưa có folder thì tạo cái mới
            {
                Directory.CreateDirectory(pathfolder);
            }

            string shift_id = Get_shift();
            string pday = setPday(shift_id);
            DateTime myPday = DateTime.ParseExact(pday, "yyyyMMdd", CultureInfo.InvariantCulture);
            pday = myPday.ToString("yyyyMMdd");
            string description = "";
            string indat = DateTime.Now.ToString("yyyyMMdd");
            string intime = DateTime.Now.ToString("HH:mm:ss");

            //string sqlRecipe = "select mater_name from pmt_recipe";
            string Plan_Id = "";
            string result = "";
            string partno = tenkeo.Trim();
            string classs = shift_id;
            string slipno = classs + Machno + "-" + pday.Substring(4, 4);
            string barcode = "";
            string effdat = "";
            string daylimt = "";
            string ptype = "";
            string machno = "V-BB37" + Machno;
            string somesx = "";
         
            ca = shift_id;
            string makeo = "";
            effdat = myPday.AddDays(7).ToString("yyyyMMdd");
            string syear = myPday.ToString("yyyy").Substring(2, 2);
            string sday = myPday.ToString("dd");
            string smonth = myPday.ToString("MM");
            switch (smonth)
            {
                case "10": smonth = "A"; break;
                case "11": smonth = "B"; break;
                case "12": smonth = "C"; break;
                default: smonth = smonth.Substring(1, 1); break;
            }
            string spday = syear + smonth + sday;
         

            string sqlRTPLAN = "select Plan_Id, P_Date from [mfnsShareDB].[dbo].[IF_RtPlan2Mixing] where  Plan_Id ='"+mesid+"'";
            DataTable dtKEORE = new DataTable();
            dtKEORE = SQLCNN8MAY.ExecuteQuery(sqlRTPLAN, Machno);
            if (dtKEORE.Rows.Count > 0)
            {
                if (dtKEORE.Rows[0][1].ToString().Trim().Replace("-","") != pday.Trim())
                {
                    string s = "Mes quá giờ không quét được";
                    return s;
                }
                if (tenkeo.Length <= 5)
                {
                    makeo = "RD"; ptype = "3";
                }
                else
                {
                    switch (tenkeo.Trim().Substring(6).ToUpper())
                    {
                        case "RM": makeo = "RD"; ptype = "3"; break;
                        case "1": makeo = "RB"; ptype = "2"; break;
                        case "9": makeo = "RD"; ptype = "3"; break;
                        case "2": makeo = "RC"; ptype = "2"; break;
                        case "3": makeo = "RC"; ptype = "2"; break;
                        case "4": makeo = "RC"; ptype = "2"; break;
                        case "5": makeo = "RC"; ptype = "2"; break;
                        case "RE": makeo = "RR"; ptype = "3"; break;
                        case "R": makeo = "RR"; ptype = "3"; break;
                        case "92": makeo = "RD"; ptype = "3"; break;
                        case "": makeo = "RD"; ptype = "3"; break;

                        case "1-EDGE": makeo = "RB"; ptype = "2"; break;
                        case "2-EDGE": makeo = "RC"; ptype = "2"; break;
                        case "3-EDGE": makeo = "RC"; ptype = "2"; break;
                        case "4-EDGE": makeo = "RC"; ptype = "2"; break;
                        case "5-EDGE": makeo = "RC"; ptype = "2"; break;
                        case "9-EDGE": makeo = "RD"; ptype = "3"; break;
                        case "EDGE": makeo = "RD"; ptype = "3"; break;

                        case "1EDGE": tenkeo = partno.Substring(0, 5) + "-1-EDGE"; makeo = "RB"; ptype = "2"; break;
                        case "2EDGE": tenkeo = partno.Substring(0, 5) + "-2-EDGE"; makeo = "RC"; ptype = "2"; break;
                        case "3EDGE": tenkeo = partno.Substring(0, 5) + "-3-EDGE"; makeo = "RC"; ptype = "2"; break;
                        case "4EDGE": tenkeo = partno.Substring(0, 5) + "-4-EDGE"; makeo = "RC"; ptype = "2"; break;
                        case "5EDGE": tenkeo = partno.Substring(0, 5) + "-5-EDGE"; makeo = "RC"; ptype = "2"; break;
                        case "9EDGE": tenkeo = partno.Substring(0, 5) + "-9-EDGE"; makeo = "RD"; ptype = "3"; break;

                        case "1THU": tenkeo = partno.Substring(0, 5) + "-1THU"; makeo = "RB"; ptype = "2"; break;
                        case "2THU": tenkeo = partno.Substring(0, 5) + "-2THU"; makeo = "RC"; ptype = "2"; break;
                        case "3THU": tenkeo = partno.Substring(0, 5) + "-3THU"; makeo = "RC"; ptype = "2"; break;
                        case "4THU": tenkeo = partno.Substring(0, 5) + "-4THU"; makeo = "RC"; ptype = "2"; break;
                        case "5THU": tenkeo = partno.Substring(0, 5) + "-5THU"; makeo = "RC"; ptype = "2"; break;
                        case "9THU": tenkeo = partno.Substring(0, 5) + "-9THU"; makeo = "RD"; ptype = "3"; break;
                    }
                }
              

               
                string ktptype = "  SELECT [ptype], LTRIM(RTRIM([rubno_7])) rubno_7 FROM[InTem].[dbo].[rubnod_Ptype] WHERE SUBSTRING(rubno_7,7,1)= '2' AND rubno_7 = '" + partno.Trim() + "'";
                DataTable KT_ptype = sql_KEORE.ExecuteQuery(ktptype);
                if (KT_ptype.Rows.Count >= 2)
                {
                    result = "Liên hệ phòng thí nghiệm (a Thuần) đóng 1 tiêu chuẩn";
                    return result;
                }
                else
                {
                    if (KT_ptype.Rows.Count > 0)
                    {
                        makeo = "RB";
                        ptype = KT_ptype.Rows[0][0].ToString().Trim();
                       
                    }
                    
                }

                // Kiểm tra hạn sử dụng
                string getKeo = "select expday from [erp].[dbo].[prdexp] where subno='4' and factory='V' and  ptype ='" + makeo + "' and rubno='" + partno.Substring(0, 5) + "'";
                DataTable keo = SQL10_33.ExecuteQuery(getKeo);

                if (keo.Rows.Count == 0)
                {
                    result = "Mã keo không được sử dụng.\n Liên hệ Duyên phòng chế tạo (755) !";
                    return result;
                }
                else
                {
                    int day = int.Parse(keo.Rows[0][0].ToString().Trim());
                    daylimt = day.ToString().Trim();
                    effdat = DateTime.Now.AddDays(day).ToString("yyyyMMdd");
                }


               
                string sqlBarcode = "SELECT MAX(SUBSTRING(Barcode,8,3)) FROM [erp].[dbo].[prdebe] where subno='4' and factory='V' and barcode like '%" + makeo + "%' and prodat = '" + pday + "'"; 
                DataTable dtBar = SQL10_33.ExecuteQuery(sqlBarcode);
              

                if (dtBar.Rows.Count == 1 && dtBar.Rows[0][0].ToString().Trim() == "")
                    barcode = makeo + spday + "001";
                else
                    barcode = makeo + spday + (int.Parse(dtBar.Rows[0][0].ToString()) + 1).ToString("000");
            }
            else
            {
                result = "Mã MES đã bị đóng! Liên hệ IT mở!";
                return result;
            }
         
            string idGrouplot = "";
            string br1 = "";

            #region kiem tra gioi han keo

            if (float.Parse(Soluong) < 30)
            {
                return "Lỗi! Trọng lượng không phù hợp!";
            }

            string ktKeo2 = "select * from [InTem].[dbo].[rubnod_Ptype] where [rubno_7] ='" + tenkeo.Trim() + "'";
            DataTable kiemkeo2 = sql_KEORE.ExecuteQuery(ktKeo2);

          

            string plannum = "SELECT Plan_Num FROM [mfnsShareDB].[dbo].[IF_RtPlan2Mixing] where Plan_Id ='" + mesid + "'";
            DataTable plann = SQLCNN8MAY.ExecuteQuery(plannum, Machno);
            if (plann.Rows.Count > 0)
            {
                br1 = plann.Rows[0][0].ToString().Trim();
            }

            string checkmes = " Select  Plan_Id,Plan_Num,Recipe_Code from [mfnsShareDB].[dbo].[IF_RtPlan2Mixing] where  replace(Recipe_Name,'-','') = '" + partno.Replace("-", "") + "' and Plan_Id ='" + mesid + "' ";
            DataTable dtMEsid = SQLCNN8MAY.ExecuteQuery(checkmes,Machno);

            if (dtMEsid.Rows.Count == 0)
            {
                result = "MES không tồn tại  IF_RtPlan2Mixing , tạo MES khác!";
                return result;
            }
           
            Plan_Id = dtMEsid.Rows[0][0].ToString().Trim();

            string recipe_name = dtMEsid.Rows[0][2].ToString().Trim();
            string planqty = dtMEsid.Rows[0][1].ToString().Trim();

           
            string DateNow = DateTime.Now.ToString("yyyy-MM-dd 06:30:00").ToString();
            string timeH = DateTime.Now.ToString("HHmm");
            if (int.Parse(timeH) <= 630)
            {
                DateNow = DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd 18:30:00").ToString();
            }
            //2024/10/01
            //string ktMESplan = " select b.FinishNum * (select sum(set_weight) from[mfns].[dbo].[pmt_weigh] where father_code = b.RecipeName) from [mfns].[dbo].[Ppt_GroupLot] b " +
            //    " where MesPlanID = '" + Plan_Id + "' and RecipeName = '" + recipe_name + "' and SetNumber = '" + planqty + "' and End_datetime is not null";

            string ktMESplan = " select b.FinishNum * (select sum(set_weight) from[mfns].[dbo].[pmt_weigh] where father_code = b.RecipeName) from [mfns].[dbo].[Ppt_GroupLot] b " +
               " where MesPlanID = '" + Plan_Id + "' and RecipeName = '" + recipe_name + "'  and End_datetime is not null";

            DataTable ktplan = SQLCNN8MAY.ExecuteQuery(ktMESplan, Machno);

            float KeoSX = 0;
            float GioiHanKeo = 0;
            if (makeo.Substring(0, 1) == "R")
            {
                string LayKeoQuetTem = " select ISNULL(sum([weight]),0) from [erp].[dbo].[prdebe] where subno='4' and factory='V' and machno ='" + machno + "' and mesid ='" + Plan_Id + "' ";
                DataTable getWeightSX = SQL10_33.ExecuteQuery(LayKeoQuetTem);

                GioiHanKeo = float.Parse(ktplan.Rows[0][0].ToString().Trim());
                KeoSX = float.Parse(getWeightSX.Rows[0][0].ToString());
                float KeoVo = KeoSX + float.Parse(Soluong);

                if (KeoVo > GioiHanKeo)
                {
                    if (GioiHanKeo < KeoSX)
                    {
                        result = "MES này quá số lượng kế hoạch, không thể quét tiếp!";
                        return result;
                    }
                    else
                    {
                        result = "Lỗi! MES này chỉ quét được " + (GioiHanKeo - KeoSX).ToString().Trim() + "KG nữa!";
                        return result;
                    }
                }
            }
            #endregion kiem tra gioi han keo

          

            if (pallet.Length < 6)
            {
                result = "Pallet không đủ ký tự, nhập lại!!!";
                return result;
            }
         
            if (pallet.Substring(0, 2) != "VB" && pallet.Substring(0, 2) != "EB" && pallet.Substring(0, 2) != "VC" && pallet.Substring(0, 2) != "VD" && pallet.Substring(0, 2) != "VE" && pallet.Substring(0, 2) != "VF" && pallet.Substring(0, 1) != "V")
            {
                result = "Pallet Không hợp lệ, nhập lại!!!";
                return result;
            }
            string SelPallet = "SELECT PALLET_NO FROM [InTem].[dbo].[PalletBB] WHERE PALLET_NO='" + pallet + "'";
            DataTable dtPallet = sql_KEORE.ExecuteQuery(SelPallet);

            if (dtPallet.Rows.Count == 0)
            {
                string sDat = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                string sqlIns = "INSERT INTO [InTem].[dbo].[PalletBB] VALUES('" + pallet + "', '" + sDat + "', '" + usrno + "','1', 'Y') ";

                bool bQuery = sql_KEORE.ExecuteNonQuery(sqlIns);
                if (!bQuery)
                {
                    result = "Lỗi cập nhật Pallet\nPalletBB";
                    return result;
                }
            }

           
            string ktXuatPallet = "select top 1 active from [dbo].[prdebe] where subno ='4' and factory='V' and pallet_no='" + pallet.Trim() + "' order by indat desc, intime desc";
            DataTable kttpallet = SQL10_33.ExecuteQuery(ktXuatPallet);
            if (kttpallet.Rows.Count > 0)
            {
                if (kttpallet.Rows[0][0].ToString() == "N")
                {
                    result = "Pallet này chưa xuất, không được trùng pallet";
                    return result;
                }
            }
            string ktTrungPallet = "select * from [dbo].[prdebe] where subno ='4' and factory='V' and mesid='" + mesid.Trim() + "' and pallet_no='" + pallet.Trim() + "' ";
            DataTable kttpallet1 = SQL10_33.ExecuteQuery(ktTrungPallet);
            if (kttpallet1.Rows.Count > 0)
            {
                result = "1 Pallet chỉ được quét 1 lần cho 1 mã MES";
                return result;
            }

          
            string active = "Y";
            if (ptype == "3")
            {
                active = "N";
            }
            if (makeo == "RB") 
            {
                candao = "N";
            }
           
            if (ptype == "2")
            {
                active = "";
            }
            if (OEM != "")
            {
                string insertTemOEM = "INSERT INTO [BB].[dbo].[TemOEMBB] ([mesid], [partno], [Barcode], [indat], [intime]) VALUES('"+Plan_Id+"', '"+partno+"', '"+barcode+"', '"+indat+"', '"+intime+"'); ";
                bool insertTemOEM2 = SQL10_33.ExecuteNonQuery(insertTemOEM);

            }

            string weightRecipe = "SELECT SUM(set_weight) as weightRecipe FROM [mfns].[dbo].[pmt_weigh] where father_code ='"+partno+"'";
            DataTable weightRC = SQLCNN8MAY.ExecuteQuery(weightRecipe, machno);
            if (weightRC.Rows[0][0].ToString().Trim() == "")
            {
                result = "Không tìm thấy số kg tiêu chuẩn";
                return result;
            }
            string weight = "";
            weight = weightRC.Rows[0][0].ToString().Trim();

            #region tính số mẻ cũ

            //double weightValue = double.Parse(weight) * 1.02;
            ////thêm 2% so với tc
            //string ktSotemin = "SELECT *  FROM [erp].[dbo].[prdebe] where mesid ='"+Plan_Id+"' order by indat desc, intime desc";
            //DataTable dtSkgTemin = SQLCNN8MAY.ExecuteQuery(ktSotemin, "33");
            //if (dtSkgTemin.Rows.Count == 0)
            //{
            //    if(double.Parse(Soluong) > weightValue)
            //    {
            //        somesx = "1-2";

            //    }
            //    if (double.Parse(Soluong) <= weightValue)
            //    {
            //        somesx = "1";
            //    }
            //}
            //if (dtSkgTemin.Rows.Count > 0)
            //{
            //    string someLonnhat = dtSkgTemin.Rows[0]["some_sx"].ToString().Trim();
            //    string lastCharStr = someLonnhat.Split('-').Last();
            //    if (double.Parse(Soluong) > weightValue)
            //    { 
            //        int meA = int.Parse(lastCharStr)+1;
            //        int meB = int.Parse(lastCharStr) + 2;

            //        somesx = meA.ToString()+"-"+meB.ToString();   


            //    }
            //    if (double.Parse(Soluong) <= weightValue)
            //    {
            //        int meA = int.Parse(lastCharStr) + 1;
            //        somesx = meA.ToString();
            //    }

            //}
            #endregion

             
            double kgTieuChuan = double.Parse(weight); //số kg tiêu chuẩn
            double kgIn = double.Parse(Soluong);  // số kg nhập vào in 

            double kgDaIn = KeoSX;  //số keo đã intem

            double epsilon = 0.000001; // dùng tránh lỗi double ko quan tâm làm chi cả 

            // tính mẻ
            int start = (int)Math.Floor(kgDaIn / kgTieuChuan) + 1;   // tính số mẻ bắt đầu 
            int end = (int)Math.Floor((kgDaIn + kgIn - epsilon) / kgTieuChuan) + 1;// tính số mẻ  kết thúc 

            // giới hạn không vượt số mẻ điều động
            int maxBatch = (int)Math.Ceiling(GioiHanKeo / kgTieuChuan);    // đây là tổng trong lượng chia cho tiêu chuẩn,  theo finishnum 
            if (end > maxBatch) end = maxBatch; // nếu end vươt finish num thì là finishnum

            // build chuỗi
           
            if (start == end)
            {
                somesx = start.ToString();
            } 
            else
            {
                somesx = string.Join("-", Enumerable.Range(start, end - start + 1));  // cắ chuỗi số mẻ
            }









            string machnoinsert = machno;

            string sql1 = "INSERT INTO [dbo].[prdebe] VALUES ('4' , 'V', '" + Plan_Id + "','" + machnoinsert + "', '" + daylimt + "','" + barcode + "','" + slipno + "','" + Soluong + "', '" + pday + "',  '" + effdat + "', " +
            "'" + classs + "', '" + ptype + "','" + candao.Trim() + "','" + partno + "', '" + intime + "', '" + indat + "', '" + usrno + "','" + pallet + "','" + active + "','" + somesx + "')";
            bool b1 = SQL10_33.ExecuteNonQuery(sql1);

            //bool b1 = true;

            if (!b1)
            {
                result = "Đợi 1 chút rồi quét lại!!!";
                return result;
            }
            else
            {
                
              
                string sqlip = "SELECT ip FROM [erp].[dbo].[P8500_IP] where machno like 'V-BB%' ";
                DataTable ip = SQL10_33.ExecuteQuery(sqlip);
                string IP = "";

                if (ip.Rows.Count > 0)
                {
                    foreach (DataRow item in ip.Rows)
                    {
                        try
                        {
                            //IP = item["ip"].ToString().Trim();
                            //string br = "";

                            //PingReply PR = ping.Send(IP);
                          
                            //if (PR.Status.ToString().Equals("Success"))
                            //{
                            //    string ser = "";
                            //    string bar = "";
                            //    string sqlmix = "SELECT * FROM IF_MixPrintLab where Plan_ID ='" + Plan_Id + "'";
                            //    DataTable sqlmix1 = sqlinsertScan.ExecuteQuery(sqlmix, IP.Trim());
                            //    br = Plan_Id + (int.Parse(sqlmix1.Rows.Count.ToString()) + 1).ToString("000").Trim();
                            //    bar = br.Substring(3).Trim();
                            //    string serialnum = "SELECT count(*) FROM IF_MixPrintLab where Convert(varchar(10), Write_Time,25) ='" + DateTime.Now.ToString("yyyy-MM-dd") + "' and  Equip_ID='" + Machno.Substring(1, 1) + "'";
                            //    DataTable serial = sqlinsertScan.ExecuteQuery(serialnum, IP.Trim());
                            //    if (serial.Rows.Count > 0)
                            //        ser = (int.Parse(serial.Rows[0][0].ToString()) + 1).ToString().Trim();
                            //    string sqlprintlab = "INSERT INTO [IF_MixPrintLab] ([MixSaveTime],[Shift],[Barcode],[Equip_ID],[Plan_ID],[Recipe_Code],[Recipe_Name],[Recipe_Type],[Set_Num],[Serial_Num],[ShiftNum],[Dailylimit],[Weight],[BarCodeLab],[Write_Time],[Read_Time],[RW_Flag],[Active])" +
                            //        " VALUES('" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "','" + classs + "','" + bar + "','" + Machno + "','" + Plan_Id + "','" + partno + "' " +
                            //        " ,'" + partno + "','" + ptype + "','" + br1 + "','" + ser + "','2','" + daylimt + "'" +
                            //        " ,'" + Soluong + "','" + barcode + "','" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "','" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "','1','0')";
                            //    bool printlab = sqlinsertScan.ExecuteNonQuery(sqlprintlab, IP.Trim());
                            //}
                        }
                        catch { continue; }

                    }
                }
                //------------------------------------------------------------------//
                result = "";
                result = Get_Excel_BB.Create_Excel(partno, effdat, pathfile, makeo.Trim(), slipno, barcode, ca, daylimt, Soluong, mesid, Machno, pday, indat, classs, intime, pallet, OEM,somesx);
                if (result != "")
                {
                    
                    result = "Kẹt lệnh in Excel 9.245";
                    return result;
                }

                result = Print_Excel(filename, pathfile);


            }

            return result;
        }
        [WebMethod]
        public string Print_BBbutem(string tenkeo, string Machno, string printername, string ca, string usrno, string candao, string Soluong, string mesid, string pallet, string shift_id, string pday, string indat, string intime, string OEM)
        {


            string filename = "_" + printername + ".xlsx";
            string pathfolder = System.AppDomain.CurrentDomain.BaseDirectory + @"Data\";
            string pathfile = System.AppDomain.CurrentDomain.BaseDirectory + @"Data\" + filename; //tạo file và folder cho file Excel
            if (!Directory.Exists(pathfolder)) // Nếu chưa có folder thì tạo cái mới
            {
                Directory.CreateDirectory(pathfolder);
            }
            DateTime myPday = DateTime.ParseExact(pday, "yyyyMMdd", CultureInfo.InvariantCulture);
            pday = myPday.ToString("yyyyMMdd");
            string description = "";

            //string sqlRecipe = "select mater_name from pmt_recipe";
            string Plan_Id = "";
            string result = "";
            string partno = tenkeo.Trim();
            string classs = shift_id;
            string slipno = classs + Machno + "-" + pday.Substring(4, 4);
            string barcode = "";
            string effdat = "";
            string daylimt = "";
            string ptype = "";
            string machno = "V-BB37" + Machno;
            string somesx = "";
            ca = shift_id;
            string makeo = "";
            effdat = myPday.AddDays(7).ToString("yyyyMMdd");
            string syear = myPday.ToString("yyyy").Substring(2, 2);
            string sday = myPday.ToString("dd");
            string smonth = myPday.ToString("MM");
            switch (smonth)
            {
                case "10": smonth = "A"; break;
                case "11": smonth = "B"; break;
                case "12": smonth = "C"; break;
                default: smonth = smonth.Substring(1, 1); break;
            }
            string spday = syear + smonth + sday;


            string sqlRTPLAN = "select Plan_Id, P_Date from [mfnsShareDB].[dbo].[IF_RtPlan2Mixing] where  Plan_Id ='" + mesid + "'";
            DataTable dtKEORE = new DataTable();
            dtKEORE = SQLCNN8MAY.ExecuteQuery(sqlRTPLAN, Machno);
            if (dtKEORE.Rows.Count > 0)
            {
                if (dtKEORE.Rows[0][1].ToString().Trim().Replace("-", "") != pday.Trim())
                {
                    string s = "Mes quá giờ không quét được";
                    return s;
                }
                if (tenkeo.Trim().Length <= 5)
                {
                    makeo = "RD"; ptype = "3";
                }
                else
                {
                    switch (tenkeo.Trim().Substring(6).ToUpper())
                    {
                        case "RM": makeo = "RD"; ptype = "3"; break;
                        case "1": makeo = "RB"; ptype = "2"; break;
                        case "9": makeo = "RD"; ptype = "3"; break;
                        case "2": makeo = "RC"; ptype = "2"; break;
                        case "3": makeo = "RC"; ptype = "2"; break;
                        case "4": makeo = "RC"; ptype = "2"; break;
                        case "5": makeo = "RC"; ptype = "2"; break;
                        case "RE": makeo = "RR"; ptype = "3"; break;
                        case "R": makeo = "RR"; ptype = "3"; break;
                        case "92": makeo = "RD"; ptype = "3"; break;
                        case "": makeo = "RD"; ptype = "3"; break;

                        case "1-EDGE": makeo = "RB"; ptype = "2"; break;
                        case "2-EDGE": makeo = "RC"; ptype = "2"; break;
                        case "3-EDGE": makeo = "RC"; ptype = "2"; break;
                        case "4-EDGE": makeo = "RC"; ptype = "2"; break;
                        case "5-EDGE": makeo = "RC"; ptype = "2"; break;
                        case "9-EDGE": makeo = "RD"; ptype = "3"; break;
                        case "EDGE": makeo = "RD"; ptype = "3"; break;

                        case "1EDGE": tenkeo = partno.Substring(0, 5) + "-1-EDGE"; makeo = "RB"; ptype = "2"; break;
                        case "2EDGE": tenkeo = partno.Substring(0, 5) + "-2-EDGE"; makeo = "RC"; ptype = "2"; break;
                        case "3EDGE": tenkeo = partno.Substring(0, 5) + "-3-EDGE"; makeo = "RC"; ptype = "2"; break;
                        case "4EDGE": tenkeo = partno.Substring(0, 5) + "-4-EDGE"; makeo = "RC"; ptype = "2"; break;
                        case "5EDGE": tenkeo = partno.Substring(0, 5) + "-5-EDGE"; makeo = "RC"; ptype = "2"; break;
                        case "9EDGE": tenkeo = partno.Substring(0, 5) + "-9-EDGE"; makeo = "RD"; ptype = "3"; break;

                        case "1THU": tenkeo = partno.Substring(0, 5) + "-1THU"; makeo = "RB"; ptype = "2"; break;
                        case "2THU": tenkeo = partno.Substring(0, 5) + "-2THU"; makeo = "RC"; ptype = "2"; break;
                        case "3THU": tenkeo = partno.Substring(0, 5) + "-3THU"; makeo = "RC"; ptype = "2"; break;
                        case "4THU": tenkeo = partno.Substring(0, 5) + "-4THU"; makeo = "RC"; ptype = "2"; break;
                        case "5THU": tenkeo = partno.Substring(0, 5) + "-5THU"; makeo = "RC"; ptype = "2"; break;
                        case "9THU": tenkeo = partno.Substring(0, 5) + "-9THU"; makeo = "RD"; ptype = "3"; break;
                    }
                }



                string ktptype = "  SELECT [ptype], LTRIM(RTRIM([rubno_7])) rubno_7 FROM[InTem].[dbo].[rubnod_Ptype] WHERE SUBSTRING(rubno_7,7,1)= '2' AND rubno_7 = '" + partno.Trim() + "'";
                DataTable KT_ptype = sql_KEORE.ExecuteQuery(ktptype);
                if (KT_ptype.Rows.Count >= 2)
                {
                    result = "Liên hệ phòng thí nghiệm (a Thuần) đóng 1 tiêu chuẩn";
                    return result;
                }
                else
                {
                    if (KT_ptype.Rows.Count > 0)
                    {
                        makeo = "RB";
                        ptype = KT_ptype.Rows[0][0].ToString().Trim();

                    }

                }

                // Kiểm tra hạn sử dụng
                string getKeo = "select expday from [erp].[dbo].[prdexp] where subno='4' and factory='V' and  ptype ='" + makeo + "' and rubno='" + partno.Substring(0, 5) + "'";
                DataTable keo = SQL10_33.ExecuteQuery(getKeo);

                if (keo.Rows.Count == 0)
                {
                    result = "Mã keo không được sử dụng.\n Liên hệ Duyên phòng chế tạo (755) !";
                    return result;
                }
                else
                {
                    int day = int.Parse(keo.Rows[0][0].ToString().Trim());

                    //day = 7;
                    daylimt = day.ToString().Trim();
                    effdat = DateTime.Now.AddDays(day).ToString("yyyyMMdd");
                }



                string sqlBarcode = "SELECT MAX(SUBSTRING(Barcode,8,3)) FROM [erp].[dbo].[prdebe] where subno='4' and factory='V' and barcode like '%" + makeo + "%' and prodat = '" + pday + "'";
                DataTable dtBar = SQL10_33.ExecuteQuery(sqlBarcode);


                if (dtBar.Rows.Count == 1 && dtBar.Rows[0][0].ToString().Trim() == "")
                    barcode = makeo + spday + "001";
                else
                    barcode = makeo + spday + (int.Parse(dtBar.Rows[0][0].ToString()) + 1).ToString("000");
            }
            else
            {
                result = "Mã MES đã bị đóng! Liên hệ IT mở!";
                return result;
            }

            string idGrouplot = "";
            string br1 = "";

            #region kiem tra gioi han keo

            if (float.Parse(Soluong) < 30)
            {
                return "Lỗi! Trọng lượng không phù hợp!";
            }

            string ktKeo2 = "select * from [InTem].[dbo].[rubnod_Ptype] where [rubno_7] ='" + tenkeo.Trim() + "'";
            DataTable kiemkeo2 = sql_KEORE.ExecuteQuery(ktKeo2);



            string plannum = "SELECT Plan_Num FROM [mfnsShareDB].[dbo].[IF_RtPlan2Mixing] where Plan_Id ='" + mesid + "'";
            DataTable plann = SQLCNN8MAY.ExecuteQuery(plannum, Machno);
            if (plann.Rows.Count > 0)
            {
                br1 = plann.Rows[0][0].ToString().Trim();
            }

            string checkmes = " Select  Plan_Id,Plan_Num,Recipe_Code from [mfnsShareDB].[dbo].[IF_RtPlan2Mixing] where  replace(Recipe_Name,'-','') = '" + partno.Replace("-", "") + "' and Plan_Id ='" + mesid + "' ";
            DataTable dtMEsid = SQLCNN8MAY.ExecuteQuery(checkmes, Machno);

            if (dtMEsid.Rows.Count == 0)
            {
                result = "MES không tồn tại  IF_RtPlan2Mixing , tạo MES khác!";
                return result;
            }

            Plan_Id = dtMEsid.Rows[0][0].ToString().Trim();

            string recipe_name = dtMEsid.Rows[0][2].ToString().Trim();
            string planqty = dtMEsid.Rows[0][1].ToString().Trim();


            string DateNow = DateTime.Now.ToString("yyyy-MM-dd 06:30:00").ToString();
            string timeH = DateTime.Now.ToString("HHmm");
            if (int.Parse(timeH) <= 630)
            {
                DateNow = DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd 18:30:00").ToString();
            }
            //2024/10/01
            //string ktMESplan = " select b.FinishNum * (select sum(set_weight) from[mfns].[dbo].[pmt_weigh] where father_code = b.RecipeName) from [mfns].[dbo].[Ppt_GroupLot] b " +
            //    " where MesPlanID = '" + Plan_Id + "' and RecipeName = '" + recipe_name + "' and SetNumber = '" + planqty + "' and End_datetime is not null";

            string ktMESplan = " select b.FinishNum * (select sum(set_weight) from[mfns].[dbo].[pmt_weigh] where father_code = b.RecipeName) from [mfns].[dbo].[Ppt_GroupLot] b " +
               " where MesPlanID = '" + Plan_Id + "' and RecipeName = '" + recipe_name + "'  and End_datetime is not null";

            DataTable ktplan = SQLCNN8MAY.ExecuteQuery(ktMESplan, Machno);


            if (makeo.Substring(0, 1) == "R")
            {
                string LayKeoQuetTem = " select ISNULL(sum([weight]),0) from [erp].[dbo].[prdebe] where subno='4' and factory='V' and machno ='" + machno + "' and mesid ='" + Plan_Id + "' ";
                DataTable getWeightSX = SQL10_33.ExecuteQuery(LayKeoQuetTem);

                float GioiHanKeo = float.Parse(ktplan.Rows[0][0].ToString().Trim());
                float KeoSX = float.Parse(getWeightSX.Rows[0][0].ToString());
                float KeoVo = KeoSX + float.Parse(Soluong);

                if (KeoVo > GioiHanKeo)
                {
                    if (GioiHanKeo < KeoSX)
                    {
                        result = "MES này quá số lượng kế hoạch, không thể quét tiếp!";
                        return result;
                    }
                    else
                    {
                        result = "Lỗi! MES này chỉ quét được " + (GioiHanKeo - KeoSX).ToString().Trim() + "KG nữa!";
                        return result;
                    }
                }
            }
            #endregion kiem tra gioi han keo



            if (pallet.Length < 6)
            {
                result = "Pallet không đủ ký tự, nhập lại!!!";
                return result;
            }

            if (pallet.Substring(0, 2) != "VB" && pallet.Substring(0, 2) != "EB" && pallet.Substring(0, 2) != "VC" && pallet.Substring(0, 2) != "VD" && pallet.Substring(0, 2) != "VE" && pallet.Substring(0, 2) != "VF" && pallet.Substring(0, 1) != "V")
            {
                result = "Pallet Không hợp lệ, nhập lại!!!";
                return result;
            }
            string SelPallet = "SELECT PALLET_NO FROM [InTem].[dbo].[PalletBB] WHERE PALLET_NO='" + pallet + "'";
            DataTable dtPallet = sql_KEORE.ExecuteQuery(SelPallet);

            if (dtPallet.Rows.Count == 0)
            {
                string sDat = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                string sqlIns = "INSERT INTO [InTem].[dbo].[PalletBB] VALUES('" + pallet + "', '" + sDat + "', '" + usrno + "','1', 'Y') ";

                bool bQuery = sql_KEORE.ExecuteNonQuery(sqlIns);
                if (!bQuery)
                {
                    result = "Lỗi cập nhật Pallet\nPalletBB";
                    return result;
                }
            }


            string ktXuatPallet = "select top 1 active from [dbo].[prdebe] where subno ='4' and factory='V' and pallet_no='" + pallet.Trim() + "' order by indat desc, intime desc";
            DataTable kttpallet = SQL10_33.ExecuteQuery(ktXuatPallet);
            if (kttpallet.Rows.Count > 0)
            {
                if (kttpallet.Rows[0][0].ToString() == "N")
                {
                    result = "Pallet này chưa xuất, không được trùng pallet";
                    return result;
                }
            }
            string ktTrungPallet = "select * from [dbo].[prdebe] where subno ='4' and factory='V' and mesid='" + mesid.Trim() + "' and pallet_no='" + pallet.Trim() + "' ";
            DataTable kttpallet1 = SQL10_33.ExecuteQuery(ktTrungPallet);
            if (kttpallet1.Rows.Count > 0)
            {
                result = "1 Pallet chỉ được quét 1 lần cho 1 mã MES";
                return result;
            }


            string active = "Y";
            if (ptype == "3")
            {
                active = "N";
            }
            if (makeo == "RB")
            {
                candao = "N";
            }

            if (ptype == "2")
            {
                active = "";
            }
            if (OEM != "")
            {
                string insertTemOEM = "INSERT INTO [BB].[dbo].[TemOEMBB] ([mesid], [partno], [Barcode], [indat], [intime]) VALUES('" + Plan_Id + "', '" + partno + "', '" + barcode + "', '" + indat + "', '" + intime + "'); ";
                bool insertTemOEM2 = SQL10_33.ExecuteNonQuery(insertTemOEM);

            }


            string weightRecipe = "SELECT SUM(set_weight) as weightRecipe FROM [mfns].[dbo].[pmt_weigh] where father_code ='" + partno + "'";
            DataTable weightRC = SQLCNN8MAY.ExecuteQuery(weightRecipe, machno);
            if (weightRC.Rows[0][0].ToString().Trim() == "")
            {
                result = "Không tìm thấy số kg tiêu chuẩn";
                return result;
            }
            string weight = "";
            weight = weightRC.Rows[0][0].ToString().Trim();

            double weightValue = double.Parse(weight) * 1.02;

            //thêm 2% so với tc
            string ktSotemin = "SELECT *  FROM [erp].[dbo].[prdebe] where mesid ='" + Plan_Id + "' order by indat desc, intime desc";
            DataTable dtSkgTemin = SQLCNN8MAY.ExecuteQuery(ktSotemin, "33");
            if (dtSkgTemin.Rows.Count == 0)
            {
                if (double.Parse(Soluong) > weightValue)
                {
                    somesx = "1-2";

                }
                if (double.Parse(Soluong) <= weightValue)
                {
                    somesx = "1";
                }
            }
            if (dtSkgTemin.Rows.Count > 0)
            {
                string someLonnhat = dtSkgTemin.Rows[0]["some_sx"].ToString().Trim();
                string lastCharStr = someLonnhat.Split('-').Last();
                if (double.Parse(Soluong) > weightValue)
                {
                    int meA = int.Parse(lastCharStr) + 1;
                    int meB = int.Parse(lastCharStr) + 2;

                    somesx = meA.ToString() + "-" + meB.ToString();


                }
                if (double.Parse(Soluong) <= weightValue)
                {
                    int meA = int.Parse(lastCharStr) + 1;
                    somesx = meA.ToString();
                }

            }




            string machnoinsert = machno;

            string sql1 = "INSERT INTO [dbo].[prdebe] VALUES ('4' , 'V', '" + Plan_Id + "','" + machnoinsert + "', '" + daylimt + "','" + barcode + "','" + slipno + "','" + Soluong + "', '" + pday + "',  '" + effdat + "', " +
            "'" + classs + "', '" + ptype + "','" + candao.Trim() + "','" + partno + "', '" + intime + "', '" + indat + "', '" + usrno + "','" + pallet + "','" + active + "','"+somesx+"')";
            bool b1 = SQL10_33.ExecuteNonQuery(sql1);

            //bool b1 = true;

            if (!b1)
            {
                result = "Đợi 1 chút rồi quét lại!!!";
                return result;
            }
            else
            {


                string sqlip = "SELECT ip FROM [erp].[dbo].[P8500_IP] where machno like 'V-BB%' ";
                DataTable ip = SQL10_33.ExecuteQuery(sqlip);
                string IP = "";

                if (ip.Rows.Count > 0)
                {
                    foreach (DataRow item in ip.Rows)
                    {
                        try
                        {
                            IP = item["ip"].ToString().Trim();
                            string br = "";

                            PingReply PR = ping.Send(IP);

                            //if (PR.Status.ToString().Equals("Success"))
                            //{
                            //    string ser = "";
                            //    string bar = "";
                            //    string sqlmix = "SELECT * FROM IF_MixPrintLab where Plan_ID ='" + Plan_Id + "'";
                            //    DataTable sqlmix1 = sqlinsertScan.ExecuteQuery(sqlmix, IP.Trim());
                            //    br = Plan_Id + (int.Parse(sqlmix1.Rows.Count.ToString()) + 1).ToString("000").Trim();
                            //    bar = br.Substring(3).Trim();
                            //    string serialnum = "SELECT count(*) FROM IF_MixPrintLab where Convert(varchar(10), Write_Time,25) ='" + DateTime.Now.ToString("yyyy-MM-dd") + "' and  Equip_ID='" + Machno.Substring(1, 1) + "'";
                            //    DataTable serial = sqlinsertScan.ExecuteQuery(serialnum, IP.Trim());
                            //    if (serial.Rows.Count > 0)
                            //        ser = (int.Parse(serial.Rows[0][0].ToString()) + 1).ToString().Trim();
                            //    string sqlprintlab = "INSERT INTO [IF_MixPrintLab] ([MixSaveTime],[Shift],[Barcode],[Equip_ID],[Plan_ID],[Recipe_Code],[Recipe_Name],[Recipe_Type],[Set_Num],[Serial_Num],[ShiftNum],[Dailylimit],[Weight],[BarCodeLab],[Write_Time],[Read_Time],[RW_Flag],[Active])" +
                            //        " VALUES('" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "','" + classs + "','" + bar + "','" + Machno + "','" + Plan_Id + "','" + partno + "' " +
                            //        " ,'" + partno + "','" + ptype + "','" + br1 + "','" + ser + "','2','" + daylimt + "'" +
                            //        " ,'" + Soluong + "','" + barcode + "','" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "','" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "','1','0')";
                            //    bool printlab = sqlinsertScan.ExecuteNonQuery(sqlprintlab, IP.Trim());
                            //}
                        }
                        catch { continue; }

                    }
                }
                //------------------------------------------------------------------//
                result = "";
                result = Get_Excel_BB.Create_Excel(partno, effdat, pathfile, makeo.Trim(), slipno, barcode, ca, daylimt, Soluong, mesid, Machno, pday, indat, classs, intime, pallet, OEM,somesx);
                if (result != "")
                {

                    result = "Kẹt lệnh in Excel 9.245";
                    return result;
                }

                result = Print_Excel(filename, pathfile);


            }

            return result;
        }


        [WebMethod]
        public string Print_BB_Again(string usrno, string Machno, string barcode, string mesid, string soluong, string printername, string tenkeo)
        {

            //- android bị truyền sai tên máy in nên sửa kiểu này
            string[] mayin = Printer();
            string matchingPrinter = Array.Find(mayin, printer => printer.Contains(printername));
            if (!string.IsNullOrEmpty(matchingPrinter))
            {
                printername = matchingPrinter;
            }
            //- android bị truyền sai tên máy in nên sửa kiểu này


            string filename = "_" + printername + ".xlsx";
            string pathfolder = System.AppDomain.CurrentDomain.BaseDirectory + @"Data\";
            string pathfile = System.AppDomain.CurrentDomain.BaseDirectory + @"Data\" + filename; //tạo file và folder cho file Excel
            if (!Directory.Exists(pathfolder)) // Nếu chưa có folder thì tạo cái mới
            {
                Directory.CreateDirectory(pathfolder);
            }
            string result = "";
            string description = "";

            string checkmes = " select Plan_Num from [mfnsShareDB].[dbo].[IF_RtPlan2Mixing] where  Plan_Id ='"+mesid+"'";
            DataTable dtMEsid = SQLCNN8MAY.ExecuteQuery(checkmes,Machno);




            if (dtMEsid.Rows.Count == 0)
            {
                result = "MES không tồn tại [IF_RtPlan2Mixing], tạo MES khác!";
                return result;
            }
            string checkoem = "SELECT [Barcode]  FROM [BB].[dbo].[TemOEMBB] where Barcode ='"+barcode+"'";
            DataTable dtoem = SQL10_33.ExecuteQuery(checkoem);
            if (dtoem.Rows.Count != 0)
            {
                description = "OEM";
            }

         

            string sql = "Select * from prdebe where subno='4' and factory='V' and machno ='" + "V-BB37" + Machno + "'  and mesid = '" + mesid + "' and barcode = '" + barcode.Trim() + "'";
            DataTable dtsql = SQL10_33.ExecuteQuery(sql);

            if (dtsql.Rows.Count == 1)
            {
                string loaikeo = dtsql.Rows[0]["barcode"].ToString().Trim().Substring(0, 2);
                string pallet = dtsql.Rows[0]["pallet_no"].ToString().Trim();
                string indat = dtsql.Rows[0]["indat"].ToString().Trim();
                string effdat = dtsql.Rows[0]["effdat"].ToString().Trim();
                string intime = dtsql.Rows[0]["intime"].ToString().Trim();
                string partno = dtsql.Rows[0]["partno"].ToString().Trim();
                string slipno = dtsql.Rows[0]["slipno"].ToString().Trim();
                string daylimt = dtsql.Rows[0]["daylimt"].ToString().Trim();
                string pday = dtsql.Rows[0]["prodat"].ToString().Trim();
                string soluong1 = dtsql.Rows[0]["weight"].ToString().Trim();
                string classs = dtsql.Rows[0]["class"].ToString().Trim();
                string ca = slipno.Substring(0, 1);
                string someSX = dtsql.Rows[0]["some_sx"].ToString().Trim();

                result = "";
                result = Get_Excel_InlaiBB.Create_Excel(partno, effdat, pathfile, loaikeo, slipno, barcode.Trim(), ca, daylimt, soluong1, mesid, Machno, pday, indat, classs, intime, pallet, description,someSX);
                result = Print_Excel(filename, pathfile);
                string print_date = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss").Trim();
                string print_again_log = "Insert into [BB].[dbo].[Print_again_log] Values ('" + mesid + "','" + barcode.Trim() + "','" + slipno + "','" + partno + "','" + soluong1 + "','" + pallet + "','" + print_date + "','" + usrno + "',N'Hiện trường tự in, Device: Android ')";
                bool b_check = SQL10_33.ExecuteNonQuery(print_again_log);
                if (result != "")
                {
                    return result;
                }
            }
            else
            {
                result = "Không có dữ liệu TEM này!";
            }
            return result;
        }
        private string Print_Excel(string filename, string pathfile)
        {
            try
            {
                bool flag = true;
                string messager = "";
                string[] sp = filename.Split('_');
                string printerName = sp[1].Substring(0, sp[1].Length - 5);

                if (printerName.Substring(0, 3) == "Fax" || printerName.Substring(0, 5) == "Foxit"
                    || printerName.Substring(0, 9) == "Microsoft")
                {
                    File.Delete(pathfile);
                    return "Chon lai may in";
                }

                PrintExcel.Print_xls_file(printerName, pathfile, ref flag, ref messager); //In
                if (flag == true)
                {
                    File.Delete(pathfile);
                }
                else
                {
                    File.Delete(pathfile);
                    return sp[0];
                }
                return "";
            }
            catch (Exception ex)
            {
                return "Lỗi kẹt lệnh in 9.245!";
            }
        }
    }
   
}
