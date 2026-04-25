using ServiceBBAPI.Helpers;
using ServiceBBAPI.Models;
using ServiceBBAPI.SQLCNN;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Web.Services;

namespace ServiceBBAPI
{
    [WebService(Namespace = "http://tempuri.org/")]
    [WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
    [System.ComponentModel.ToolboxItem(false)]
    public class ServiceBBintem : WebService
    {
        private readonly DatabaseHelper _erpDb = DatabaseConnections.GetErpDb();
        private readonly DatabaseHelper _erp34Db = DatabaseConnections.GetErp34Db();
        private readonly DatabaseHelper _inTemDb = DatabaseConnections.GetInTemDb();

        private DatabaseHelper GetMachineDb(string machineKey)
        {
            return DatabaseConnections.GetMachineDb(machineKey);
        }

        [WebMethod]
        public string HelloWorld()
        {
            return "Hello World";
        }

        [WebMethod]
        public string[] Login(string empno)
        {
            string[] result = new string[4];

            DataTable dt = _erp34Db.ExecuteQuery(
                "SELECT name, depno, empno FROM peremp WHERE empno = @empno AND subno = '4'",
                new SqlParameter("@empno", empno));

            if (dt.Rows.Count == 0)
            {
                result[0] = "Số thẻ không tồn tại";
                return result;
            }

            result[0] = "";
            result[1] = dt.Rows[0]["name"].ToString().Trim();
            result[2] = dt.Rows[0]["depno"].ToString().Trim();
            result[3] = dt.Rows[0]["empno"].ToString().Trim();
            return result;
        }

        [WebMethod]
        public string[] may()
        {
            return new[] { "01", "02", "03", "04", "05", "06", "07", "08" };
        }

        [WebMethod]
        public string Get_shift()
        {
            return ShiftHelper.GetShiftId();
        }

        [WebMethod]
        public string Get_shiftMayBB()
        {
            return ShiftHelper.GetShiftMayBB();
        }

        [WebMethod]
        public DataSet Get_Mesid_list(string Machno)
        {
            string shiftId = ShiftHelper.GetShiftMayBB();
            string pday = ShiftHelper.GetPdayBB(shiftId);
            DateTime myPday = DateTime.ParseExact(pday, "yyyyMMdd", CultureInfo.InvariantCulture);
            string pdayFormatted = myPday.ToString("yyyy-MM-dd");

            var machineDb = GetMachineDb(Machno);
            DataTable dtMes = machineDb.ExecuteQuery(
                @"SELECT a.Plan_Id, a.Recipe_Code 
                  FROM [mfnsShareDB].[dbo].[IF_RtPlan2Mixing] a
                  INNER JOIN [mfns].[dbo].[Ppt_GroupLot] b ON a.Plan_Id = b.MesPlanID
                  WHERE a.Shift_Id = @shiftId 
                    AND a.P_Date = @pday 
                    AND a.Plan_Id NOT LIKE 'V%' 
                    AND b.End_datetime IS NOT NULL",
                new SqlParameter("@shiftId", shiftId),
                new SqlParameter("@pday", pdayFormatted));

            DataSet ds = new DataSet();
            ds.Tables.Add(dtMes);
            return ds;
        }

        [WebMethod]
        public string[] Printer()
        {
            DataTable dtPrinter = _erpDb.ExecuteQuery(
                "SELECT [TenMay], [MaMay] FROM [BB].[dbo].[Printer_BB] WHERE IP = 'BB'");

            return dtPrinter.AsEnumerable()
                .Select(row => row[1].ToString().Trim())
                .ToArray();
        }

        [WebMethod]
        public string Finish(string tenkeo, string mesid, string machno)
        {
            return "";
        }

        [WebMethod]
        public DataSet Get_Barcode(string mesid, string machno)
        {
            string fullMachno = "V-BB37" + machno;
            DataTable dtBar = _erpDb.ExecuteQuery(
                "SELECT barcode, weight FROM prdebe WHERE subno = '4' AND factory = 'V' AND mesid = @mesid AND machno = @machno",
                new SqlParameter("@mesid", mesid),
                new SqlParameter("@machno", fullMachno));

            DataSet ds = new DataSet();
            ds.Tables.Add(dtBar);
            return ds;
        }

        [WebMethod]
        public DataSet Get_Mesid_list_Inlai(string Machno)
        {
            string shiftId = ShiftHelper.GetShiftMayBB();
            string pday = ShiftHelper.GetPdayBB(shiftId);
            DateTime myPday = DateTime.ParseExact(pday, "yyyyMMdd", CultureInfo.InvariantCulture);
            string pdayFormatted = myPday.ToString("yyyy-MM-dd");

            var machineDb = GetMachineDb(Machno);
            DataTable dtMes = machineDb.ExecuteQuery(
                @"SELECT a.Plan_Id, a.Recipe_Code 
                  FROM [mfnsShareDB].[dbo].[IF_RtPlan2Mixing] a
                  INNER JOIN [mfns].[dbo].[Ppt_GroupLot] b ON a.Plan_Id = b.MesPlanID
                  WHERE a.Shift_Id = @shiftId 
                    AND a.P_Date = @pday 
                    AND b.End_datetime IS NOT NULL",
                new SqlParameter("@shiftId", shiftId),
                new SqlParameter("@pday", pdayFormatted));

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

        #region Print_BB

        [WebMethod]
        public string Print_BB(string tenkeo, string Machno, string printername, string ca, string usrno,
            string candao, string Soluong, string mesid, string pallet, string mac_pda, string OEM, string depno)
        {
            UpdatePdaMacAddress(mac_pda);

            string shiftId = ShiftHelper.GetShiftId();
            string pday = ShiftHelper.GetPday(shiftId);
            DateTime myPday = DateTime.ParseExact(pday, "yyyyMMdd", CultureInfo.InvariantCulture);
            pday = myPday.ToString("yyyyMMdd");
            string indat = DateTime.Now.ToString("yyyyMMdd");
            string intime = DateTime.Now.ToString("HH:mm:ss");
            string machno = "V-BB37" + Machno;
            string classs = shiftId;
            string slipno = classs + Machno + "-" + pday.Substring(4, 4);
            string spday = ShiftHelper.BuildSpday(myPday);
            ca = shiftId;

            var machineDb = GetMachineDb(Machno);

            string validationError = ValidateMesPlan(mesid, pday, machineDb);
            if (validationError != null) return validationError;

            RubberTypeResult rubberType = RubberTypeMapper.Resolve(tenkeo);
            string makeo = rubberType.Makeo;
            string ptype = rubberType.Ptype;
            tenkeo = rubberType.NormalizedTenkeo;
            string partno = tenkeo.Trim();

            string ptypeError = CheckAndOverridePtype(partno, ref makeo, ref ptype);
            if (ptypeError != null) return ptypeError;

            string expiryError = GetExpiryInfo(partno, makeo, out string daylimt, out string effdat);
            if (expiryError != null) return expiryError;

            string barcode = GenerateBarcode(makeo, spday, pday);

            if (!float.TryParse(Soluong, out float soluongValue) || soluongValue < 30)
                return "Lỗi! Trọng lượng không phù hợp!";

            string planNum = GetPlanNum(mesid, machineDb);

            string mesValidation = ValidateMesExists(partno, mesid, machineDb, out string planId, out string recipeName);
            if (mesValidation != null) return mesValidation;

            float keoSX = 0;
            float gioiHanKeo = 0;
            string limitError = CheckKeoLimit(planId, recipeName, makeo, machno, Soluong, machineDb, out keoSX, out gioiHanKeo);
            if (limitError != null) return limitError;

            string palletError = ValidateAndInsertPallet(pallet, usrno, mesid);
            if (palletError != null) return palletError;

            string active = DetermineActiveStatus(ptype, makeo);
            if (makeo == "RB") candao = "N";

            if (!string.IsNullOrEmpty(OEM))
            {
                InsertTemOEM(planId, partno, barcode, indat, intime);
            }

            string weightError = GetWeightRecipe(partno, machno, machineDb, out double kgTieuChuan);
            if (weightError != null) return weightError;

            string somesx = BatchCalculator.Calculate(kgTieuChuan, double.Parse(Soluong), keoSX, gioiHanKeo);

            bool insertSuccess = InsertPrdebe(planId, machno, daylimt, barcode, slipno, Soluong, pday,
                effdat, classs, ptype, candao, partno, intime, indat, usrno, pallet, active, somesx);

            if (!insertSuccess)
                return "Đợi 1 chút rồi quét lại!!!";

            return PrintLabel(printername, partno, effdat, makeo, slipno, barcode, ca,
                daylimt, Soluong, mesid, Machno, pday, indat, classs, intime, pallet, OEM, somesx);
        }

        #endregion

        #region Print_BBbutem

        [WebMethod]
        public string Print_BBbutem(string tenkeo, string Machno, string printername, string ca, string usrno,
            string candao, string Soluong, string mesid, string pallet, string shift_id, string pday,
            string indat, string intime, string OEM)
        {
            DateTime myPday = DateTime.ParseExact(pday, "yyyyMMdd", CultureInfo.InvariantCulture);
            pday = myPday.ToString("yyyyMMdd");
            string machno = "V-BB37" + Machno;
            string classs = shift_id;
            string slipno = classs + Machno + "-" + pday.Substring(4, 4);
            string spday = ShiftHelper.BuildSpday(myPday);
            ca = shift_id;

            var machineDb = GetMachineDb(Machno);

            string validationError = ValidateMesPlan(mesid, pday, machineDb);
            if (validationError != null) return validationError;

            RubberTypeResult rubberType = RubberTypeMapper.Resolve(tenkeo);
            string makeo = rubberType.Makeo;
            string ptype = rubberType.Ptype;
            tenkeo = rubberType.NormalizedTenkeo;
            string partno = tenkeo.Trim();

            string ptypeError = CheckAndOverridePtype(partno, ref makeo, ref ptype);
            if (ptypeError != null) return ptypeError;

            string expiryError = GetExpiryInfo(partno, makeo, out string daylimt, out string effdat);
            if (expiryError != null) return expiryError;

            string barcode = GenerateBarcode(makeo, spday, pday);

            if (!float.TryParse(Soluong, out float soluongValue) || soluongValue < 30)
                return "Lỗi! Trọng lượng không phù hợp!";

            string planNum = GetPlanNum(mesid, machineDb);

            string mesValidation = ValidateMesExists(partno, mesid, machineDb, out string planId, out string recipeName);
            if (mesValidation != null) return mesValidation;

            if (makeo.StartsWith("R"))
            {
                string limitError = CheckKeoLimit(planId, recipeName, makeo, machno, Soluong, machineDb,
                    out float keoSXTemp, out float gioiHanTemp);
                if (limitError != null) return limitError;
            }

            string palletError = ValidateAndInsertPallet(pallet, usrno, mesid);
            if (palletError != null) return palletError;

            string active = DetermineActiveStatus(ptype, makeo);
            if (makeo == "RB") candao = "N";

            if (!string.IsNullOrEmpty(OEM))
            {
                InsertTemOEM(planId, partno, barcode, indat, intime);
            }

            string weightError = GetWeightRecipe(partno, machno, machineDb, out double kgTieuChuan);
            if (weightError != null) return weightError;

            double weightValue = kgTieuChuan * 1.02;
            string somesx = CalculateBatchOldMethod(planId, Soluong, weightValue);

            bool insertSuccess = InsertPrdebe(planId, machno, daylimt, barcode, slipno, Soluong, pday,
                effdat, classs, ptype, candao, partno, intime, indat, usrno, pallet, active, somesx);

            if (!insertSuccess)
                return "Đợi 1 chút rồi quét lại!!!";

            return PrintLabel(printername, partno, effdat, makeo, slipno, barcode, ca,
                daylimt, Soluong, mesid, Machno, pday, indat, classs, intime, pallet, OEM, somesx);
        }

        #endregion

        #region Print_BB_Again

        [WebMethod]
        public string Print_BB_Again(string usrno, string Machno, string barcode, string mesid,
            string soluong, string printername, string tenkeo)
        {
            string[] availablePrinters = Printer();
            string matchingPrinter = availablePrinters.FirstOrDefault(p => p.Contains(printername));
            if (!string.IsNullOrEmpty(matchingPrinter))
            {
                printername = matchingPrinter;
            }

            string filename = "_" + printername + ".xlsx";
            string pathfolder = AppDomain.CurrentDomain.BaseDirectory + @"Data\";
            string pathfile = pathfolder + filename;
            EnsureDirectoryExists(pathfolder);

            var machineDb = GetMachineDb(Machno);

            DataTable dtMes = machineDb.ExecuteQuery(
                "SELECT Plan_Num FROM [mfnsShareDB].[dbo].[IF_RtPlan2Mixing] WHERE Plan_Id = @mesid",
                new SqlParameter("@mesid", mesid));

            if (dtMes.Rows.Count == 0)
                return "MES không tồn tại [IF_RtPlan2Mixing], tạo MES khác!";

            string description = "";
            DataTable dtOem = _erpDb.ExecuteQuery(
                "SELECT [Barcode] FROM [BB].[dbo].[TemOEMBB] WHERE Barcode = @barcode",
                new SqlParameter("@barcode", barcode));
            if (dtOem.Rows.Count > 0)
                description = "OEM";

            string fullMachno = "V-BB37" + Machno;
            DataTable dtsql = _erpDb.ExecuteQuery(
                @"SELECT * FROM prdebe 
                  WHERE subno = '4' AND factory = 'V' 
                    AND machno = @machno AND mesid = @mesid AND barcode = @barcode",
                new SqlParameter("@machno", fullMachno),
                new SqlParameter("@mesid", mesid),
                new SqlParameter("@barcode", barcode.Trim()));

            if (dtsql.Rows.Count != 1)
                return "Không có dữ liệu TEM này!";

            DataRow row = dtsql.Rows[0];
            string loaikeo = row["barcode"].ToString().Trim().Substring(0, 2);
            string pallet = row["pallet_no"].ToString().Trim();
            string indat = row["indat"].ToString().Trim();
            string effdat = row["effdat"].ToString().Trim();
            string intime = row["intime"].ToString().Trim();
            string partno = row["partno"].ToString().Trim();
            string slipno = row["slipno"].ToString().Trim();
            string daylimt = row["daylimt"].ToString().Trim();
            string pday = row["prodat"].ToString().Trim();
            string soluong1 = row["weight"].ToString().Trim();
            string classs = row["class"].ToString().Trim();
            string ca = slipno.Substring(0, 1);
            string someSX = row["some_sx"].ToString().Trim();

            string result = Get_Excel_InlaiBB.Create_Excel(partno, effdat, pathfile, loaikeo, slipno,
                barcode.Trim(), ca, daylimt, soluong1, mesid, Machno, pday, indat, classs, intime,
                pallet, description, someSX);

            if (!string.IsNullOrEmpty(result))
                return result;

            result = Print_Excel(filename, pathfile);

            string printDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            _erpDb.ExecuteNonQuery(
                @"INSERT INTO [BB].[dbo].[Print_again_log] 
                  VALUES (@mesid, @barcode, @slipno, @partno, @soluong, @pallet, @printDate, @usrno, N'Hiện trường tự in, Device: Android ')",
                new SqlParameter("@mesid", mesid),
                new SqlParameter("@barcode", barcode.Trim()),
                new SqlParameter("@slipno", slipno),
                new SqlParameter("@partno", partno),
                new SqlParameter("@soluong", soluong1),
                new SqlParameter("@pallet", pallet),
                new SqlParameter("@printDate", printDate),
                new SqlParameter("@usrno", usrno));

            return result;
        }

        #endregion

        #region Private Helper Methods

        private void UpdatePdaMacAddress(string macPda)
        {
            macPda = macPda?.Trim() ?? "";
            if (string.IsNullOrEmpty(macPda)) return;

            string macDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            try
            {
                DataTable dtMac = _erpDb.ExecuteQuery(
                    "SELECT * FROM [QL PDA].[dbo].[KV2 PDA] WHERE MacAddress = @mac",
                    new SqlParameter("@mac", macPda));

                if (dtMac.Rows.Count > 0)
                {
                    _erpDb.ExecuteNonQuery(
                        "UPDATE [QL PDA].[dbo].[KV2 PDA] SET Intime = @intime WHERE MacAddress = @mac",
                        new SqlParameter("@intime", macDate),
                        new SqlParameter("@mac", macPda));
                }
                else
                {
                    string ip = GetClientIp();
                    _erpDb.ExecuteNonQuery(
                        "UPDATE [QL PDA].[dbo].[KV2 PDA] SET Intime = @intime, MacAddress = @mac WHERE IP = @ip",
                        new SqlParameter("@intime", macDate),
                        new SqlParameter("@mac", macPda),
                        new SqlParameter("@ip", ip?.Trim() ?? ""));
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[PDA MAC Update Error] {ex.Message}");
            }
        }

        private string ValidateMesPlan(string mesid, string pday, DatabaseHelper machineDb)
        {
            DataTable dtPlan = machineDb.ExecuteQuery(
                "SELECT Plan_Id, P_Date FROM [mfnsShareDB].[dbo].[IF_RtPlan2Mixing] WHERE Plan_Id = @planId",
                new SqlParameter("@planId", mesid));

            if (dtPlan.Rows.Count == 0)
                return "Mã MES đã bị đóng! Liên hệ IT mở!";

            string planDate = dtPlan.Rows[0]["P_Date"].ToString().Trim().Replace("-", "");
            if (planDate != pday.Trim())
                return "Mes quá giờ không quét được";

            return null;
        }

        private string CheckAndOverridePtype(string partno, ref string makeo, ref string ptype)
        {
            DataTable dtPtype = _inTemDb.ExecuteQuery(
                @"SELECT [ptype], LTRIM(RTRIM([rubno_7])) AS rubno_7 
                  FROM [InTem].[dbo].[rubnod_Ptype] 
                  WHERE SUBSTRING(rubno_7, 7, 1) = '2' AND rubno_7 = @partno",
                new SqlParameter("@partno", partno.Trim()));

            if (dtPtype.Rows.Count >= 2)
                return "Liên hệ phòng thí nghiệm (a Thuần) đóng 1 tiêu chuẩn";

            if (dtPtype.Rows.Count > 0)
            {
                makeo = "RB";
                ptype = dtPtype.Rows[0]["ptype"].ToString().Trim();
            }

            return null;
        }

        private string GetExpiryInfo(string partno, string makeo, out string daylimt, out string effdat)
        {
            daylimt = "";
            effdat = "";

            string rubno = partno.Length >= 5 ? partno.Substring(0, 5) : partno;
            DataTable dtExpiry = _erpDb.ExecuteQuery(
                @"SELECT expday FROM [erp].[dbo].[prdexp] 
                  WHERE subno = '4' AND factory = 'V' AND ptype = @ptype AND rubno = @rubno",
                new SqlParameter("@ptype", makeo),
                new SqlParameter("@rubno", rubno));

            if (dtExpiry.Rows.Count == 0)
                return "Mã keo không được sử dụng.\n Liên hệ Duyên phòng chế tạo (755) !";

            int days = int.Parse(dtExpiry.Rows[0]["expday"].ToString().Trim());
            daylimt = days.ToString();
            effdat = DateTime.Now.AddDays(days).ToString("yyyyMMdd");
            return null;
        }

        private string GenerateBarcode(string makeo, string spday, string pday)
        {
            DataTable dtBar = _erpDb.ExecuteQuery(
                @"SELECT MAX(SUBSTRING(Barcode, 8, 3)) 
                  FROM [erp].[dbo].[prdebe] 
                  WHERE subno = '4' AND factory = 'V' AND barcode LIKE @pattern AND prodat = @prodat",
                new SqlParameter("@pattern", "%" + makeo + "%"),
                new SqlParameter("@prodat", pday));

            if (dtBar.Rows.Count == 1 && string.IsNullOrEmpty(dtBar.Rows[0][0]?.ToString()?.Trim()))
                return makeo + spday + "001";

            int nextNumber = int.Parse(dtBar.Rows[0][0].ToString()) + 1;
            return makeo + spday + nextNumber.ToString("000");
        }

        private string GetPlanNum(string mesid, DatabaseHelper machineDb)
        {
            DataTable dtPlan = machineDb.ExecuteQuery(
                "SELECT Plan_Num FROM [mfnsShareDB].[dbo].[IF_RtPlan2Mixing] WHERE Plan_Id = @planId",
                new SqlParameter("@planId", mesid));

            return dtPlan.Rows.Count > 0 ? dtPlan.Rows[0][0].ToString().Trim() : "";
        }

        private string ValidateMesExists(string partno, string mesid, DatabaseHelper machineDb,
            out string planId, out string recipeName)
        {
            planId = "";
            recipeName = "";

            DataTable dtMes = machineDb.ExecuteQuery(
                @"SELECT Plan_Id, Plan_Num, Recipe_Code 
                  FROM [mfnsShareDB].[dbo].[IF_RtPlan2Mixing] 
                  WHERE REPLACE(Recipe_Name, '-', '') = @recipeName AND Plan_Id = @planId",
                new SqlParameter("@recipeName", partno.Replace("-", "")),
                new SqlParameter("@planId", mesid));

            if (dtMes.Rows.Count == 0)
                return "MES không tồn tại  IF_RtPlan2Mixing , tạo MES khác!";

            planId = dtMes.Rows[0]["Plan_Id"].ToString().Trim();
            recipeName = dtMes.Rows[0]["Recipe_Code"].ToString().Trim();
            return null;
        }

        private string CheckKeoLimit(string planId, string recipeName, string makeo, string machno,
            string soluong, DatabaseHelper machineDb, out float keoSX, out float gioiHanKeo)
        {
            keoSX = 0;
            gioiHanKeo = 0;

            if (!makeo.StartsWith("R"))
                return null;

            DataTable dtPlan = machineDb.ExecuteQuery(
                @"SELECT b.FinishNum * (SELECT SUM(set_weight) FROM [mfns].[dbo].[pmt_weigh] WHERE father_code = b.RecipeName) 
                  FROM [mfns].[dbo].[Ppt_GroupLot] b 
                  WHERE MesPlanID = @planId AND RecipeName = @recipeName AND End_datetime IS NOT NULL",
                new SqlParameter("@planId", planId),
                new SqlParameter("@recipeName", recipeName));

            if (dtPlan.Rows.Count == 0 || dtPlan.Rows[0][0] == DBNull.Value)
                return null;

            DataTable dtWeight = _erpDb.ExecuteQuery(
                @"SELECT ISNULL(SUM([weight]), 0) FROM [erp].[dbo].[prdebe] 
                  WHERE subno = '4' AND factory = 'V' AND machno = @machno AND mesid = @mesid",
                new SqlParameter("@machno", machno),
                new SqlParameter("@mesid", planId));

            gioiHanKeo = float.Parse(dtPlan.Rows[0][0].ToString().Trim());
            keoSX = float.Parse(dtWeight.Rows[0][0].ToString());
            float keoVo = keoSX + float.Parse(soluong);

            if (keoVo > gioiHanKeo)
            {
                if (gioiHanKeo < keoSX)
                    return "MES này quá số lượng kế hoạch, không thể quét tiếp!";

                return "Lỗi! MES này chỉ quét được " + (gioiHanKeo - keoSX).ToString().Trim() + "KG nữa!";
            }

            return null;
        }

        private string ValidateAndInsertPallet(string pallet, string usrno, string mesid)
        {
            string validationError = PalletValidator.Validate(pallet);
            if (validationError != null)
                return validationError;

            pallet = pallet.Trim();

            DataTable dtPallet = _inTemDb.ExecuteQuery(
                "SELECT PALLET_NO FROM [InTem].[dbo].[PalletBB] WHERE PALLET_NO = @pallet",
                new SqlParameter("@pallet", pallet));

            if (dtPallet.Rows.Count == 0)
            {
                string now = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                bool inserted = _inTemDb.ExecuteNonQuery(
                    "INSERT INTO [InTem].[dbo].[PalletBB] VALUES (@pallet, @dat, @usrno, '1', 'Y')",
                    new SqlParameter("@pallet", pallet),
                    new SqlParameter("@dat", now),
                    new SqlParameter("@usrno", usrno));

                if (!inserted)
                    return "Lỗi cập nhật Pallet\nPalletBB";
            }

            DataTable dtActive = _erpDb.ExecuteQuery(
                @"SELECT TOP 1 active FROM [dbo].[prdebe] 
                  WHERE subno = '4' AND factory = 'V' AND pallet_no = @pallet 
                  ORDER BY indat DESC, intime DESC",
                new SqlParameter("@pallet", pallet));

            if (dtActive.Rows.Count > 0 && dtActive.Rows[0][0].ToString() == "N")
                return "Pallet này chưa xuất, không được trùng pallet";

            DataTable dtDup = _erpDb.ExecuteQuery(
                @"SELECT * FROM [dbo].[prdebe] 
                  WHERE subno = '4' AND factory = 'V' AND mesid = @mesid AND pallet_no = @pallet",
                new SqlParameter("@mesid", mesid.Trim()),
                new SqlParameter("@pallet", pallet));

            if (dtDup.Rows.Count > 0)
                return "1 Pallet chỉ được quét 1 lần cho 1 mã MES";

            return null;
        }

        private string DetermineActiveStatus(string ptype, string makeo)
        {
            if (ptype == "3") return "N";
            if (ptype == "2") return "";
            return "Y";
        }

        private void InsertTemOEM(string planId, string partno, string barcode, string indat, string intime)
        {
            _erpDb.ExecuteNonQuery(
                @"INSERT INTO [BB].[dbo].[TemOEMBB] ([mesid], [partno], [Barcode], [indat], [intime]) 
                  VALUES (@mesid, @partno, @barcode, @indat, @intime)",
                new SqlParameter("@mesid", planId),
                new SqlParameter("@partno", partno),
                new SqlParameter("@barcode", barcode),
                new SqlParameter("@indat", indat),
                new SqlParameter("@intime", intime));
        }

        private string GetWeightRecipe(string partno, string machno, DatabaseHelper machineDb, out double kgTieuChuan)
        {
            kgTieuChuan = 0;

            DataTable dtWeight = machineDb.ExecuteQuery(
                "SELECT SUM(set_weight) AS weightRecipe FROM [mfns].[dbo].[pmt_weigh] WHERE father_code = @partno",
                new SqlParameter("@partno", partno));

            if (dtWeight.Rows.Count == 0 || dtWeight.Rows[0][0] == DBNull.Value ||
                string.IsNullOrEmpty(dtWeight.Rows[0][0].ToString().Trim()))
            {
                return "Không tìm thấy số kg tiêu chuẩn";
            }

            kgTieuChuan = double.Parse(dtWeight.Rows[0][0].ToString().Trim());
            return null;
        }

        private bool InsertPrdebe(string planId, string machno, string daylimt, string barcode,
            string slipno, string soluong, string pday, string effdat, string classs, string ptype,
            string candao, string partno, string intime, string indat, string usrno, string pallet,
            string active, string somesx)
        {
            return _erpDb.ExecuteNonQuery(
                @"INSERT INTO [dbo].[prdebe] VALUES 
                  ('4', 'V', @planId, @machno, @daylimt, @barcode, @slipno, @soluong, @pday, @effdat, 
                   @classs, @ptype, @candao, @partno, @intime, @indat, @usrno, @pallet, @active, @somesx)",
                new SqlParameter("@planId", planId),
                new SqlParameter("@machno", machno),
                new SqlParameter("@daylimt", daylimt),
                new SqlParameter("@barcode", barcode),
                new SqlParameter("@slipno", slipno),
                new SqlParameter("@soluong", soluong),
                new SqlParameter("@pday", pday),
                new SqlParameter("@effdat", effdat),
                new SqlParameter("@classs", classs),
                new SqlParameter("@ptype", ptype),
                new SqlParameter("@candao", candao.Trim()),
                new SqlParameter("@partno", partno),
                new SqlParameter("@intime", intime),
                new SqlParameter("@indat", indat),
                new SqlParameter("@usrno", usrno),
                new SqlParameter("@pallet", pallet),
                new SqlParameter("@active", active),
                new SqlParameter("@somesx", somesx));
        }

        private string PrintLabel(string printername, string partno, string effdat, string makeo,
            string slipno, string barcode, string ca, string daylimt, string soluong, string mesid,
            string machno, string pday, string indat, string classs, string intime, string pallet,
            string oem, string somesx)
        {
            string filename = "_" + printername + ".xlsx";
            string pathfolder = AppDomain.CurrentDomain.BaseDirectory + @"Data\";
            string pathfile = pathfolder + filename;
            EnsureDirectoryExists(pathfolder);

            string result = Get_Excel_BB.Create_Excel(partno, effdat, pathfile, makeo.Trim(), slipno,
                barcode, ca, daylimt, soluong, mesid, machno, pday, indat, classs, intime, pallet, oem, somesx);

            if (!string.IsNullOrEmpty(result))
                return "Kẹt lệnh in Excel 9.245";

            return Print_Excel(filename, pathfile);
        }

        private string CalculateBatchOldMethod(string planId, string soluong, double weightValue)
        {
            var erp33Db = DatabaseConnections.GetMachineDb("33");
            DataTable dtExisting = erp33Db.ExecuteQuery(
                "SELECT * FROM [erp].[dbo].[prdebe] WHERE mesid = @mesid ORDER BY indat DESC, intime DESC",
                new SqlParameter("@mesid", planId));

            double soluongValue = double.Parse(soluong);

            if (dtExisting.Rows.Count == 0)
            {
                return soluongValue > weightValue ? "1-2" : "1";
            }

            string lastSomesx = dtExisting.Rows[0]["some_sx"].ToString().Trim();
            string lastPart = lastSomesx.Split('-').Last();
            int lastNum = int.Parse(lastPart);

            if (soluongValue > weightValue)
            {
                return (lastNum + 1) + "-" + (lastNum + 2);
            }

            return (lastNum + 1).ToString();
        }

        private void EnsureDirectoryExists(string path)
        {
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
        }

        private string Print_Excel(string filename, string pathfile)
        {
            try
            {
                string[] sp = filename.Split('_');
                string printerName = sp[1].Substring(0, sp[1].Length - 5);

                if (printerName.Length >= 9 && printerName.Substring(0, 9) == "Microsoft")
                {
                    File.Delete(pathfile);
                    return "Chon lai may in";
                }

                if (printerName.Length >= 5 && printerName.Substring(0, 5) == "Foxit")
                {
                    File.Delete(pathfile);
                    return "Chon lai may in";
                }

                if (printerName.Length >= 3 && printerName.Substring(0, 3) == "Fax")
                {
                    File.Delete(pathfile);
                    return "Chon lai may in";
                }

                bool flag = true;
                string messager = "";
                PrintExcel.Print_xls_file(printerName, pathfile, ref flag, ref messager);

                if (File.Exists(pathfile))
                    File.Delete(pathfile);

                return flag ? "" : sp[0];
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[Print Error] {ex.Message}");
                return "Lỗi kẹt lệnh in 9.245!";
            }
        }

        #endregion
    }
}
