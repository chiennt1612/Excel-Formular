using System;
using System.Text;
using OfficeOpenXml;
using Newtonsoft.Json.Linq;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Data;

namespace Utils
{
    public class ReadExcelForm
    {
        #region Properties
        ExcelPackage package;
        private dynamic DLanguge;
        private dynamic dParam;
        private dynamic dData;
        private dynamic dPivot;
        private ObjSheets[] ObjSheet;
        private ObjSheetsPivot[] ObjSheetPivot;

        private const int MaxWhileCount = 10000;
        private const int SheetCount = 10;
        private const int PivotCount = 5;
        private const int DataSourceCount = 10;
        private const int DataSourceCloumnCount = 1000;
        private const string CharSplitDataSource = ".";
        private const string CharSplitDataPivot = "|";

        private const string EndTag = "]";
        private const string BeginVarable = "&=$[";
        private const string BeginVarableLang = "&=$L[";
        private const string BeginStaticValue = "&==["; // xử lý phiên bản sau
        private const string BeginRepeatDynamicFormula = "&=$D[";
        private const string BeginDataPivot = "&=Pivot[";
        private const string BeginDataSource = "&=[";
        #endregion

        #region Contructor
        public ReadExcelForm(ExcelPackage _package, string _dParam, string _dData, string _dPivot)
        {
            DLanguge = null;
            dParam = JObject.Parse(_dParam);
            dData = JObject.Parse(_dData);
            dPivot = JObject.Parse(_dPivot);
            package = _package;
            ObjSheet = new ObjSheets[SheetCount];
            ObjSheetPivot = new ObjSheetsPivot[SheetCount];
            ReadExcelPivot();
            WriteExcelPivot(ObjSheetPivot);
            ReadExcel();
            for (int i = 0; i < ObjSheet.Length; i++)
            {
                if (ObjSheet[i] != null)
                {
                    if (ObjSheet[i].IsDataList)
                    {
                        ExcelWorksheet w = package.Workbook.Worksheets[ObjSheet[i].SheetName];
                        int row; int l; int k;
                        for (l = 0; l < ObjSheet[i].CountDataSource; l++)
                        {
                            var ObjData = ObjSheet[i].DataSrcList[l];
                            string PivotFormula = ""; int PivotCol = 0;
                            if (ObjSheetPivot[i] != null)
                            {
                                if (ObjSheetPivot[i].IsPivot)
                                {
                                    var obj = ObjSheetPivot[i].CellPivot[0];
                                    if (obj.Formula != "" && ObjData.DataSrcName == obj.DataSource)
                                    {
                                        PivotCol = obj.CellCol;
                                        PivotFormula = obj.Formula;
                                    }
                                }
                            }
                            dynamic d = dData[ObjData.DataSrcName];
                            int ColStart = ObjData.CellDataSrcStart;
                            int ColEnd = ObjData.CellDataSrcEnd;
                            row = ObjData.CellRowStart + (l > 0 ? (d.Items.Count - 1) : 0);
                            if (d.Items.Count > 0)
                            {
                                WriteCellData(d, 0, row, w, ObjSheet[i], ColStart, ColEnd, PivotCol, PivotFormula);
                                for (k = 1; k < d.Items.Count; k++)//10; k++)// 
                                {
                                    row++;
                                    w.InsertRow(row, 1, row - 1);
                                    WriteCellData(d, k, row, w, ObjSheet[i], ColStart, ColEnd, PivotCol, PivotFormula);
                                }
                            }

                        }
                    }
                }
            }
        }
        public ReadExcelForm(ExcelPackage _package, dynamic _DLanguge, string _dParam, string _dData, string _dPivot)
        {
            DLanguge = _DLanguge;
            dParam = JObject.Parse(_dParam);
            dData = JObject.Parse(_dData);
            dPivot = JObject.Parse(_dPivot);
            package = _package;
            ObjSheet = new ObjSheets[SheetCount];
            ObjSheetPivot = new ObjSheetsPivot[SheetCount];
            ReadExcelPivot();
            WriteExcelPivot(ObjSheetPivot);
            ReadExcel();

            for(int i=0; i <  ObjSheet.Length; i ++) {
                if (ObjSheet[i] != null && ObjSheet[i].IsDataList)
                {
                    ExcelWorksheet w = package.Workbook.Worksheets[ObjSheet[i].SheetName];
                    int row;

                    for(int l = 0; l< ObjSheet[i].CountDataSource; l++) {
                        var ObjData = ObjSheet[i].DataSrcList[l];
                        string PivotFormula = ""; int PivotCol = 0;
                        if (ObjSheetPivot[i] != null)
                        {
                            if (ObjSheetPivot[i].IsPivot)
                            {
                                var obj = ObjSheetPivot[i].CellPivot[0];
                                if (obj.Formula != "" && ObjData.DataSrcName == obj.DataSource)
                                {
                                    PivotCol = obj.CellCol;
                                    PivotFormula = obj.Formula;
                                }
                            }
                        }
                        dynamic d = dData[ObjData.DataSrcName];
                        int ColStart = ObjData.CellDataSrcStart;
                        int ColEnd = ObjData.CellDataSrcEnd;
                        //DataTable t = new DataTable();
                        //ColDataTable(t, ColStart, ColEnd);
                        row = ObjData.CellRowStart + (l > 0 ? (d.Items.Count - 1) : 0);
                        int dCnt = d.Items.Count;
                        if (dCnt > 0)
                        {
                            //WriteCellData(d, 0, row, w, ObjSheet[i], ColStart, ColEnd, PivotCol, PivotFormula);
                            //RowToDataTable(t, d, 0, row, ObjSheet[i], ColStart, ColEnd, PivotCol, PivotFormula);

                            //List<dynamic> items = d.Items.ToObject<List<dynamic>>();
                            w.InsertRow(row + 1, dCnt - 1, row);
                            //Parallel.For(0, items.Count, k =>
                            //{                                
                            //RowToDataTable(t, d, k, row, ObjSheet[i], ColStart, ColEnd, PivotCol, PivotFormula);
                            //w.InsertRow(row, 1, row - 1);
                            //WriteCellData(d, k, row + k, w, ObjSheet[i], ColStart, ColEnd, PivotCol, PivotFormula);
                            //row++;
                            //});
                            for (int k = 0; k < dCnt; k++)
                            {
                                //RowToDataTable(t, d, k, row, ObjSheet[i], ColStart, ColEnd, PivotCol, PivotFormula);
                                //w.InsertRow(row, 1, row - 1);
                                WriteCellData(d, k, row, w, ObjSheet[i], ColStart, ColEnd, PivotCol, PivotFormula);
                                row++;
                            };
                            //w.InsertRow(row, d.Items.Count);
                            //w.Cells[row, ColStart, d.Items.Count, ColEnd].LoadFromDataTable(t, true);
                        }
                    };
                }
            };
            //for (int i = 0; i < ObjSheet.Length; i++)
            //{
            //    if (ObjSheet[i] != null)
            //    {
            //        if (ObjSheet[i].IsDataList)
            //        {
            //            ExcelWorksheet w = package.Workbook.Worksheets[ObjSheet[i].SheetName];
            //            int row; int l; int k;
            //            for (l = 0; l < ObjSheet[i].CountDataSource; l++)
            //            {
            //                var ObjData = ObjSheet[i].DataSrcList[l];
            //                string PivotFormula = ""; int PivotCol = 0;
            //                if (ObjSheetPivot[i] != null)
            //                {
            //                    if (ObjSheetPivot[i].IsPivot)
            //                    {
            //                        var obj = ObjSheetPivot[i].CellPivot[0];
            //                        if (obj.Formula != "" && ObjData.DataSrcName == obj.DataSource)
            //                        {
            //                            PivotCol = obj.CellCol;
            //                            PivotFormula = obj.Formula;
            //                        }
            //                    }
            //                }
            //                dynamic d = dData[ObjData.DataSrcName];
            //                int ColStart = ObjData.CellDataSrcStart;
            //                int ColEnd = ObjData.CellDataSrcEnd;
            //                row = ObjData.CellRowStart + (l > 0 ? (d.Items.Count - 1) : 0);
            //                if (d.Items.Count > 0)
            //                {
            //                    WriteCellData(d, 0, row, ref w, ref ObjSheet[i], ColStart, ColEnd, PivotCol, PivotFormula);
            //                    for (k = 1; k < d.Items.Count; k++)//10; k++)// 
            //                    {
            //                        row++;
            //                        w.InsertRow(row, 1, row - 1);
            //                        WriteCellData(d, k, row, ref w, ref ObjSheet[i], ColStart, ColEnd, PivotCol, PivotFormula);
            //                    }
            //                }

            //            }
            //        }
            //    }
            //}
        }
        ~ReadExcelForm()
        {
            dParam = null; dData = null; dPivot = null;
            package = null;
            ObjSheet = null;
        }
        #endregion

        #region Private
        // Xử lý Pivot và Biến; Language
        private void ReadCellDataPivot(ref ExcelWorksheet worksheet, ref ObjSheetsPivot ObjS)
        {
            var rowCount = worksheet.Dimension?.Rows;
            var colCount = worksheet.Dimension?.Columns;
            if (!rowCount.HasValue || !colCount.HasValue)
            {
                return;
            }
            for (int row = 1; row <= rowCount.Value; row++)
            {
                for (int col = 1; col <= colCount.Value; col++)
                {
                    dynamic CellData = worksheet.Cells[row, col].Value;
                    if (CellData != null)
                    {
                        // thay thế nếu là Varable languageư
                        string cellData = CellData.ToString(); bool IsTag = false; string Formula = "";
                        try
                        {
                            cellData = ReplaceStringValue(cellData, BeginVarableLang, EndTag, ref IsTag, true);
                            cellData = ReplaceStringValue(cellData, BeginVarable, EndTag, ref IsTag, false);

                        }
                        catch (Exception) { }
                        
                        if (IsTag) // Varable language
                            worksheet.Cells[row, col].Value = cellData;
                        else
                        {
                            SearchTagByCell(cellData, BeginStaticValue, EndTag, ref IsTag, ref Formula);
                            if (IsTag)
                            {
                                worksheet.Cells[row, col].Formula = Formula;
                            }
                            else
                            {
                                // thay thế nếu là Varable Param
                                //cellData = CellData.ToString(); IsTag = false;
                                //cellData = ReplaceStringValue(cellData, BeginVarable, EndTag, ref IsTag, false);
                                //if (IsTag) // Varable
                                //    worksheet.Cells[row, col].Value = cellData;
                                //else
                                //{
                                //BeginDataPivot
                                cellData = CellData.ToString(); IsTag = false;
                                string[] a;
                                SearchTagByCell(cellData, BeginDataPivot, EndTag, ref IsTag, ref Formula);
                                if (IsTag)
                                {
                                    //PivotColumn.Name.ID:ListData:2:Formula
                                    string[] c = Formula.Split(new string[] { CharSplitDataPivot }, StringSplitOptions.None);
                                    a = c[0].Split(new string[] { CharSplitDataSource }, StringSplitOptions.None);
                                    ObjS.IsPivot = IsTag;
                                    ObjS.CellPivot[ObjS.CountCellPivot] = new ObjPivot();
                                    ObjS.CellPivot[ObjS.CountCellPivot].CellRow = row;
                                    ObjS.CellPivot[ObjS.CountCellPivot].CellCol = col;
                                    ObjS.CellPivot[ObjS.CountCellPivot].PivotVal = a; // PivotColumn.Name.ID
                                    ObjS.CellPivot[ObjS.CountCellPivot].DataSource = c[1]; //ListData

                                    if (c.Length > 2)
                                        ObjS.CellPivot[ObjS.CountCellPivot].DataSourceRow = int.Parse(c[2]); //2
                                    else
                                        ObjS.CellPivot[ObjS.CountCellPivot].DataSourceRow = 1;

                                    if (c.Length > 3)
                                    {
                                        ObjS.CellPivot[ObjS.CountCellPivot].Formula = c[3]; //Formula
                                    }
                                    else
                                        ObjS.CellPivot[ObjS.CountCellPivot].Formula = "";

                                    ObjS.CountCellPivot++;
                                }
                                //}
                            }
                        }
                    }
                }
            }
        }
        private void ReadExcelPivot()
        {
            for (int i = 1; i <= (package.Workbook.Worksheets.Count > SheetCount ? SheetCount : package.Workbook.Worksheets.Count); i++)
            {
                var worksheet = package.Workbook.Worksheets[i];
                ObjSheetPivot[i - 1] = new ObjSheetsPivot();
                ObjSheetPivot[i - 1].SheetName = worksheet.Name;
                ObjSheetPivot[i - 1].IsPivot = false;
                ObjSheetPivot[i - 1].CountCellPivot = 0;
                ObjSheetPivot[i - 1].CellPivot = new ObjPivot[PivotCount];
                ReadCellDataPivot(ref worksheet, ref ObjSheetPivot[i - 1]);
            }
        }
        private void WriteExcelPivot(ObjSheetsPivot[] objSheet)
        {
            for (int i = 0; i < objSheet.Length; i++)
            {
                if (objSheet[i] != null)
                {
                    if (objSheet[i].IsPivot)
                    {
                        ExcelWorksheet w = package.Workbook.Worksheets[objSheet[i].SheetName];
                        int row; int col; int l; int cntColAdded = 0;
                        for (int j = 0; j < objSheet[i].CountCellPivot; j++)
                        {
                            /*
                            ObjS.CellPivot[ObjS.CountCellPivot].PivotVal = a; // PivotColumn.Name.ID
                            ObjS.CellPivot[ObjS.CountCellPivot].DataSource = c[1]; //ListData
                            ObjS.CellPivot[ObjS.CountCellPivot].DataSourceRow = int.Parse(c[2]); //2
                            */
                            var obj = objSheet[i].CellPivot[j];
                            row = obj.CellRow;
                            col = obj.CellCol + cntColAdded;
                            dynamic d1 = dPivot[obj.PivotVal[0]];
                            //GetDataJson(HRSContext context, dynamic d, int i, string sKey, int j = 0, bool IsFormat = false)
                            object ValS1 = Tools.GetDataJson(DLanguge, d1, 0, obj.PivotVal[1], 0, 2);
                            w.Cells[row, col].Value = ValS1; // d1.Items[0][obj.PivotVal[1]].ToString();
                            ValS1 = Tools.GetDataJson(DLanguge, d1, 0, obj.PivotVal[2], 0, 2);
                            w.Cells[row + obj.DataSourceRow, col].Value = BeginDataSource + obj.DataSource + "." + ValS1/*d1.Items[0][obj.PivotVal[2]].ToString()*/ + EndTag;

                            //=SUM([{C}10:{C}11]).4_=CHOOSE(WEEKDAY([{C}8]), "Sun", "Mon","Tue","Wed","Thu","Fri","Sat").-1_=DAY([{C}8]).-2]
                            string[] aFormual = obj.Formula.Split(new string[] { "_" }, StringSplitOptions.None);
                            if (obj.Formula != "")
                            {
                                for (int ia = 0; ia < aFormual.Length; ia++)
                                {
                                    string[] b = aFormual[ia].Split(new string[] { "." }, StringSplitOptions.None);
                                    //formual[ia] = w.Cells[row + int.Parse(aFormual[ia]), col].Formula;
                                    //formual = formual.Replace("{c}", col.ToString());
                                    w.Cells[row + int.Parse(b[1]), col].Formula = CreateFormula(w, col, b[0]);
                                }
                                //if (obj.Formula != "") SetValidation(ref w, w.Cells[row + obj.DataSourceRow, col].Address, aFormual[0]);
                                //w.Cells[row, col].Value = 0;
                            }
                            for (l = 1; l < d1.Items.Count; l++)
                            {
                                col++;
                                w.InsertColumn(col, 1, col - 1); cntColAdded++;
                                ValS1 = Tools.GetDataJson(DLanguge, d1, l, obj.PivotVal[1], 0, 2);
                                w.Cells[row, col].Value = ValS1; // d1.Items[l][obj.PivotVal[1]].ToString();
                                ValS1 = Tools.GetDataJson(DLanguge, d1, l, obj.PivotVal[2], 0, 2);
                                w.Cells[row + obj.DataSourceRow, col].Value = BeginDataSource + obj.DataSource + "." + ValS1/*d1.Items[l][obj.PivotVal[2]].ToString()*/ + EndTag;
                                if (obj.Formula != "")
                                {
                                    for (int ia = 0; ia < aFormual.Length; ia++)
                                    {
                                        string[] b = aFormual[ia].Split(new string[] { "." }, StringSplitOptions.None);
                                        //formual[ia] = w.Cells[row + int.Parse(aFormual[ia]), col].Formula;
                                        //formual = formual.Replace("{c}", col.ToString());
                                        w.Cells[row + int.Parse(b[1]), col].Formula = CreateFormula(w, col, b[0]);
                                    }
                                    //if (obj.Formula != "") SetValidation(ref w, w.Cells[row + obj.DataSourceRow, col].Address, obj.Formula);
                                }
                            }
                        }
                    }
                }
            }
        }
        private static string CreateFormula(ExcelWorksheet w, int col, string ExcelFormula)
        {
            string BeginTagCell = "{"; string EndTagCell = "}";
            string CellAddress = ExcelFormula;
            int i; int j;
            i = CellAddress.IndexOf(BeginTagCell);
            j = CellAddress.IndexOf(EndTagCell);
            if (i == -1 || j == -1 || i == j) return CellAddress;
            string s1; s1 = CellAddress.Substring(i + BeginTagCell.Length, j - i - EndTagCell.Length);
            string[] a = s1.Split(new string[] { ":" }, StringSplitOptions.None);
            if (a.Length < 2) //=DAY([8])
            {
                CellAddress = CellAddress.Replace(BeginTagCell + s1 + EndTagCell, w.Cells[int.Parse(a[0]), col].Address);
            }
            else // =SUM([10:11])
            {
                CellAddress = CellAddress.Replace(BeginTagCell + s1 + EndTagCell, w.Cells[int.Parse(a[0]), col, int.Parse(a[1]), col].Address);
            }
            return CellAddress;
        }
        // hàm dùng chung
        private static void SetValidation(ref ExcelWorksheet worksheet, string CellAddress, string ExcelFormula)
        {
            try
            {
                var validation = worksheet.DataValidations.AddListValidation(CellAddress);
                validation.ShowErrorMessage = true;
                validation.ErrorStyle = OfficeOpenXml.DataValidation.ExcelDataValidationWarningStyle.warning;
                validation.ErrorTitle = "An invalid value was entered";
                validation.Error = "Select a value from the list";                                                                 
                validation.Formula.ExcelFormula = ExcelFormula; 
            }
            catch
            {

            }
        }
        private string ReplaceStringValue(string s, string BeginTag, string _EndTag, ref bool IsVal, int Row, int Dir)
        {
            int i; int j;
            i = s.IndexOf(BeginTag);
            j = s.IndexOf(_EndTag);
            if (i == -1 || j == -1 || i == j)
            {
                IsVal = false;
                return s;
            }
            else
            {
                int iWhile = 0; // chặn lỗi Out Of Memory
                while (!(i == -1 || j == -1 || i == j) && iWhile < MaxWhileCount)
                {
                    string s1; iWhile++;
                    if (Dir == 0)
                    {
                        s1 = s.Substring(i + BeginTag.Length - 1, j - i - BeginTag.Length);
                        s = s.Replace(BeginTag + s1 + _EndTag, Row.ToString());
                    }
                    else if (Dir == 1)
                    {
                        s1 = s.Substring(i + BeginTag.Length, j - i - BeginTag.Length);
                        s = s.Replace(BeginTag + s1 + _EndTag, (Row + int.Parse(s1)).ToString());
                    }
                    else
                    {
                        s1 = s.Substring(i + BeginTag.Length, j - i - BeginTag.Length);
                        s = s.Replace(BeginTag + s1 + _EndTag, (Row - int.Parse(s1)).ToString());
                    }
                    i = s.IndexOf(BeginTag);
                    j = s.IndexOf(_EndTag);
                }
                IsVal = true;
                return s;
            }
        }
        private object GetLanguageLabel(string Key, bool IsLang = true)
        {
            object r;
            if (IsLang)
            {
                if (DLanguge == null) return Key;

                try
                {
                    r = DLanguge[Key];//Tools.GetDataJson(dLang, Key);
                }
                catch
                {
                    r = Key;
                }
            } 
            else
            {
                r = dParam[Key];// Tools.GetDataJson(dParam, Key);
                if (DLanguge == null) return Key;
                try
                {
                    r = DLanguge[r];
                }
                catch { }
            }

            return r;
        }
        private void SearchTagByCell(string s, string BeginTag, string _EndTag, ref bool IsVal, ref string[] a, ref int aCount)
        {
            int i; int j;
            i = s.IndexOf(BeginTag);
            j = s.IndexOf(_EndTag);

            if (i == -1 || j == -1 || i == j)
            {
                IsVal = false;
                a = "".Split(new string[] { CharSplitDataSource }, StringSplitOptions.None);
                aCount = a.Length;
            }
            else
            {
                IsVal = true; string s1 = s.Substring(i + BeginTag.Length, j - i - BeginTag.Length);
                a = s1.Split(new string[] { CharSplitDataSource }, StringSplitOptions.None);
                aCount = a.Length;
            }
        }
        private void SearchTagByCell(string s, string BeginTag, string _EndTag, ref bool IsVal, ref string Formula)
        {
            int i; int j; Formula = "";
            i = s.IndexOf(BeginTag);
            j = s.IndexOf(_EndTag);

            if (i == -1 || j == -1 || i == j)
            {
                IsVal = false;
                Formula = "";
            }
            else
            {
                IsVal = true; Formula = s.Substring(i + BeginTag.Length, j - i - BeginTag.Length);
            }
        }
        private string ReplaceStringValue(string s, string BeginTag, string _EndTag, ref bool IsVal, bool IsLang = true)
        {
            int i; int j;
            i = s.IndexOf(BeginTag);
            j = s.IndexOf(_EndTag);
            if (i == -1 || j == -1 || i == j)
            {
                //IsVal = false;
                return s;
            }
            else
            {
                int iWhile = 0; // chặn lỗi Out Of Memory
                while (!(i == -1 || j == -1 || i == j) && iWhile < MaxWhileCount)
                {
                    string s1 = s.Substring(i + BeginTag.Length, j - i - BeginTag.Length);
                    s = s.Replace(BeginTag + s1 + _EndTag, GetLanguageLabel(s1, IsLang).ToString());
                    i = s.IndexOf(BeginTag);
                    j = s.IndexOf(_EndTag);
                    iWhile++;
                }
                IsVal = true;
                return s;
            }
        }

        private void ColDataTable(DataTable t, int CellDataSrcStart = 0, int CellDataSrcEnd = 0)
        {
            for (int i = 0; i <= (CellDataSrcEnd - CellDataSrcStart) ; i++)
            {
                t.Columns.Add("C" + i, typeof(object));
            }
        }
        private void RowToDataTable(DataTable t, dynamic d, int iData, int row,
            ObjSheets ObjS, int CellDataSrcStart = 0, int CellDataSrcEnd = 0, int PivotCol = 0, string PivotFormula = "")
        {
            DataRow dtRow = t.NewRow();
            for (int i = 0; i < t.Columns.Count; i++)
            {              
                //if (ObjS.CellDataSource[i].Formula != "")
                //{
                //    if (!ObjS.CellDataSource[i].IsValue)
                //        worksheet.Cells[row, ObjS.CellDataSource[i].CellCol].Formula = ExecFormula(ObjS.CellDataSource[i].Formula, row);
                //    else
                //        worksheet.Cells[row, ObjS.CellDataSource[i].CellCol].Value = ObjS.CellDataSource[i].Formula;
                //}
                //else
                //{
                    //if (PivotFormula != "" && ObjS.CellDataSource[i].CellCol >= PivotCol) SetValidation(ref worksheet, worksheet.Cells[row, ObjS.CellDataSource[i].CellCol].Address, PivotFormula);
                    object ValS1 = Tools.GetDataJson(DLanguge, d, iData, ObjS.CellDataSource[i].DataSource[1], 0, 2);
                //worksheet.Cells[row, ObjS.CellDataSource[i].CellCol].Value = ValS1;
                dtRow[i] = ValS1;
                //}

            }
            t.Rows.Add(dtRow);
        }
        // Xử lý datasource và Formula
        private void WriteCellData(dynamic d, int iData, int row, ExcelWorksheet worksheet, ObjSheets ObjS, int CellDataSrcStart = 0, int CellDataSrcEnd = 0,
               int PivotCol = 0, string PivotFormula = "")
        {
            for (int i = CellDataSrcStart; i <= CellDataSrcEnd; i++)
            {
                if (ObjS.CellDataSource[i].Formula != "")
                {
                    if (!ObjS.CellDataSource[i].IsValue)
                        worksheet.Cells[row, ObjS.CellDataSource[i].CellCol].Formula = ExecFormula(ObjS.CellDataSource[i].Formula, row);
                    else
                        worksheet.SetValue(row, ObjS.CellDataSource[i].CellCol, ObjS.CellDataSource[i].Formula);
                }
                else
                {
                    //if (PivotFormula != "" && ObjS.CellDataSource[i].CellCol >= PivotCol) SetValidation(ref worksheet, worksheet.Cells[row, ObjS.CellDataSource[i].CellCol].Address, PivotFormula);
                    object ValS1 = /*Random.Equals(1, 100000); //*/Tools.GetDataJson(DLanguge, d, iData, ObjS.CellDataSource[i].DataSource[1], 0, 2);
                    worksheet.SetValue(row, ObjS.CellDataSource[i].CellCol, ValS1);
                }

            }
        }
        private string ExecFormula(string Formula, int row)
        {
            string BeginTag = "{r"; string BeginTagInc = "{r+"; string BeginTagDec = "{r-";
            string EndTag = "}"; string r; bool IsVal = false;
            r = ReplaceStringValue(Formula, BeginTagInc, EndTag, ref IsVal, row, 1);
            if (!IsVal)
            {
                r = ReplaceStringValue(Formula, BeginTagDec, EndTag, ref IsVal, row, -1);
                if (!IsVal)
                {
                    r = ReplaceStringValue(Formula, BeginTag, EndTag, ref IsVal, row, 0);
                    if (!IsVal)
                    {
                        r = Formula;
                    }
                }
            }
            return r;
        }
        private void ReadExcel()
        {
            for (int i = 1; i <= (package.Workbook.Worksheets.Count > SheetCount ? SheetCount : package.Workbook.Worksheets.Count); i++)
            {
                var worksheet = package.Workbook.Worksheets[i];
                ObjSheet[i - 1] = new ObjSheets();
                ObjSheet[i - 1].SheetName = worksheet.Name;
                ObjSheet[i - 1].IsDataList = false;
                ObjSheet[i - 1].CountCellData = 0;
                ObjSheet[i - 1].CellDataSource = new ObjDataSource[DataSourceCloumnCount];
                ObjSheet[i - 1].CountDataSource = 0;
                ObjSheet[i - 1].DataSrcList = new ClsDataSrcList[DataSourceCount];
                ReadCellData(ref worksheet, ref ObjSheet[i - 1]);
            }
        }
        private void ReadCellData(ref ExcelWorksheet worksheet, ref ObjSheets ObjS)
        {
            var rowCount = worksheet.Dimension?.Rows;
            var colCount = worksheet.Dimension?.Columns;
            if (!rowCount.HasValue || !colCount.HasValue)
            {
                return;
            }
            string DataSrcNamePrev = "";
            for (int row = 1; row <= rowCount.Value; row++)
            {
                for (int col = 1; col <= colCount.Value; col++)
                {
                    dynamic CellData = worksheet.Cells[row, col].Value;
                    if (CellData != null)
                    {
                        //BeginDataSource
                        string cellData = CellData.ToString(); bool IsTag = false; string Formula = ""; int aCount; string[] a;
                        cellData = CellData.ToString(); IsTag = false;
                        a = "a".Split(new string[] { CharSplitDataSource }, StringSplitOptions.None); aCount = 0;
                        SearchTagByCell(cellData, BeginDataSource, EndTag, ref IsTag, ref a, ref aCount);
                        if (IsTag) ObjS.IsDataList = IsTag;
                        if (IsTag)
                        {
                            if (DataSrcNamePrev != a[0])
                            {
                                ObjS.DataSrcList[ObjS.CountDataSource] = new ClsDataSrcList();
                                ObjS.DataSrcList[ObjS.CountDataSource].CellRowStart = row;
                                ObjS.DataSrcList[ObjS.CountDataSource].CellDataSrcStart = ObjS.CountCellData;
                                if (ObjS.CountDataSource > 0) ObjS.DataSrcList[ObjS.CountDataSource - 1].CellDataSrcEnd = ObjS.CountCellData - 1;
                                ObjS.DataSrcList[ObjS.CountDataSource].DataSrcName = a[0];
                                DataSrcNamePrev = a[0];
                                ObjS.CountDataSource++;
                            }
                            ObjS.DataSrcList[ObjS.CountDataSource - 1].CellDataSrcEnd = ObjS.CountCellData;

                            ObjS.CellDataSource[ObjS.CountCellData] = new ObjDataSource();
                            ObjS.CellDataSource[ObjS.CountCellData].CellRow = row;
                            ObjS.CellDataSource[ObjS.CountCellData].CellCol = col;
                            ObjS.CellDataSource[ObjS.CountCellData].DataSource = a;
                            ObjS.CellDataSource[ObjS.CountCellData].aCount = aCount;
                            ObjS.CellDataSource[ObjS.CountCellData].Formula = "";
                            ObjS.CellDataSource[ObjS.CountCellData].IsValue = false;
                            ObjS.CountCellData++;
                        }
                        SearchTagByCell(cellData, BeginRepeatDynamicFormula, EndTag, ref IsTag, ref Formula);
                        if (IsTag)
                        {
                            ObjS.DataSrcList[ObjS.CountDataSource - 1].CellDataSrcEnd = ObjS.CountCellData;

                            ObjS.CellDataSource[ObjS.CountCellData] = new ObjDataSource();
                            ObjS.CellDataSource[ObjS.CountCellData].CellRow = row;
                            ObjS.CellDataSource[ObjS.CountCellData].CellCol = col;
                            ObjS.CellDataSource[ObjS.CountCellData].Formula = Formula;
                            ObjS.CellDataSource[ObjS.CountCellData].IsValue = false;
                            ObjS.CountCellData++;
                        }
                    }
                }
            }
        }
        #endregion
    }
    public class ObjPivot
    {
        public int CellRow;
        public int CellCol;
        public string[] PivotVal;
        public string DataSource;
        public int DataSourceRow;
        public string Formula;
    }
    public class ObjDataSource
    {
        public int CellRow;
        public int CellCol;
        public string[] DataSource;
        public int aCount;
        public string Formula;
        public bool IsValue;
    }
    public class ClsDataSrcList
    {
        public int CellRowStart;
        public int CellDataSrcStart;
        public int CellDataSrcEnd;
        public string DataSrcName;
    }
    public class ObjSheetsPivot
    {
        public string SheetName;
        public bool IsPivot;
        public int CountCellPivot;
        public ObjPivot[] CellPivot;
    }
    public class ObjSheets
    {
        public string SheetName;
        public bool IsDataList;
        public int CountCellData;
        public ObjDataSource[] CellDataSource;
        public int CountDataSource;
        public ClsDataSrcList[] DataSrcList;
    }
}
