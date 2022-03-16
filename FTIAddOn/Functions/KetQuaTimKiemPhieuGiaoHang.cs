using ERPConnect;
using FTIAddOn;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;
using System.Collections;
using AddOn_AC_AL.Models;

namespace AddOn_AC_AL.Functions
{
    public class KetQuaTimKiemPhieuGiaoHang
    {
        private SAPbouiCOM.Application SBO_Application;
        private Program program;
        private RFCTable oD_Ts;

        private string formID = "";
        private string formType => this.program.formTypeKQTKPGH;

        private const string FILE_NAME = "KetQuaTimKiemPhieuGiaoHang.srf";
        private const string DT0_ID = "DT_0";
        private const string DT1_ID = "DT_1";
        private const string GRID_ID = "Item_0";
        private const string BTN_COLLAPSE_ID = "Item_1";
        private const string BTN_EXPAND_ID = "Item_2";
        private const string BTN_SELECT_ALL_ID = "Item_3";
        private const string BTN_TAO_PO_ID = "Item_5";
        private const string BTN_TIM_KIEM_ID = "Item_13";
        private const string BTN_CB_LOC_ID = "Item_15";
        private const string UD_SO_PGH_ID = "UD_0";
        private const string UD_TU_NGAY_ID = "UD_1";
        private const string UD_DEN_NGAY_ID = "UD_2";
        private const string UD_LOC_VALUE_ID = "UD_3";

        private const string KEY_HT_CACHE_DT = "HTCACHEDT";
        private const string KEY_HT_CACHE_DT2 = "HTCACHED2";

        private string Parameters;

        private SAPbouiCOM.Form oForm => SBO_Application.Forms.Item(formID);
        private SAPbouiCOM.DataTable oDataTable0
        {
            get
            {
                return oForm.DataSources.DataTables.Item(DT0_ID);
            }
            set
            {
                oDataTable0 = value;
            }
        }
        private SAPbouiCOM.DataTable oDataTable1 => oForm.DataSources.DataTables.Item(DT1_ID);
        private SAPbouiCOM.Grid oGrid => oForm.Items.Item(GRID_ID).Specific;
        private SAPbouiCOM.UserDataSource uDSoPGH => oForm.DataSources.UserDataSources.Item(UD_SO_PGH_ID);
        private SAPbouiCOM.UserDataSource uDTuNgay => oForm.DataSources.UserDataSources.Item(UD_TU_NGAY_ID);
        private SAPbouiCOM.UserDataSource uDDenNgay => oForm.DataSources.UserDataSources.Item(UD_DEN_NGAY_ID);
        private SAPbouiCOM.UserDataSource uDLocDuLieu => oForm.DataSources.UserDataSources.Item(UD_LOC_VALUE_ID);
        private SAPbouiCOM.Button oBtnCollapse => oForm.Items.Item(BTN_COLLAPSE_ID).Specific;
        private SAPbouiCOM.Button oBtnExpand => oForm.Items.Item(BTN_EXPAND_ID).Specific;
        private SAPbouiCOM.ButtonCombo oBtnComboLoc => oForm.Items.Item(BTN_CB_LOC_ID).Specific;

        public KetQuaTimKiemPhieuGiaoHang(SAPbouiCOM.Application SBO_Application, Program program, RFCTable oD_Ts, string parameters, string formID)
        {
            this.SBO_Application = SBO_Application;
            this.program = program;
            this.formID = formID;
            this.oD_Ts = oD_Ts;
            this.Parameters = parameters;
        }

        public KetQuaTimKiemPhieuGiaoHang(SAPbouiCOM.Application SBO_Application, Program program, string formID)
        {
            this.SBO_Application = SBO_Application;
            this.program = program;
            this.formID = formID;
        }

        public void OpenForm()
        {
            try
            {
                var oXmlDoc = new System.Xml.XmlDocument();
                oXmlDoc.Load("Forms\\" + FILE_NAME);
                var nodeForm = oXmlDoc.ChildNodes.Item(1).ChildNodes.Item(0).ChildNodes.Item(0).ChildNodes.Item(0);
                nodeForm.Attributes["uid"].Value = formID;
                nodeForm.Attributes["FormType"].Value = formType;
                var sXML = oXmlDoc.InnerXml.ToString();
                SBO_Application.LoadBatchActions(ref sXML);
                oForm.Left = 250;
                oForm.Top = 50;
                uDSoPGH.Value = "%";
                uDTuNgay.Value = string.Format("{0:dd.MM.yy}", DateTime.Now);
                uDDenNgay.Value = string.Format("{0:dd.MM.yy}", DateTime.Now);

                oForm.Items.Item("Item_4").TextStyle = 1;
                oForm.Items.Item("Item_6").TextStyle = 1;

                oBtnComboLoc.Item.DisplayDesc = true;
                oBtnComboLoc.ValidValues.Add("1", "Tất cả");
                oBtnComboLoc.ValidValues.Add("2", "OD chưa có trong hệ thống");
                oBtnComboLoc.ValidValues.Add("3", "OD đã tồn tại trong hệ thống");
                oBtnComboLoc.Select(0, BoSearchKey.psk_Index);
                

                //oGrid.DataTable = oDataTable0;
                //oGrid.CollapseLevel = 1;
                //oGrid.AutoResizeColumns();
                oGrid.SelectionMode = BoMatrixSelect.ms_Auto;

                //SetTitleGrid();

                oBtnCollapse.Item.Enabled = false;
                oBtnExpand.Item.Enabled = false;
                //oBtnComboLoc.Item.Enabled = false;

                oForm.Visible = true;
            }
            catch (Exception ex)
            {
                SBO_Application.SetStatusBarMessage(ex.Message, BoMessageTime.bmt_Short, true);
            }
        }

        public void SBO_Application_ItemEvent_BeforeAction(string formUID, ref ItemEvent pVal, out bool bubbleEvent)
        {
            bubbleEvent = true;
            try
            {
                switch (pVal.EventType)
                {
                    case BoEventTypes.et_CLICK:
                        BeforeAction_Click(formUID, ref pVal, out bubbleEvent);
                        break;
                }
            }
            catch (Exception ex)
            {
                SBO_Application.SetStatusBarMessage(ex.Message, BoMessageTime.bmt_Short, true);
            }
        }

        private void BeforeAction_Click(string formUID, ref ItemEvent pVal, out bool bubbleEvent)
        {
            bubbleEvent = true;
            try
            {
                switch (pVal.ItemUID)
                {
                    case GRID_ID:
                        if (pVal.ColUID == "RowsHeader")
                        {
                            if (pVal.Modifiers == BoModifiersEnum.mt_SHIFT)
                            {
                                bubbleEvent = false;
                                SBO_Application.SetStatusBarMessage("Modifier keys 'SHIFT' not use...", BoMessageTime.bmt_Short, true);
                            }
                            if (oGrid.Rows.IsLeaf(pVal.Row))
                            {
                                bubbleEvent = false;
                                SBO_Application.SetStatusBarMessage("Row IsLeaf disable for select...", BoMessageTime.bmt_Short, true);
                            }
                        }
                        break;
                }
            }
            catch (Exception ex)
            {
                SBO_Application.SetStatusBarMessage(ex.Message, BoMessageTime.bmt_Short, true);
            }
        }

        public void SBO_Application_ItemEvent_AfterAction(string formUID, ref ItemEvent pVal, out bool bubbleEvent)
        {
            bubbleEvent = true;
            try
            {
                switch (pVal.EventType)
                {
                    case BoEventTypes.et_CLICK:
                        AfterAction_Click(formUID, ref pVal, out bubbleEvent);
                        break;
                    case BoEventTypes.et_CHOOSE_FROM_LIST:
                        AfterAction_CFL(formUID, ref pVal, out bubbleEvent);
                        break;
                    case BoEventTypes.et_COMBO_SELECT:
                        var a = uDLocDuLieu.Value;
                        break;
                }
            }
            catch (Exception ex)
            {
                SBO_Application.SetStatusBarMessage(ex.Message, BoMessageTime.bmt_Short, true);
            }
        }

        private void AfterAction_CFL(string formUID, ref ItemEvent pVal, out bool bubbleEvent)
        {
            bubbleEvent = true;
            try
            {
                switch (pVal.ItemUID)
                {
                    case GRID_ID:
                        switch (pVal.ColUID)
                        {
                            case "WHSE":
                                Hashtable hTDT = null;
                                if (this.program.hTFormData.ContainsKey(formID) && ((Hashtable)this.program.hTFormData[formID]).ContainsKey(KEY_HT_CACHE_DT2))
                                {
                                    hTDT = (Hashtable)((Hashtable)this.program.hTFormData[formID])[KEY_HT_CACHE_DT2];
                                }
                                if (hTDT == null)
                                {
                                    throw new Exception($"Không tìm thấy dữ liệu cache, vui lòng đóng cửa sổ và thử tải lại!");
                                }
                                var oCFLEvent = (IChooseFromListEvent)pVal;
                                var oDataTable = oCFLEvent.SelectedObjects;
                                var whse = oDataTable.GetValue(0, 0).ToString();
                                var dTRI = oGrid.GetDataTableRowIndex(pVal.Row);
                                if (dTRI != -1)
                                {
                                    oDataTable0.SetValue("WHSE", dTRI, whse);
                                    var row = (Models.DataTable.Row)hTDT[dTRI];
                                    foreach (var cell in row.Cells.Cell)
                                    {
                                        if (cell.ColumnUid == "WHSE")
                                        {
                                            cell.Value = whse;
                                        }
                                    }
                                }
                                break;
                        }
                        break;
                }
            }
            catch (Exception ex)
            {
                SBO_Application.SetStatusBarMessage(ex.Message, BoMessageTime.bmt_Short, true);
            }
        }

        private void AfterAction_Click(string formUID, ref ItemEvent pVal, out bool bubbleEvent)
        {
            bubbleEvent = true;
            try
            {
                switch (pVal.ItemUID)
                {
                    case BTN_TIM_KIEM_ID:
                        var soOD = uDSoPGH.Value;
                        if (string.IsNullOrEmpty(uDTuNgay.Value))
                        {
                            SBO_Application.SetStatusBarMessage("Không được bỏ trống trường 'Từ ngày'!");
                            return;
                        }
                        if (string.IsNullOrEmpty(uDDenNgay.Value))
                        {
                            SBO_Application.SetStatusBarMessage("Không được bỏ trống trường 'Đến ngày'!");
                            return;
                        }
                        var tuNgay = DateTime.ParseExact(uDTuNgay.Value, "dd.MM.yy", null);
                        var denNgay = DateTime.ParseExact(uDDenNgay.Value, "dd.MM.yy", null);
                        Call_YAC_FM_FTI_GET_OD(soOD, tuNgay, denNgay);
                        break;
                    case BTN_COLLAPSE_ID:
                        oGrid.Rows.CollapseAll();
                        break;
                    case BTN_EXPAND_ID:
                        oGrid.Rows.ExpandAll();
                        break;
                    case BTN_SELECT_ALL_ID:
                        oGrid.Rows.SelectedRows.Clear();
                        oDataTable1.Rows.Clear();
                        program.oProgBar = SBO_Application.StatusBar.CreateProgressBar("Đang thực hiện select all...", oGrid.Rows.Count, true);
                        for (var i = 0; i < oGrid.Rows.Count; i++)
                        {
                            if (oGrid.Rows.IsLeaf(i))
                                continue;
                            oGrid.Rows.SelectedRows.Add(i);
                            oDataTable1.Rows.Add(1);
                            oDataTable1.SetValue(0, oDataTable1.Rows.Count - 1, i);
                            program.oProgBar.Value = i + 1;
                        }
                        program.oProgBar.Stop();
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(program.oProgBar);
                        program.oProgBar = null;
                        break;
                    case BTN_TAO_PO_ID:
                        Hashtable hTDT = null;
                        if (!this.program.hTFormData.ContainsKey(formID))
                        {
                            throw new Exception($"Không tìm thấy dữ liệu cache, vui lòng đóng cửa sổ và thử tải lại!");
                        }
                        else
                        {
                            hTDT = (Hashtable)((Hashtable)this.program.hTFormData[formID])[KEY_HT_CACHE_DT];
                        }
                        if (oDataTable1.Rows.Count == 0)
                        {
                            throw new Exception($"Không có bản ghi nào được lựa chọn...");
                        }
                        var rowDatas = new List<RowData>();
                        for (var i = 0; i < oDataTable1.Rows.Count; i++)
                        {
                            var indexGrid = (int)oDataTable1.GetValue(0, i);
                            var rows = (List<Models.DataTable.Row>)hTDT[indexGrid];
                            foreach (var row in rows)
                            {
                                var rowData = new RowData();
                                foreach (var cell in row.Cells.Cell)
                                {
                                    #region "Map data"
                                    switch (cell.ColumnUid)
                                    {
                                        case "WERKS":
                                            rowData.WERKS = cell.Value.ToString();
                                            break;
                                        case "VBELN":
                                            rowData.VBELN = cell.Value.ToString();
                                            break;
                                        case "POSNR":
                                            rowData.POSNR = cell.Value.ToString();
                                            break;
                                        case "KUNAG":
                                            rowData.KUNAG = cell.Value.ToString();
                                            break;
                                        case "SUPPLIER":
                                            rowData.SUPPLIER = cell.Value.ToString();
                                            break;
                                        case "BLDAT":
                                            rowData.BLDAT = cell.Value.ToString();
                                            break;
                                        case "LFDAT":
                                            rowData.LFDAT = cell.Value.ToString();
                                            break;
                                        case "WADAT_IST":
                                            rowData.WADAT_IST = cell.Value.ToString();
                                            break;
                                        case "MATNR":
                                            rowData.MATNR = cell.Value.ToString();
                                            break;
                                        case "ARKTX":
                                            rowData.ARKTX = cell.Value.ToString();
                                            break;
                                        case "WHSE":
                                            rowData.WHSE = cell.Value.ToString();
                                            break;
                                        case "VRKME":
                                            rowData.VRKME = cell.Value.ToString();
                                            break;
                                        case "LFIMG":
                                            rowData.LFIMG = cell.Value.ToString();
                                            break;
                                        case "UNITPRICE":
                                            rowData.UNITPRICE = cell.Value.ToString();
                                            break;
                                        case "KPEIN":
                                            rowData.KPEIN = cell.Value.ToString();
                                            break;
                                        case "WAERK":
                                            rowData.WAERK = cell.Value.ToString();
                                            break;
                                        case "UNITVAT":
                                            rowData.UNITVAT = cell.Value.ToString();
                                            break;
                                        case "BATCH":
                                            rowData.BATCH = cell.Value.ToString();
                                            break;
                                        case "SERIALNO":
                                            rowData.SERIALNO = cell.Value.ToString();
                                            break;
                                        case "VGBEL":
                                            rowData.VGBEL = cell.Value.ToString();
                                            break;
                                        case "VGPOS":
                                            rowData.VGPOS = cell.Value.ToString();
                                            break;
                                        case "LGORT":
                                            rowData.LGORT = cell.Value.ToString();
                                            break;
                                        case "LGORT_AL":
                                            rowData.LGORT_AL = cell.Value.ToString();
                                            break;
                                        default:
                                            break;
                                    }
                                    #endregion
                                }
                                rowDatas.Add(rowData);
                            }
                        }
                        TaoPO(rowDatas);
                        break;
                    case GRID_ID:
                        switch (pVal.ColUID)
                        {
                            case "RowsHeader":
                                if (pVal.Modifiers != BoModifiersEnum.mt_CTRL && !oDataTable1.IsEmpty)
                                {
                                    oDataTable1.Rows.Clear();
                                }
                                oDataTable1.Rows.Add(1);
                                oDataTable1.Rows.Offset = oDataTable1.Rows.Count - 1;
                                oDataTable1.SetValue(0, oDataTable1.Rows.Offset, pVal.Row);
                                break;
                            case "WHSE":
                                break;
                        }
                        //if (pVal.ColUID == "RowsHeader")
                        //{
                        //    if (pVal.Modifiers != BoModifiersEnum.mt_CTRL && !oDataTable1.IsEmpty)
                        //    {
                        //        oDataTable1.Rows.Clear();
                        //    }
                        //    oDataTable1.Rows.Add(1);
                        //    oDataTable1.Rows.Offset = oDataTable1.Rows.Count - 1;
                        //    oDataTable1.SetValue(0, oDataTable1.Rows.Offset, pVal.Row);
                        //}
                        break;
                }
            }
            catch (Exception ex)
            {
                SBO_Application.SetStatusBarMessage(ex.Message, BoMessageTime.bmt_Short, true);
            }
        }

        private void Call_YAC_FM_FTI_GET_OD(string soOD, DateTime tuNgay, DateTime denNgay)
        {
            R3Connection con = null;
            program.oProgBar = SBO_Application.StatusBar.CreateProgressBar("Đang tải dữ liệu...", 10, true);
            try
            {
                var fileConfig = "config.ini";
                ConnectSAP.Class.GetData.openconn(ref con, fileConfig);
                program.oProgBar.Value = 3;
                RFCFunction func = con.CreateFunction("YAC_FM_FTI_GET_OD");
                func.Exports["IM_VBELN"].ParamValue = soOD;
                func.Exports["IM_FR_POSTDATE"].ParamValue = string.Format("{0:yyyyMMdd}", tuNgay);
                func.Exports["IM_TO_POSTDATE"].ParamValue = string.Format("{0:yyyyMMdd}", denNgay);
                var exValue = func.Imports["EX_VALUE"];
                var oD_Ts = func.Tables["LISTOD_T"];
                program.oProgBar.Value = 8;
                func.Execute();
                program.oProgBar.Value = 10;
                program.oProgBar.Stop();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(program.oProgBar);
                program.oProgBar = null;
                if (exValue.ParamValue.Equals(0))
                {
                    SBO_Application.SetStatusBarMessage("Fail OD không hợp lệ!", BoMessageTime.bmt_Short, true);
                    return;
                }
                this.oD_Ts = oD_Ts;
                DisplayData();
            }
            catch (Exception ex)
            {
                program.oProgBar.Stop();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(program.oProgBar);
                program.oProgBar = null;
                con.Close();
                SBO_Application.SetStatusBarMessage(ex.Message, BoMessageTime.bmt_Short, true);
            }
        }

        /// <summary>
        /// Từ chuỗi XML cấu trúc DataTable của SAP tạo đối tượng DataTable trong .NET
        /// </summary>
        /// <returns></returns>
        private Models.DataTable.DataTable GenNetDataTableFormXML()
        {
            try
            {
                var xml = oDataTable0.SerializeAsXML(BoDataTableXmlSelect.dxs_All);
                var serializer = new XmlSerializer(typeof(Models.DataTable.DataTable));
                Models.DataTable.DataTable dataTable = null;
                using (var sr = new StringReader(xml))
                {
                    dataTable = (Models.DataTable.DataTable)serializer.Deserialize(sr);
                }
                dataTable.Rows = new Models.DataTable.Rows();
                dataTable.Rows.Row = new List<Models.DataTable.Row>();
                return dataTable;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        /// <summary>
        /// Map dữ liệu trả về từ RFC (System A1) sang dữ liệu DataTable trong .NET
        /// Chuẩn bị dữ liệu HashTable để lưu vào memory
        /// </summary>
        /// <returns>
        /// + Item1 Models.DataTable.DataTable
        /// + Item2 (Key: VBELN, Value: Models.DataTable.Row)
        /// + Item3 (Key: IndexOfDataTable SAP, Value: Models.DataTable.Row)
        /// + Item4 List<string> thứ tự VBELN RFC trả về
        /// </returns>
        private Tuple<Models.DataTable.DataTable, Hashtable, Hashtable, List<string>> MapDataRFCToDataTable()
        {
            program.oProgBar = SBO_Application.StatusBar.CreateProgressBar("Đang xử lý dữ liệu...", oD_Ts.RowCount, true);
            try
            {
                var dataTable = GenNetDataTableFormXML();
                var hTODTs = new Hashtable();
                var hTDTs = new Hashtable();
                var vBELNs = new List<string>();
                for (var i = 0; i < oD_Ts.RowCount; i++)
                {
                    var keyHT = (string)oD_Ts[i, "VBELN"];
                    var row = new Models.DataTable.Row();
                    row.Cells = new Models.DataTable.Cells();
                    row.Cells.Cell = new List<Models.DataTable.Cell>();
                    for (var j = 0; j < oD_Ts.Columns.Count; j++)
                    {
                        var columnName = oD_Ts.Columns[j].Name;
                        var cellValue = oD_Ts[i, columnName];
                        row.Cells.Cell.Add(new Models.DataTable.Cell()
                        {
                            ColumnUid = columnName,
                            Value = cellValue,
                        });
                    }
                    row.Cells.Cell.Add(new Models.DataTable.Cell()
                    {
                        ColumnUid = "WHSE",
                        Value = "",
                    });
                    if (hTODTs.ContainsKey(keyHT))
                    {
                        ((List<Models.DataTable.Row>)hTODTs[keyHT]).Add(row);
                    }
                    else
                    {
                        hTODTs[keyHT] = new List<Models.DataTable.Row>() { row };
                        vBELNs.Add(keyHT);
                    }
                    hTDTs.Add(i, row);
                    dataTable.Rows.Row.Add(row);
                    program.oProgBar.Value = i + 1;
                }
                program.oProgBar.Stop();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(program.oProgBar);
                program.oProgBar = null;
                return new Tuple<Models.DataTable.DataTable, Hashtable, Hashtable, List<string>> (dataTable, hTODTs, hTDTs, vBELNs);
            }
            catch (Exception ex)
            {
                program.oProgBar.Stop();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(program.oProgBar);
                program.oProgBar = null;
                throw new Exception($"MapDataRFCToDataTable -> {ex.Message}");
            }
        }

        /// <summary>
        /// Xử lý và lưu dữ liệu vào memory để xử lý nhanh hơn.
        /// </summary>
        /// <param name="hTODTs"></param>
        /// <param name="hTDTs"></param>
        /// <param name="vBELNs"></param>
        private void StoreDataCache(Hashtable hTODTs, Hashtable hTDTs, List<string> vBELNs)
        {
            try
            {
                var hTVal = new Hashtable();
                var hTIndexGrid = new Hashtable();
                var indexGrid = 0;
                foreach (var vBELN in vBELNs)
                {
                    var valueHT = (List<Models.DataTable.Row>)hTODTs[vBELN];
                    hTIndexGrid.Add(indexGrid, valueHT);
                    indexGrid += valueHT.Count + 1;
                }
                hTVal.Add(KEY_HT_CACHE_DT, hTIndexGrid);
                hTVal.Add(KEY_HT_CACHE_DT2, hTDTs);
                if (this.program.hTFormData.ContainsKey(formID))
                {
                    this.program.hTFormData.Remove(formID);
                }
                this.program.hTFormData.Add(formID, hTVal);
            }
            catch (Exception ex)
            {
                throw new Exception($"StoreDataCache -> {ex.Message}");
            }
        }

        /// <summary>
        /// Tải dữ liệu vào DataTable SAP từ XML (Nhanh hơn)
        /// </summary>
        /// <param name="dataTable"></param>
        private void DataTableLoadDataFromXML(Models.DataTable.DataTable dataTable)
        {
            try
            {
                var ser = new XmlSerializer(typeof(Models.DataTable.DataTable));
                var ms = new MemoryStream();
                ser.Serialize(ms, dataTable);
                ms.Position = 0;
                var r = new StreamReader(ms);
                var res = r.ReadToEnd();
                oDataTable0.LoadFromXML(res);
            } catch(Exception ex)
            {
                throw new Exception($"DataTableLoadDataFromXML -> {ex.Message}");
            }
        }

        private void DisplayData()
        {
            oForm.Freeze(true);
            try
            {
                var t = MapDataRFCToDataTable();
                StoreDataCache(t.Item2, t.Item3, t.Item4);
                DataTableLoadDataFromXML(t.Item1);

                oGrid.DataTable = oDataTable0;
                oGrid.CollapseLevel = 1;
                oGrid.AutoResizeColumns();
                oGrid.SelectionMode = BoMatrixSelect.ms_Auto;

                for (var i = 0; i < oGrid.Columns.Count; i++)
                {
                    if (oGrid.Columns.Item(i).UniqueID == "WHSE")
                        continue;
                    oGrid.Columns.Item(i).Editable = false;
                }

                SetTitleGrid();

                oGrid.Columns.Item("WHSE").Width = 130;
                oGrid.Columns.Item("WHSE").Type = BoGridColumnType.gct_EditText;
                EditTextColumn editCol = (EditTextColumn)oGrid.Columns.Item("WHSE");
                editCol.ChooseFromListUID = "CFL_0";
                editCol.ChooseFromListAlias = "WhsCode";
                editCol.LinkedObjectType = "64";

                oForm.Freeze(false);
            }
            catch (Exception ex)
            {
                oForm.Freeze(false);
                SBO_Application.SetStatusBarMessage(ex.Message, BoMessageTime.bmt_Short, true);
            }
        }

        private void SetTitleGrid()
        {
            try
            {
                oGrid.Columns.Item("WERKS").TitleObject.Caption = "Plant - công ty An Cường Group";
                oGrid.Columns.Item("VBELN").TitleObject.Caption = "Mã OD-Giao hàng";
                oGrid.Columns.Item("VBELN").ForeColor = System.Convert.ToInt32(0xFF0000);
                oGrid.Columns.Item("VBELN").TextStyle = 1;
                oGrid.Columns.Item("POSNR").TitleObject.Caption = "Hạng mục OD";
                oGrid.Columns.Item("KUNAG").TitleObject.Caption = "Mã Khách hàng";
                oGrid.Columns.Item("SUPPLIER").TitleObject.Caption = "SUPPLIER";
                oGrid.Columns.Item("BLDAT").TitleObject.Caption = "Document Date";
                oGrid.Columns.Item("LFDAT").TitleObject.Caption = "Delivery Date";
                oGrid.Columns.Item("WADAT_IST").TitleObject.Caption = "Posting Date";
                oGrid.Columns.Item("MATNR").TitleObject.Caption = "Mã hàng hóa";
                oGrid.Columns.Item("ARKTX").TitleObject.Caption = "Tên hàng hóa";
                oGrid.Columns.Item("WHSE").TitleObject.Caption = "Kho";
                oGrid.Columns.Item("VRKME").TitleObject.Caption = "Đơn vị tính";
                oGrid.Columns.Item("LFIMG").TitleObject.Caption = "Số lượng giao";
                oGrid.Columns.Item("UNITPRICE").TitleObject.Caption = "Đơn giá";
                oGrid.Columns.Item("KPEIN").TitleObject.Caption = "Đơn vị giá";
                oGrid.Columns.Item("WAERK").TitleObject.Caption = "Loại tiền";
                oGrid.Columns.Item("BATCH").TitleObject.Caption = "Số Lô";
                oGrid.Columns.Item("VGBEL").TitleObject.Caption = "Số SO-Đơn hàng";
                oGrid.Columns.Item("VGPOS").TitleObject.Caption = "Hạng mục SO-Đơn hàng";
                oGrid.Columns.Item("LGORT").TitleObject.Caption = "Mã Kho xuất của An Cường";
                oGrid.Columns.Item("LGORT_AL").TitleObject.Caption = "Kho ảo Ái Linh";
                oGrid.Columns.Item("SERIALNO").TitleObject.Caption = "Số Serial number";
            }
            catch (Exception ex)
            {
                SBO_Application.SetStatusBarMessage(ex.Message, BoMessageTime.bmt_Short, true);
            }
        }

        private void TaoPO(List<RowData> rowDatas)
        {
            try
            {
                var oCompany = (SAPbobsCOM.Company)SBO_Application.Company.GetDICompany();
                var sQL = "SELECT \"UgpEntry\", \"UgpCode\" FROM \"OUGP\"";
                var oRecordSet = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRecordSet.DoQuery(sQL);
                var hTOUGP = new Hashtable();
                while (!oRecordSet.EoF)
                {
                    var ugpEntry = oRecordSet.Fields.Item(0).Value;
                    var ugpCode = oRecordSet.Fields.Item(1).Value;
                    if (!hTOUGP.ContainsKey(ugpCode))
                    {
                        hTOUGP.Add(ugpCode, ugpEntry);
                    }
                    oRecordSet.MoveNext();
                }
                var docEntrys = new List<IDValue>();
                foreach (var gR in rowDatas.GroupBy(it => it.VBELN))
                {
                    try
                    {
                        var first = gR.First();
                        var oDocuments = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);
                        oDocuments.DocObjectCode = SAPbobsCOM.BoObjectTypes.oPurchaseOrders;
                        oDocuments.CardCode = first.SUPPLIER;
                        oDocuments.TaxDate = DateTime.ParseExact(first.BLDAT, "yyyyMMdd", null);
                        oDocuments.DocDueDate = DateTime.ParseExact(first.LFDAT, "yyyyMMdd", null);
                        oDocuments.DocDate = DateTime.ParseExact(first.WADAT_IST, "yyyyMMdd", null);
                        var oLines = oDocuments.Lines;
                        var lineNum = 0;
                        foreach (var row in gR)
                        {
                            oLines.Add();
                            oLines.SetCurrentLine(lineNum);
                            oLines.ItemCode = row.MATNR;
                            if (hTOUGP.ContainsKey(row.VRKME))
                            {
                                oLines.UoMEntry = (int)hTOUGP[row.VRKME];
                            }
                            oLines.Quantity = double.Parse(row.LFIMG);
                            oLines.UnitPrice = double.Parse(row.UNITPRICE);
                            oLines.LineTotal = oLines.Quantity * oLines.UnitPrice;
                            oLines.WarehouseCode = row.WHSE;
                            lineNum++;
                        }
                        var ret = oDocuments.Add();
                        if (ret == 0)
                        {
                            var docEntry = "";
                            oCompany.GetNewObjectCode(out docEntry);
                            docEntrys.Add(new IDValue()
                            {
                                IDS = docEntry,
                                Value = first.VBELN
                            });
                        }
                        else
                        {
                            var errCode = 0;
                            var errMes = "";
                            oCompany.GetLastError(out errCode, out errMes);
                            throw new Exception($"Lỗi xảy ra khi tạo PO: {first.VBELN} -> {errCode}-{errMes}");
                        }
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oLines);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oDocuments);
                    }
                    catch (Exception ex)
                    {
                        SBO_Application.SetStatusBarMessage(ex.Message, BoMessageTime.bmt_Short, true);
                    }
                }
                if (docEntrys.Count != 0)
                {
                    var form = new DanhSachPOTaoThanhCong(this.SBO_Application, this.program, docEntrys, Guid.NewGuid().ToString().Substring(0, 8));
                    form.OpenForm();
                    form = null;
                }
                else
                {
                    SBO_Application.SetStatusBarMessage("Không có PO nào được tạo thành công!", BoMessageTime.bmt_Short, true);
                }
            }
            catch (Exception ex)
            {
                SBO_Application.SetStatusBarMessage(ex.Message, BoMessageTime.bmt_Short, true);
            }
        }
    }
}
