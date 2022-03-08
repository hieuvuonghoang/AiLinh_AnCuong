﻿using ERPConnect;
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
        private const string BTN_TAO_SO_ID = "Item_5";
        private const string BTN_TIM_KIEM_ID = "Item_13";
        private const string UD_SO_PGH_ID = "UD_0";
        private const string UD_TU_NGAY_ID = "UD_1";
        private const string UD_DEN_NGAY_ID = "UD_2";

        private const string KEY_HT_CACHE_DT = "HTCACHEDT";

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
                //nodeForm.Attributes["title"].Value += ": " + Parameters;
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

                oGrid.DataTable = oDataTable0;
                oGrid.CollapseLevel = 1;
                oGrid.AutoResizeColumns();
                oGrid.SelectionMode = BoMatrixSelect.ms_Auto;

                SetTitleGrid();

                oForm.Visible = true;
            }
            catch (Exception ex)
            {
                SBO_Application.SetStatusBarMessage(ex.Message, BoMessageTime.bmt_Medium, true);
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
                SBO_Application.SetStatusBarMessage(ex.Message, BoMessageTime.bmt_Medium, true);
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
                SBO_Application.SetStatusBarMessage(ex.Message, BoMessageTime.bmt_Medium, true);
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
                }
            }
            catch (Exception ex)
            {
                SBO_Application.SetStatusBarMessage(ex.Message, BoMessageTime.bmt_Medium, true);
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
                    case BTN_TAO_SO_ID:
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
                        for (var i = 0; i < oDataTable1.Rows.Count; i++)
                        {
                            var indexGrid = (int)oDataTable1.GetValue(0, i);
                            var rows = (List<Models.Row>)hTDT[indexGrid];
                            foreach(var row in rows)
                            {
                                foreach(var cell in row.Cells.Cell)
                                {
                                    var cellVal = cell.Value;
                                }
                            }
                        }
                        break;
                    case GRID_ID:
                        if (pVal.ColUID == "RowsHeader")
                        {
                            if (pVal.Modifiers != BoModifiersEnum.mt_CTRL && !oDataTable1.IsEmpty)
                            {
                                oDataTable1.Rows.Clear();
                            }
                            oDataTable1.Rows.Add(1);
                            oDataTable1.Rows.Offset = oDataTable1.Rows.Count - 1;
                            oDataTable1.SetValue(0, oDataTable1.Rows.Offset, pVal.Row);
                        }
                        break;
                }
            }
            catch (Exception ex)
            {
                SBO_Application.SetStatusBarMessage(ex.Message, BoMessageTime.bmt_Medium, true);
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
                    SBO_Application.SetStatusBarMessage("Fail OD không hợp lệ!", BoMessageTime.bmt_Medium, true);
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
                SBO_Application.SetStatusBarMessage(ex.Message, BoMessageTime.bmt_Medium, true);
            }
        }

        private void DisplayData()
        {
            program.oProgBar = SBO_Application.StatusBar.CreateProgressBar("Đang xử lý dữ liệu...", oD_Ts.RowCount, true);
            oForm.Freeze(true);
            try
            {
                var xml = oDataTable0.SerializeAsXML(BoDataTableXmlSelect.dxs_All);
                var serializer = new XmlSerializer(typeof(Models.DataTable));
                Models.DataTable dataTable = null;
                using (var sr = new StringReader(xml))
                {
                    dataTable = (Models.DataTable)serializer.Deserialize(sr);
                }

                dataTable.Rows = new Models.Rows();
                dataTable.Rows.Row = new List<Models.Row>();

                var hTODTs = new Hashtable();
                var sTTValues = new List<IDValue>();
                var sTT = 1;
                for (var i = 0; i < oD_Ts.RowCount; i++)
                {
                    var keyHT = (string)oD_Ts[i, "VBELN"];
                    var row = new Models.Row();
                    row.Cells = new Models.Cells();
                    row.Cells.Cell = new List<Models.Cell>();
                    for (var j = 0; j < oD_Ts.Columns.Count; j++)
                    {
                        var columnName = oD_Ts.Columns[j].Name;
                        var cellValue = oD_Ts[i, columnName];
                        row.Cells.Cell.Add(new Models.Cell()
                        {
                            ColumnUid = columnName,
                            Value = cellValue,
                        });
                    }
                    if (hTODTs.ContainsKey(keyHT))
                    {
                        ((List<Models.Row>)hTODTs[keyHT]).Add(row);
                    }
                    else
                    {
                        hTODTs[keyHT] = new List<Models.Row>() { row };
                        sTTValues.Add(new IDValue() { ID = sTT, Value = keyHT });
                    }
                    dataTable.Rows.Row.Add(row);
                    program.oProgBar.Value = i + 1;
                }
                program.oProgBar.Stop();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(program.oProgBar);
                program.oProgBar = null;

                var ser = new XmlSerializer(typeof(Models.DataTable));

                var ms = new MemoryStream();
                ser.Serialize(ms, dataTable);
                ms.Position = 0;

                var r = new StreamReader(ms);
                var res = r.ReadToEnd();

                oDataTable0.LoadFromXML(res);

                oGrid.DataTable = oDataTable0;
                oGrid.CollapseLevel = 1;
                oGrid.AutoResizeColumns();
                oGrid.SelectionMode = BoMatrixSelect.ms_Auto;

                for (var i = 0; i < oGrid.Columns.Count; i++)
                {
                    oGrid.Columns.Item(i).Editable = false;
                }

                SetTitleGrid();

                var hTVal = new Hashtable();
                var hTIndexGrid = new Hashtable();
                var indexGrid = 0;
                foreach (var iDValue in sTTValues)
                {
                    var valueHT = (List<Models.Row>)hTODTs[iDValue.Value];
                    hTIndexGrid.Add(indexGrid, valueHT);
                    indexGrid += valueHT.Count + 1;
                }
                hTVal.Add(KEY_HT_CACHE_DT, hTIndexGrid);
                if (this.program.hTFormData.ContainsKey(formID))
                {
                    this.program.hTFormData.Remove(formID);
                }
                this.program.hTFormData.Add(formID, hTVal);
                oForm.Freeze(false);
            } catch (Exception ex)
            {
                program.oProgBar.Stop();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(program.oProgBar);
                program.oProgBar = null;
                oForm.Freeze(false);
                SBO_Application.SetStatusBarMessage(ex.Message, BoMessageTime.bmt_Medium, true);
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
            } catch (Exception ex)
            {
                SBO_Application.SetStatusBarMessage(ex.Message, BoMessageTime.bmt_Medium, true);
            }
        }
    }
}
