using ERPConnect;
using FTIAddOn;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AddOn_AC_AL.Functions
{
    public class TimKiemPhieuGiaoHang
    {
        private Program program;
        private SAPbouiCOM.Application SBO_Application;

        private string formID = "";
        private string formType => this.program.formTypeTKPGH;

        private const string FILE_NAME = "TimKiemPhieuGiaoHang.srf";
        private const string TXT_SO_PHEU = "Item_1";
        private const string TXT_TU_NGAY = "Item_3";
        private const string TXT_DEN_NGAY = "Item_5";
        private const string UDS_SO_PHEU = "UD_0";
        private const string UDS_TU_NGAY = "UD_1";
        private const string UDS_DEN_NGAY = "UD_2";
        private const string BTN_OK = "Item_6";

        private SAPbouiCOM.Form oForm => SBO_Application.Forms.Item(formID);
        private SAPbouiCOM.UserDataSource uDS_SoPhieu => oForm.DataSources.UserDataSources.Item(UDS_SO_PHEU);
        private SAPbouiCOM.UserDataSource uDS_TuNgay => oForm.DataSources.UserDataSources.Item(UDS_TU_NGAY);
        private SAPbouiCOM.UserDataSource uDS_DenNgay => oForm.DataSources.UserDataSources.Item(UDS_DEN_NGAY);

        public TimKiemPhieuGiaoHang(SAPbouiCOM.Application SBO_Application, Program program, string formID)
        {
            this.SBO_Application = SBO_Application;
            this.formID = formID;
            this.program = program;
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
                // load the form to the SBO application in one batch
                var sXML = oXmlDoc.InnerXml.ToString();
                SBO_Application.LoadBatchActions(ref sXML);
                oForm.Left = 400;
                oForm.Top = 100;
                //uDS_SoPhieu.Value = "%";
                uDS_SoPhieu.Value = "6000603981";
                uDS_TuNgay.Value = string.Format("{0:dd.MM.yy}", DateTime.Now);
                uDS_DenNgay.Value = string.Format("{0:dd.MM.yy}", DateTime.Now);
                oForm.Visible = true;
            }
            catch (Exception ex)
            {
                SBO_Application.SetStatusBarMessage(ex.Message, BoMessageTime.bmt_Medium, true);
            }
        }

        public void SBO_Application_ItemEvent_AfterAction(string FormUID, ref ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                switch (pVal.EventType)
                {
                    case BoEventTypes.et_CLICK:
                        AfterAction_Click(FormUID, ref pVal, out BubbleEvent);
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
                var kQTKPGH = new KetQuaTimKiemPhieuGiaoHang(SBO_Application, program, oD_Ts, string.Format("{0} | {1:dd/MM/yyyy} | {2:dd/MM/yyyy}", soOD, tuNgay, denNgay), Guid.NewGuid().ToString().Substring(0, 8));
                kQTKPGH.OpenForm();
                kQTKPGH = null;
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

        private void AfterAction_Click(string formUID, ref ItemEvent pVal, out bool bubbleEvent)
        {
            bubbleEvent = true;
            try
            {
                switch (pVal.ItemUID)
                {
                    case BTN_OK:
                        var soOD = uDS_SoPhieu.Value;
                        if (string.IsNullOrEmpty(uDS_TuNgay.Value))
                        {
                            SBO_Application.SetStatusBarMessage("Không được bỏ trống trường 'Từ ngày'!");
                            return;
                        }
                        if (string.IsNullOrEmpty(uDS_DenNgay.Value))
                        {
                            SBO_Application.SetStatusBarMessage("Không được bỏ trống trường 'Đến ngày'!");
                            return;
                        }
                        var tuNgay = DateTime.ParseExact(uDS_TuNgay.Value, "dd.MM.yy", null);
                        var denNgay = DateTime.ParseExact(uDS_DenNgay.Value, "dd.MM.yy", null);
                        Call_YAC_FM_FTI_GET_OD(soOD, tuNgay, denNgay);
                        break;
                }
            }
            catch (Exception ex)
            {
                SBO_Application.SetStatusBarMessage(ex.Message, BoMessageTime.bmt_Medium, true);
            }
        }
    }
}
