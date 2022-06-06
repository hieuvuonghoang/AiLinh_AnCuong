using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FTIAddOn;
using SAPbouiCOM;
using System.IO;
using System.Xml.Serialization;
using AddOn_AC_AL.Models.OUGP;
using System.Collections;

namespace AddOn_AC_AL.Functions
{
    class TestFunction
    {
        private Program program;
        private Application sBO_Application;

        private string formID;
        private string formType => this.program.formTypeTFUNC;

        private const string FILE_NAME = "TestFunction.srf";

        private const string BTN_TEST_ID = "Item_0";
       

        private SAPbouiCOM.Form oForm => sBO_Application.Forms.Item(formID);

        public TestFunction(Application sBO_Application, Program program, string formID)
        {
            this.sBO_Application = sBO_Application;
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
                sBO_Application.LoadBatchActions(ref sXML);
                oForm.Left = 250;
                oForm.Top = 50;

                oForm.Visible = true;
            }
            catch (Exception ex)
            {
                sBO_Application.SetStatusBarMessage(ex.Message, BoMessageTime.bmt_Medium, true);
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
                sBO_Application.SetStatusBarMessage(ex.Message, BoMessageTime.bmt_Medium, true);
            }
        }

        private void AfterAction_Click(string formUID, ref ItemEvent pVal, out bool bubbleEvent)
        {
            bubbleEvent = true;
            try
            {
                switch (pVal.ItemUID)
                {
                    case BTN_TEST_ID:


                        //var sQL = "SELECT UgpEntry, UgpCode FROM OUGP";
                        //var oRecordset = this.program.Recordset;
                        //oRecordset.DoQuery(sQL);
                        //var xml = oRecordset.GetAsXML();
                        //var serializer = new XmlSerializer(typeof(BOM));
                        //var hT = new Hashtable();
                        //using (var reader = new StringReader(xml))
                        //{
                        //    var bOM = (BOM)serializer.Deserialize(reader);
                        //    foreach(var row in bOM.BO.OUGP.Row)
                        //    {
                        //        if(!hT.ContainsKey(row.UgpEntry))
                        //        {
                        //            hT.Add(row.UgpEntry, row);
                        //        }
                        //    }
                        //}
                        break;
                }
            }
            catch (Exception ex)
            {
                sBO_Application.SetStatusBarMessage(ex.Message, BoMessageTime.bmt_Medium, true);
            }
        }

        private void CreateSO()
        {
            try
            {
                var oCompany = (SAPbobsCOM.Company)sBO_Application.Company.GetDICompany();
                var oCompanyService = oCompany.GetCompanyService();
                var oAdminInfo = oCompanyService.GetAdminInfo();
                oAdminInfo.EnableApprovalProcedureInDI = SAPbobsCOM.BoYesNoEnum.tYES;
                oAdminInfo.DocConfirmation = SAPbobsCOM.BoYesNoEnum.tYES;
                oCompanyService.UpdateAdminInfo(oAdminInfo);

                var oDocument = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders);
                if (oDocument.GetApprovalTemplates() == 0 && oDocument.Document_ApprovalRequests.ApprovalTemplatesID > 0)
                {
                    for (var j = 0; j < oDocument.Document_ApprovalRequests.Count; j++)
                    {
                        oDocument.Document_ApprovalRequests.SetCurrentLine(j);
                        var templateName = oDocument.Document_ApprovalRequests.ApprovalTemplatesName;
                        var templateID = oDocument.Document_ApprovalRequests.ApprovalTemplatesID;
                    }
                }
            }
            catch (Exception ex)
            {
                sBO_Application.SetStatusBarMessage(ex.Message, BoMessageTime.bmt_Medium, true);
            }
        }

    }
}
