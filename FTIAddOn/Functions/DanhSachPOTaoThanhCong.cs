using AddOn_AC_AL.Models;
using FTIAddOn;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AddOn_AC_AL.Functions
{
    public class DanhSachPOTaoThanhCong
    {
        private SAPbouiCOM.Application SBO_Application;
        private Program program;
        private List<IDValue> iDValues;

        private string formID = "";
        private string formType => this.program.formTypeDSPOTTC;

        private const string FILE_NAME = "DanhSachPOTaoThanhCong.srf";
        private const string DT0_ID = "DT_0";
        private const string MATRIX_ID = "Item_0";

        private SAPbouiCOM.Form oForm => SBO_Application.Forms.Item(formID);
        private SAPbouiCOM.Matrix oMatrix => oForm.Items.Item(MATRIX_ID).Specific;
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

        public DanhSachPOTaoThanhCong(SAPbouiCOM.Application SBO_Application, Program program, List<IDValue> iDValues, string formID)
        {
            this.SBO_Application = SBO_Application;
            this.program = program;
            this.iDValues = iDValues;
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

                oDataTable0.Rows.Add(iDValues.Count);
                for(var i = 0; i < iDValues.Count; i++)
                {
                    oDataTable0.SetValue("Col_0", i, iDValues[i].IDS);
                    oDataTable0.SetValue("Col_1", i, iDValues[i].Value);
                }
                oMatrix.Columns.Item("Col_0").DataBind.Bind(DT0_ID, "Col_0");
                oMatrix.Columns.Item("Col_1").DataBind.Bind(DT0_ID, "Col_1");

                oMatrix.LoadFromDataSourceEx();
                oMatrix.AutoResizeColumns();

                oForm.Visible = true;
            }
            catch (Exception ex)
            {
                SBO_Application.SetStatusBarMessage(ex.Message, BoMessageTime.bmt_Short, true);
            }
        }
    }
}
