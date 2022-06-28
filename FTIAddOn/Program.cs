using System;
using SAPbouiCOM;
using AddOn_AC_AL.Functions;
using System.Collections;
using SAPbobsCOM;

//0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056

namespace FTIAddOn
{

    public class Run
    {
        static void Main()
        {
            var program = new Program();
            System.Windows.Forms.Application.Run();
        }
    }

    public class Program
    {
        /// <summary>
        /// ProgressBar System
        /// </summary>
        public SAPbouiCOM.ProgressBar oProgBar { get; set; }
        /// <summary>
        /// ID: FORMID, Value: Object
        /// </summary>
        public Hashtable hTFormData { get; set; }

        private SAPbouiCOM.Application SBO_Application;
        private SAPbobsCOM.Company oCompany;

        public SAPbobsCOM.Company Company => oCompany;

        /// <summary>
        /// MENUID: Tìm kiếm phiếu giao hàng
        /// </summary>
        private const string MENU_TKPGH_ID = "fa2372be";

        /// <summary>
        /// FORM TYPE: Tìm kiếm phiếu giao hàng
        /// </summary>
        private const string FORM_TYPE_TKPGH = "UF_TKPGH";
        public string formTypeTKPGH => FORM_TYPE_TKPGH;

        /// <summary>
        /// FORM TYPE: Kết quản tìm kiếm phiếu giao hàng
        /// </summary>
        private const string FORM_TYPE_KQTKPGH = "UF_KQTKPGH";
        public string formTypeKQTKPGH => FORM_TYPE_KQTKPGH;

        private const string FORM_TYPE_DSPOTTC = "UF_DSPOTTC";
        public string formTypeDSPOTTC => FORM_TYPE_DSPOTTC;

        private BoDataServerTypes oDBServerType;
        public BoDataServerTypes DBServerType => oDBServerType;

        public Program()
        {
            hTFormData = new Hashtable();
            SetApplication();
            CreateMenu();
            SetFilters();
            EventHandlers();
        }

        private void SetApplication()
        {

            // *******************************************************************
            // Use an SboGuiApi object to establish connection
            // with the SAP Business One application and return an
            // initialized appliction object
            // *******************************************************************

            SAPbouiCOM.SboGuiApi SboGuiApi = null;
            string sConnectionString = null;

            SboGuiApi = new SAPbouiCOM.SboGuiApi();

            // by following the steped specified above the following
            // statment should be suficient for either development or run mode

            sConnectionString = System.Convert.ToString(Environment.GetCommandLineArgs().GetValue(1));

            // connect to a running SBO Application

            SboGuiApi.Connect(sConnectionString);

            // get an initialized application object

            SBO_Application = SboGuiApi.GetApplication(-1);

            oCompany = SBO_Application.Company.GetDICompany();

            oDBServerType = oCompany.DbServerType;

            SBO_Application.SetStatusBarMessage("Connected!", SAPbouiCOM.BoMessageTime.bmt_Short, false);
        }

        private void SetFilters()
        {
            SAPbouiCOM.EventFilters oFilters;
            SAPbouiCOM.EventFilter oFilter;

            // Create a new EventFilters object
            oFilters = new SAPbouiCOM.EventFilters();

            // add an event type to the container
            // this method returns an EventFilter object

            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_CLICK);
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST);
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_MENU_CLICK);
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_COMBO_SELECT);
            oFilter.AddEx(FORM_TYPE_KQTKPGH);

            SBO_Application.SetFilter(oFilters);
        }

        private void EventHandlers()
        {
            SBO_Application.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(SBO_Application_AppEvent);
            SBO_Application.MenuEvent += new SAPbouiCOM._IApplicationEvents_MenuEventEventHandler(SBO_Application_MenuEvent);
            SBO_Application.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(SBO_Application_ItemEvent);
            SBO_Application.ProgressBarEvent += new SAPbouiCOM._IApplicationEvents_ProgressBarEventEventHandler(SBO_Application_ProgressBarEvent);
            //SBO_Application.StatusBarEvent += new SAPbouiCOM._IApplicationEvents_StatusBarEventEventHandler(SBO_Application_StatusBarEvent);
            //SBO_Application.FormDataEvent += new SAPbouiCOM._IApplicationEvents_FormDataEventEventHandler(SBO_Application_FormDataEvent);
            //SBO_Application.RightClickEvent += new SAPbouiCOM._IApplicationEvents_RightClickEventEventHandler(SBO_Application_RightClickEvent);
            //SBO_Application.LayoutKeyEvent += new SAPbouiCOM._IApplicationEvents_LayoutKeyEventEventHandler(SBO_Application_LayoutKeyEvent);
        }

        private void SBO_Application_ProgressBarEvent(ref ProgressBarEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            if (pVal.EventType == SAPbouiCOM.BoProgressBarEventTypes.pbet_ProgressBarStopped & pVal.BeforeAction)
            {
                SBO_Application.MessageBox("Progress Bar stopped by user, releasing progress bar", 1, "Ok", "", "");
                // Stopping the progress bar, thus loosing it's values.
                if (oProgBar != null)
                {
                    oProgBar.Stop();
                    oProgBar = null;
                }
            }
        }

        private void SBO_Application_MenuEvent(ref MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            if (!pVal.BeforeAction)
            {
                switch (pVal.MenuUID)
                {
                    case MENU_TKPGH_ID:
                        var kQTKPGH = new KetQuaTimKiemPhieuGiaoHang(SBO_Application, this, Guid.NewGuid().ToString().Substring(0, 8));
                        kQTKPGH.OpenForm();
                        kQTKPGH = null;
                        break;
                }
            }
        }

        private void SBO_Application_ItemEvent(string FormUID, ref ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            if (!pVal.BeforeAction)
            {
                switch (pVal.FormTypeEx)
                {
                    case FORM_TYPE_KQTKPGH:
                        var kQTKPGH = new KetQuaTimKiemPhieuGiaoHang(SBO_Application, this, pVal.FormUID);
                        kQTKPGH.SBO_Application_ItemEvent_AfterAction(FormUID, ref pVal, out BubbleEvent);
                        kQTKPGH = null;
                        break;
                }
            }
            else
            {
                switch (pVal.FormTypeEx)
                {
                    case FORM_TYPE_KQTKPGH:
                        var kQTKPGH = new KetQuaTimKiemPhieuGiaoHang(SBO_Application, this, pVal.FormUID);
                        kQTKPGH.SBO_Application_ItemEvent_BeforeAction(FormUID, ref pVal, out BubbleEvent);
                        kQTKPGH = null;
                        break;
                }
            }
        }

        private void SBO_Application_AppEvent(BoAppEventTypes EventType)
        {
            switch (EventType)
            {
                case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
                    System.Windows.Forms.Application.Exit();
                    break;
            }
        }

        private void CreateMenu()
        {
            FTIGlobal.PublicFunctions.CreateMenu(MENU_TKPGH_ID, "Tìm kiếm phiếu giao hàng - An Cường", BoMenuType.mt_STRING, "2304", SBO_Application);
        }

    }
}
