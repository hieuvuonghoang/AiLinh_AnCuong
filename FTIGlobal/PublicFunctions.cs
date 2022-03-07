using FTIB1Core.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;

namespace FTIGlobal
{
    public static class PublicFunctions
    {
        #region Các hàm dùng chung

        public static DataTable RsToDataTable(SAPbobsCOM.Recordset _rs)
        {
            DataTable dt = new DataTable();
            for (int i = 0; i < _rs.Fields.Count; i++)
                dt.Columns.Add(_rs.Fields.Item(i).Description);
            while (!_rs.EoF)
            {
                DataRow row = dt.NewRow();
                for (int i = 0; i < _rs.Fields.Count; i++)
                    row[i] = _rs.Fields.Item(i).Value;
                dt.Rows.Add(row.ItemArray);
                _rs.MoveNext();
            }
            return dt;
        }

        /// <summary>
        /// Method create menu user
        /// </summary>
        /// <param name="menuId">Identity menu</param>
        /// <param name="menuName">Menu name (display)</param>
        /// <param name="menuType">Menu type</param>
        /// <param name="parentMenuId">Identity menu parent</param>
        /// <param name="SBO_Application">SAPbouiCOM Application</param>
        /// <remarks></remarks>
        public static void CreateMenu(string menuId, string menuName, SAPbouiCOM.BoMenuType menuType, string parentMenuId,
            SAPbouiCOM.Application SBO_Application)
        {
            if (SBO_Application.Menus.Exists(menuId)) return;
            try
            {
                SAPbouiCOM.MenuCreationParams oMenuCreationParams = default(SAPbouiCOM.MenuCreationParams);
                SAPbouiCOM.MenuItem oMenuItem = default(SAPbouiCOM.MenuItem);
                SAPbouiCOM.Menus oMenus = default(SAPbouiCOM.Menus);
                oMenuCreationParams = (SAPbouiCOM.MenuCreationParams)SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                oMenuCreationParams.Type = menuType;
                oMenuCreationParams.UniqueID = menuId;
                oMenuCreationParams.String = menuName;
                oMenuCreationParams.Enabled = true;
                if (SBO_Application.Menus.Exists(parentMenuId))
                {
                    oMenuItem = SBO_Application.Menus.Item(parentMenuId);
                    oMenus = oMenuItem.SubMenus;
                    oMenuCreationParams.Position = oMenus.Count + 1;
                    oMenus.AddEx(oMenuCreationParams);
                }
                oMenuCreationParams = null;
                oMenuItem = null;
                oMenus = null;
            }
            catch (Exception ex)
            {
                SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        public static SAPbouiCOM.Form GetFormUDO(SAPbouiCOM.Application SBO_Application, string UniqueID, string FileSrf, string ObjectType = "", bool checkHasForm = false)
        {
            SAPbouiCOM.Form functionReturnValue = default(SAPbouiCOM.Form);
            SAPbouiCOM.FormCreationParams fcp = default(SAPbouiCOM.FormCreationParams);
            try
            {
                fcp = (SAPbouiCOM.FormCreationParams)SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
                fcp.FormType = UniqueID;
                fcp.UniqueID = UniqueID;
                if (!string.IsNullOrEmpty(ObjectType))
                    fcp.ObjectType = ObjectType;
                fcp.XmlData = LoadFromXML(FileSrf);
                try
                {
                    functionReturnValue = SBO_Application.Forms.AddEx(fcp);
                    //SetLangueForItem(functionReturnValue, functionReturnValue.UniqueID, GetCaptionItem());
                    checkHasForm = false;
                    return functionReturnValue;
                }
                catch (Exception e)
                {
                    Console.Write("\n" + e.Message);
                    functionReturnValue = SBO_Application.Forms.Item(UniqueID);
                    functionReturnValue.Select();
                    checkHasForm = true;
                    return functionReturnValue;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static string LoadFromXML(string FileName)
        {
            System.Xml.XmlDocument oXmlDoc = null;
            oXmlDoc = new System.Xml.XmlDocument();
            //sPath = Application.StartupPath & "\"
            oXmlDoc.Load(FileName);
            return (oXmlDoc.InnerXml);
        }

        private static void SetLangueForItem(SAPbouiCOM.Form Form, string FormUID, System.Data.DataTable g_CaptionItem)
        {
            SAPbouiCOM.Item Item = default(SAPbouiCOM.Item);
            SAPbouiCOM.StaticText Static = default(SAPbouiCOM.StaticText);
            SAPbouiCOM.Folder Folder = default(SAPbouiCOM.Folder);
            SAPbouiCOM.Column Column = default(SAPbouiCOM.Column);
            SAPbouiCOM.Button Button = default(SAPbouiCOM.Button);
            SAPbouiCOM.Matrix Matrix = default(SAPbouiCOM.Matrix);
            SAPbouiCOM.CheckBox CheckBox = default(SAPbouiCOM.CheckBox);
            System.Data.DataRow[] RowArray = null;
            int Count = 0;
            if (g_CaptionItem == null)
                return;
            if (g_CaptionItem.Rows.Count <= 0) return;
            RowArray = g_CaptionItem.Select("U_FormUID='" + FormUID.ToUpper() + "'");
            if (RowArray.Length > 0)
            {
                for (Count = 0; Count <= RowArray.Length - 1; Count++)
                {
                    try
                    {
                        if (!string.IsNullOrEmpty(RowArray[Count]["U_Item"].ToString().Trim()))
                        {
                            if (!string.IsNullOrEmpty(RowArray[Count]["U_Text"].ToString().Trim()))
                            {
                                Item = Form.Items.Item(RowArray[Count]["U_Item"].ToString().Trim());
                                switch (Item.Type)
                                {
                                    case SAPbouiCOM.BoFormItemTypes.it_BUTTON:
                                        Button = (SAPbouiCOM.Button)Item.Specific;
                                        Button.Caption = RowArray[Count]["U_Text"].ToString().Trim();
                                        break;
                                    case SAPbouiCOM.BoFormItemTypes.it_FOLDER:
                                        Folder = (SAPbouiCOM.Folder)Item.Specific;
                                        Folder.Caption = RowArray[Count]["U_Text"].ToString().Trim();
                                        break;
                                    case SAPbouiCOM.BoFormItemTypes.it_MATRIX:
                                        Matrix = (SAPbouiCOM.Matrix)Item.Specific;
                                        Column = Matrix.Columns.Item(RowArray[Count]["U_Column"].ToString().Trim());
                                        Column.TitleObject.Caption = RowArray[Count]["U_Text"].ToString().Trim();
                                        break;
                                    case SAPbouiCOM.BoFormItemTypes.it_GRID:
                                        break;
                                    //Grid = Item.Specific
                                    //Grid.Columns.Item(RowArray(Count).Item("U_Column").ToString.Trim()).TitleObject.Caption = RowArray(Count).Item("U_Text").ToString.Trim
                                    case SAPbouiCOM.BoFormItemTypes.it_STATIC:
                                        Static = (SAPbouiCOM.StaticText)Item.Specific;
                                        Static.Caption = RowArray[Count]["U_Text"].ToString().Trim();
                                        break;
                                    case SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX:
                                        CheckBox = (SAPbouiCOM.CheckBox)Item.Specific;
                                        CheckBox.Caption = RowArray[Count]["U_Text"].ToString().Trim();
                                        break;
                                }
                            }
                        }
                        else
                        {
                            if (!string.IsNullOrEmpty(RowArray[Count]["U_Text"].ToString().Trim()))
                            {
                                Form.Title = RowArray[Count]["U_Text"].ToString().Trim();
                            }
                        }

                    }
                    catch (Exception ex)
                    {
                        throw ex;
                        //MessageBox.Show(ex.ToString());
                    }
                }
            }
        }

        public static void SaveDataTableToXML(string filePath, System.Data.DataTable dt)
        {
            try
            {
                if (File.Exists(filePath))
                    File.Delete(filePath);
                System.IO.StringWriter sw = new System.IO.StringWriter();
                dt.WriteXml(sw, System.Data.XmlWriteMode.IgnoreSchema, true);
                WriteFile(sw.ToString(), filePath);
            }
            catch (Exception ex)
            {
                throw ex;
            }

        }

        /// <summary>
        /// Method convert from file xml to datatable
        /// </summary>
        /// <param name="path">Path file xml</param>
        /// <returns>DataTable</returns>
        public static System.Data.DataTable ConvertXmlToDataTable(string path)
        {
            try
            {
                if (!File.Exists(path)) return null;
                System.Data.DataSet ds = new System.Data.DataSet();
                ds.ReadXml(path);
                return ds.Tables[0];
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// Method write string to file 
        /// </summary>
        /// <param name="content">String content</param>
        /// <param name="path">Path save file</param>
        public static void WriteFile(String content, String path)
        {
            //System.Web.Hosting.HostingEnvironment.ApplicationPhysicalPath + @"\tmp.txt"
            FileStream fs = new FileStream(path, FileMode.OpenOrCreate, FileAccess.Write);
            StreamWriter sw = new StreamWriter(fs);
            sw.BaseStream.Seek(0, SeekOrigin.End);
            sw.WriteLine(content);
            sw.Flush();
            sw.Close();
        }
        #endregion

        #region Matrix processing
        public static void HeaderDataBind(SAPbouiCOM.Application SBO_Application, SAPbouiCOM.Form Form, string TableName, string DocNum = "")
        {
            SAPbouiCOM.ComboBox ComboBox = default(SAPbouiCOM.ComboBox);
            SAPbouiCOM.EditText EditText = default(SAPbouiCOM.EditText);
            SAPbouiCOM.CheckBox CheckBox = default(SAPbouiCOM.CheckBox);
            try
            {
                for (int Count = 0; Count <= Form.Items.Count - 1; Count++)
                {
                    switch (Form.Items.Item(Count).Type)
                    {
                        case SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX:
                            try
                            {
                                Form.Items.Item(Count).DisplayDesc = true;
                                ComboBox = (SAPbouiCOM.ComboBox)Form.Items.Item(Count).Specific;
                                ComboBox.DataBind.SetBound(true, TableName, Form.Items.Item(Count).UniqueID);
                            }
                            catch { }
                            break;
                        case (SAPbouiCOM.BoFormItemTypes.it_EDIT):
                            try
                            {
                                EditText = (SAPbouiCOM.EditText)Form.Items.Item(Count).Specific;
                                EditText.DataBind.SetBound(true, TableName, Form.Items.Item(Count).UniqueID);
                            }
                            catch { }
                            break;
                        case SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX:
                            try
                            {
                                CheckBox = (SAPbouiCOM.CheckBox)Form.Items.Item(Count).Specific;
                                CheckBox.DataBind.SetBound(true, TableName, Form.Items.Item(Count).UniqueID);
                            }
                            catch { }
                            break;
                    }
                }

                if (!string.IsNullOrWhiteSpace(DocNum))
                    Form.DataBrowser.BrowseBy = DocNum;
                Form.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                Form.PaneLevel = 1;
            }
            catch (Exception ex)
            {
                SBO_Application.StatusBar.SetText(ex.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        public static void MatrixDataBind(SAPbouiCOM.Application SBO_Application, SAPbouiCOM.Form Form, string tableName, string Matrix_Name, bool AutoResize = false)
        {
            if (string.IsNullOrWhiteSpace(tableName)) return;
            SAPbouiCOM.Matrix Matrix = default(SAPbouiCOM.Matrix);
            try
            {
                Matrix = (SAPbouiCOM.Matrix)Form.Items.Item(Matrix_Name).Specific;
                for (int i = 1; i <= Matrix.Columns.Count - 1; i++)
                    if (!string.IsNullOrEmpty(Matrix.Columns.Item(i).Description))
                    {
                        try
                        {
                            Matrix.Columns.Item(i).DataBind.SetBound(true, tableName, Matrix.Columns.Item(i).Description);
                            if (Matrix.Columns.Item(i).Type == SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
                                Matrix.Columns.Item(i).DisplayDesc = true;
                            else
                                Matrix.Columns.Item(i).DisplayDesc = false;
                        }
                        catch { }
                    }
                if (AutoResize == true)
                    Matrix.AutoResizeColumns();
            }
            catch (Exception ex) { SBO_Application.StatusBar.SetText("MatrixDataBind: " + ex.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error); }
        }

        public static void AddRowMatrix(SAPbouiCOM.Application Aplication, string formUID, string matrixUID, string oDBDataSourceName, params string[] col_val)
        {
            SAPbouiCOM.Form oForm = null;
            SAPbouiCOM.Matrix oMatrix = default(SAPbouiCOM.Matrix);
            SAPbouiCOM.DBDataSource oDBDataSource = default(SAPbouiCOM.DBDataSource);
            try
            {
                oForm = Aplication.Forms.Item(formUID);
                oForm.Freeze(true);
                oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item(matrixUID).Specific;
                oDBDataSource = oForm.DataSources.DBDataSources.Item(oDBDataSourceName);
                if (oMatrix.RowCount > 0)
                {
                    oMatrix.FlushToDataSource();
                    oDBDataSource.InsertRecord(oDBDataSource.Size);
                }
                if (oDBDataSource.Size == 0)
                    oDBDataSource.InsertRecord(oDBDataSource.Size);
                if (col_val != null)
                    if (col_val.Length > 1)
                        for (int i = 0; i < col_val.Length; i += 2)
                            oDBDataSource.SetValue(col_val[i], oDBDataSource.Size - 1, col_val[i + 1]);
                oMatrix.LoadFromDataSource();
                if (oForm.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE & oForm.Mode != SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                oForm.Freeze(false);
            }
            catch (Exception ex)
            {
                if ((oForm != null))
                    oForm.Freeze(false);
                Aplication.StatusBar.SetText(ex.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        public static void DeleteRowMatrix(SAPbouiCOM.Application Aplication, string formUID, string matrixUID, string oDBDataSourceName, int rows)
        {
            SAPbouiCOM.Form oForm = null;
            SAPbouiCOM.Matrix oMatrix = default(SAPbouiCOM.Matrix);
            SAPbouiCOM.DBDataSource oDBDataSource = default(SAPbouiCOM.DBDataSource);
            try
            {
                oForm = Aplication.Forms.Item(formUID);
                oForm.Freeze(true);
                oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item(matrixUID).Specific;
                oDBDataSource = oForm.DataSources.DBDataSources.Item(oDBDataSourceName);

                oMatrix.FlushToDataSource();
                if (oMatrix.RowCount > 0)
                {
                    if (rows < oDBDataSource.Size)
                        oDBDataSource.RemoveRecord(rows);
                }
                oMatrix.LoadFromDataSource();
                if (oForm.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE & oForm.Mode != SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                oForm.Freeze(false);
            }
            catch (Exception ex)
            {
                if ((oForm != null))
                    oForm.Freeze(false);
                Aplication.StatusBar.SetText(ex.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }
        #endregion

        #region Choose List
        public static void SetChooseFormListColumn(SAPbouiCOM.Application SBO_Application, SAPbouiCOM.Form form, string Matrix_Name, string Column_Name, string UniqueID, string ObjectType, string Alias_Field_Name)
        {
            SAPbouiCOM.ChooseFromList oCFL1 = default(SAPbouiCOM.ChooseFromList);
            SAPbouiCOM.ChooseFromListCollection oCFLs1 = default(SAPbouiCOM.ChooseFromListCollection);
            SAPbouiCOM.ChooseFromListCreationParams oCFLCreationParams = default(SAPbouiCOM.ChooseFromListCreationParams);
            SAPbouiCOM.Column Column = default(SAPbouiCOM.Column);
            SAPbouiCOM.Matrix Matrix = default(SAPbouiCOM.Matrix);
            try
            {
                oCFLs1 = form.ChooseFromLists;
                Matrix = (SAPbouiCOM.Matrix)form.Items.Item(Matrix_Name).Specific;
                Column = Matrix.Columns.Item(Column_Name);
                if (Column.ChooseFromListUID.ToString() != UniqueID)
                {
                    oCFLCreationParams = (SAPbouiCOM.ChooseFromListCreationParams)SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams);
                    oCFLCreationParams.MultiSelection = false;
                    oCFLCreationParams.ObjectType = ObjectType;
                    oCFLCreationParams.UniqueID = UniqueID;
                    oCFL1 = oCFLs1.Add(oCFLCreationParams);
                    Column.ChooseFromListUID = UniqueID;
                    if (!string.IsNullOrEmpty(Alias_Field_Name.ToString()))
                        Column.ChooseFromListAlias = Alias_Field_Name;
                }
            }
            catch (Exception ex)
            {
                SBO_Application.StatusBar.SetText("SetChooseFormListColumn: " + ex.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        public static void SetChooseFormListToItem(SAPbouiCOM.Application SBO_Application, SAPbouiCOM.Form form, string Item_Name, string UniqueID, string ObjectType, string Alias_Field_Name)
        {
            SAPbouiCOM.ChooseFromList oCFL1 = default(SAPbouiCOM.ChooseFromList);
            SAPbouiCOM.ChooseFromListCollection oCFLs1 = default(SAPbouiCOM.ChooseFromListCollection);
            SAPbouiCOM.ChooseFromListCreationParams oCFLCreationParams = default(SAPbouiCOM.ChooseFromListCreationParams);
            SAPbouiCOM.EditText EditText = default(SAPbouiCOM.EditText);
            try
            {
                oCFLs1 = form.ChooseFromLists;
                EditText = (SAPbouiCOM.EditText)form.Items.Item(Item_Name).Specific;
                if (EditText.ChooseFromListUID.ToString() != UniqueID)
                {
                    oCFLCreationParams = (SAPbouiCOM.ChooseFromListCreationParams)SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams);
                    oCFLCreationParams.MultiSelection = false;
                    oCFLCreationParams.ObjectType = ObjectType;
                    oCFLCreationParams.UniqueID = UniqueID;
                    oCFL1 = oCFLs1.Add(oCFLCreationParams);
                    EditText.ChooseFromListUID = UniqueID;
                    if (!string.IsNullOrEmpty(Alias_Field_Name.ToString()))
                        EditText.ChooseFromListAlias = Alias_Field_Name;
                }
            }
            catch (Exception ex)
            {
                SBO_Application.StatusBar.SetText("SetChooseFormListToItem: " + ex.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        public static void FillDataChooseFormLisToColumnx(SAPbouiCOM.Application cApplication, SAPbouiCOM.ItemEvent pVal, string colAlias, params string[] colMap_ColVal)
        {
            SAPbouiCOM.DataTable DataTable = default(SAPbouiCOM.DataTable);
            SAPbouiCOM.Form oForm = default(SAPbouiCOM.Form);
            SAPbouiCOM.Matrix oMatrix = default(SAPbouiCOM.Matrix);
            SAPbouiCOM.IChooseFromListEvent MakerCFLEvnt = default(SAPbouiCOM.IChooseFromListEvent);
            try
            {
                oForm = cApplication.Forms.Item(pVal.FormUID);
                oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item(pVal.ItemUID).Specific;
                MakerCFLEvnt = (SAPbouiCOM.ChooseFromListEvent)pVal;
                if (MakerCFLEvnt.SelectedObjects == null)
                    return;
                DataTable = MakerCFLEvnt.SelectedObjects;
                if ((DataTable != null))
                {
                    if (colMap_ColVal != null)
                        if (colMap_ColVal.Length > 1)
                            for (int i = 0; i < colMap_ColVal.Length; i += 2)
                            {
                                try
                                {
                                    oMatrix.SetCellWithoutValidation(pVal.Row, colMap_ColVal[i], DataTable.GetValue(colMap_ColVal[i + 1], 0).ToString().Trim());
                                }
                                catch { }
                            }
                    try
                    {
                        oMatrix.SetCellWithoutValidation(pVal.Row, pVal.ColUID, DataTable.GetValue(colAlias, 0).ToString().Trim());
                    }
                    catch { }
                    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                }
                DataTable = null;
                oForm = null;
            }
            catch { }
        }
        public static void FillDataChooseFormLisToColumn(SAPbouiCOM.Application cApplication, SAPbouiCOM.ItemEvent pVal, string colAlias, params string[] colMap_ColVal)
        {
            SAPbouiCOM.DataTable DataTable = default(SAPbouiCOM.DataTable);
            SAPbouiCOM.Form oForm = default(SAPbouiCOM.Form);
            SAPbouiCOM.Matrix oMatrix = default(SAPbouiCOM.Matrix);
            SAPbouiCOM.IChooseFromListEvent MakerCFLEvnt = default(SAPbouiCOM.IChooseFromListEvent);
            try
            {
                oForm = cApplication.Forms.Item(pVal.FormUID);
                oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item(pVal.ItemUID).Specific;
                MakerCFLEvnt = (SAPbouiCOM.ChooseFromListEvent)pVal;
                if (MakerCFLEvnt.SelectedObjects == null)
                    return;
                DataTable = MakerCFLEvnt.SelectedObjects;
                if ((DataTable != null))
                {
                    if (colMap_ColVal != null)
                        if (colMap_ColVal.Length > 1)
                            for (int i = 0; i < colMap_ColVal.Length; i += 2)
                            {
                                try
                                {
                                    oMatrix.SetCellWithoutValidation(pVal.Row, colMap_ColVal[i], DataTable.GetValue(colMap_ColVal[i + 1], 0).ToString().Trim());
                                }
                                catch { }
                            }
                    try
                    {
                        oMatrix.SetCellWithoutValidation(pVal.Row, pVal.ColUID, DataTable.GetValue(colAlias, 0).ToString().Trim());
                        AddRowMatrix(cApplication, oForm.UniqueID, "38", "@PROMO_L", "U_ItemCode", "");

                    }
                    catch { }
                    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                }
                DataTable = null;
                oForm = null;
            }
            catch { }
        }
        public static void FillDataChooseFormLisToItem(SAPbouiCOM.Application SBO_Application, SAPbouiCOM.ItemEvent pVal, string itemMap = "", string col_AliasMap = "", bool formUpd = true)
        {
            SAPbouiCOM.DataTable DataTable = default(SAPbouiCOM.DataTable);
            SAPbouiCOM.IChooseFromListEvent MakerCFLEvnt = default(SAPbouiCOM.IChooseFromListEvent);
            try
            {
                MakerCFLEvnt = (SAPbouiCOM.ChooseFromListEvent)pVal;
                if (MakerCFLEvnt.SelectedObjects == null)
                    return;
                DataTable = MakerCFLEvnt.SelectedObjects;
                if (!DataTable.IsEmpty)
                {
                    SAPbouiCOM.Form oForm = SBO_Application.Forms.Item(pVal.FormUID);
                    SAPbouiCOM.EditText oEdit;
                    if (!string.IsNullOrWhiteSpace(itemMap) & !string.IsNullOrWhiteSpace(col_AliasMap))
                    {
                        try
                        {
                            oEdit = (SAPbouiCOM.EditText)oForm.Items.Item(itemMap).Specific;
                            oEdit.Value = DataTable.GetValue(col_AliasMap, 0).ToString();
                        }
                        catch { }
                    }
                    try
                    {
                        oEdit = (SAPbouiCOM.EditText)oForm.Items.Item(pVal.ItemUID).Specific;
                        oEdit.Value = DataTable.GetValue(oEdit.ChooseFromListAlias, 0).ToString();
                    }
                    catch { }
                    //if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE & formUpd)
                    //    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                }
            }
            catch (Exception ex)
            {
                SBO_Application.StatusBar.SetText("FillDataChooseFormLisToItem: " + ex.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }
        #endregion

        #region Other form controls
        public static void GetConditions(ref SAPbouiCOM.Conditions oConditions, SAPbouiCOM.Condition oCondition, string field, string val, SAPbouiCOM.BoConditionOperation operation, SAPbouiCOM.BoConditionRelationship relationship)
        {
            oCondition = oConditions.Add();
            //oCondition.BracketOpenNum = backetNumber
            oCondition.Alias = field;
            oCondition.Operation = operation;
            oCondition.CondVal = val;
            //oCondition.BracketCloseNum = backetNumber
            oCondition.Relationship = relationship;
        }

        public static void GetConditions(ref SAPbouiCOM.Conditions oConditions, string field, string val, SAPbouiCOM.BoConditionOperation operation, SAPbouiCOM.BoConditionRelationship relationship)
        {
            SAPbouiCOM.Condition oCondition = default(SAPbouiCOM.Condition);
            oCondition = oConditions.Add();
            //oCondition.BracketOpenNum = backetNumber
            oCondition.Alias = field;
            oCondition.Operation = operation;
            oCondition.CondVal = val;
            //oCondition.BracketCloseNum = backetNumber
            oCondition.Relationship = relationship;
        }

        public static void SetLinkButtonToColumn(SAPbouiCOM.Application SBO_Application, SAPbouiCOM.Form form, string item, string col, SAPbouiCOM.BoLinkedObject objType, string user_Object)
        {
            try
            {
                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)form.Items.Item(item).Specific;
                if (!string.IsNullOrWhiteSpace(user_Object))
                    ((SAPbouiCOM.LinkedButton)oMatrix.Columns.Item(col).ExtendedObject).LinkedObjectType = user_Object;
                else
                    ((SAPbouiCOM.LinkedButton)oMatrix.Columns.Item(col).ExtendedObject).LinkedObject = objType;
            }
            catch (Exception ex)
            {
                SBO_Application.StatusBar.SetText("SetLinkButtonToColumn: " + ex.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        #endregion

        public static System.Xml.XmlNode GetNodeRows(System.Xml.XmlDocument oXmlDoc)
        {
            System.Xml.XmlNode xmlNode = null;
            var matrixNodeName = "Matrix";
            var rowsNodeName = "Rows";
            for (var iA = 0; iA < oXmlDoc.ChildNodes.Count; iA++)
            {
                if (oXmlDoc.ChildNodes.Item(iA).Name == matrixNodeName)
                {
                    for (var iB = 0; iB < oXmlDoc.ChildNodes.Item(iA).ChildNodes.Count; iB++)
                    {
                        if (oXmlDoc.ChildNodes.Item(iA).ChildNodes.Item(iB).Name == rowsNodeName)
                        {
                            xmlNode = oXmlDoc.ChildNodes.Item(iA).ChildNodes.Item(iB);
                            break;
                        }
                    }
                    break;
                }
            }
            return xmlNode;
        }

        public static Matrix GetNetMatrix(SAPbouiCOM.Matrix oMatrix)
        {
            Matrix netMatrix = null;
            try
            {
                var xmlSerializer
                    = new System.Xml.Serialization.XmlSerializer(typeof(Matrix));
                // Read Matrix from UI API matrix object
                var sMatrix = oMatrix.SerializeAsXML(SAPbouiCOM.BoMatrixXmlSelect.mxs_All);
                var strReader = new StringReader(sMatrix);
                // Call the Deserialize method and cast to the object type
                netMatrix = (Matrix)xmlSerializer.Deserialize(strReader);
            }
            catch (Exception)
            {
                netMatrix = null;
            }
            return netMatrix;
        }

        public static List<ColumnsInfor> GetColumnsInforMatrixs(Matrix netMatrix)
        {
            List<ColumnsInfor> columnsInfors = new List<ColumnsInfor>();
            for (var i = 0; i < netMatrix.ColumnsInfo.Length; i++)
            {
                var dataBind = netMatrix.ColumnsInfo[i].DataBind;
                if (dataBind != null)
                {
                    var iD = netMatrix.ColumnsInfo[i].UniqueID;
                    //var tableName = dataBind.TableName;
                    var aliasName = dataBind.Alias;
                    columnsInfors.Add(new ColumnsInfor()
                    {
                        ID = iD,
                        //TableName = tableName,
                        Alias = aliasName,
                        IndexCloumn = i,
                    });
                }
            }
            return columnsInfors;
        }

    }
}
