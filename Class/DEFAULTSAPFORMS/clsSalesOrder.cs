using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbouiCOM;
using SAPbobsCOM;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
using System.Threading;
using System.Runtime.InteropServices;
using CoreSuteConnect.Class.EXIM;
using CoreSuteConnect.Class.JOBWORK;

namespace CoreSuteConnect.Class.DEFAULTSAPFORMS
{
    class clsSalesOrder
    {
        #region VariableDeclaration
        private SAPbouiCOM.Form oForm;
        public string cFormID = string.Empty;

        public static string getSalesForm = string.Empty;

        public string JDEntry = null;

        SAPbouiCOM.ChooseFromListCollection oCFLs = null;
        SAPbouiCOM.Conditions oCons = null;
        SAPbouiCOM.Condition oCon = null;
        CommonUtility objCU = new CommonUtility();

        #endregion VariableDeclaration

        public bool MenuEvent(ref SAPbouiCOM.MenuEvent pVal, string FormId)
        {
            bool BubbleEvent = true;
            try
            {
                cFormID = FormId;
                oForm = SBOMain.SBO_Application.Forms.Item(FormId);
                if (oForm.Mode == BoFormMode.fm_ADD_MODE)
                {


                }
            }
            catch (Exception ex)
            { 

            }
            finally
            {

            }

            return BubbleEvent;
        }

        public bool ItemEvent(ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent, string FormId)
        {
            BubbleEvent = true;
            oForm = SBOMain.SBO_Application.Forms.Item(FormId);
            try
            {
                switch (pVal.EventType)
                { // KEY Down event for multiplication


                    case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
                        //oForm = SBOMain.SBO_Application.Forms.ActiveForm;
                        if (pVal.BeforeAction == true)
                        {

                        }
                        if (pVal.BeforeAction == false)
                        {
                            SAPbouiCOM.IChooseFromListEvent oCFLEvento = null;
                            oCFLEvento = ((SAPbouiCOM.IChooseFromListEvent)(pVal));
                            string sCFL_ID = oCFLEvento.ChooseFromListUID;
                            SAPbouiCOM.ChooseFromList oCFL = null;
                            oCFL = oForm.ChooseFromLists.Item(sCFL_ID);
                            SAPbouiCOM.DataTable oDataTable = null;
                            oDataTable = oCFLEvento.SelectedObjects;
                            if (oDataTable != null)
                            {
                                try
                                {
                                    if (pVal.ItemUID == "39" && pVal.ColUID == "U_ItemCode_CSJW")
                                    {
                                        SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("39").Specific;

                                        ((SAPbouiCOM.EditText)oMatrix.Columns.Item("U_ItemCode_CSJW").Cells.Item(pVal.Row).Specific).Value = oDataTable.GetValue("ItemCode", 0).ToString();
                                        ((SAPbouiCOM.EditText)oMatrix.Columns.Item("U_Desc_CSJW").Cells.Item(pVal.Row).Specific).Value = oDataTable.GetValue("ItemName", 0).ToString();
                                        ((SAPbouiCOM.EditText)oMatrix.Columns.Item("U_UOM_CSJW").Cells.Item(pVal.Row).Specific).Value = oDataTable.GetValue("SalUnitMsr", 0).ToString();

                                        string getQuery = @"SELECT AcctCode FROM OACT WHERE ExportCode = 'JOBWORK-In'";
                                        SAPbobsCOM.Recordset rec;
                                        rec = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                        rec.DoQuery(getQuery);
                                        if (rec.RecordCount > 0)
                                        {
                                            ((SAPbouiCOM.EditText)oMatrix.Columns.Item("2").Cells.Item(pVal.Row).Specific).Value = Convert.ToString(rec.Fields.Item("AcctCode").Value);
                                        }
                                    }
                                }
                                catch (Exception ex)
                                { 

                                }
                            }
                        }
                        break;

                    case SAPbouiCOM.BoEventTypes.et_FORM_LOAD:
                        // oForm = SBOMain.SBO_Application.Forms.ActiveForm;
                        oForm = SBOMain.SBO_Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);

                        if (pVal.BeforeAction == true)
                        {

                        }
                        if (pVal.BeforeAction == false)
                        {
                            SAPbouiCOM.Item oItem = (SAPbouiCOM.Item)oForm.Items.Item("2");  /// Existing Item on the form of Cancel Button

                            #region EXIM TRACKING Button
                            SAPbouiCOM.Button btn = (SAPbouiCOM.Button)oForm.Items.Add("btnExTrk", SAPbouiCOM.BoFormItemTypes.it_BUTTON).Specific;
                            btn.Item.Top = oItem.Top;
                            btn.Item.Left = oItem.Left + oItem.Width +   7;
                            btn.Item.Width = oItem.Width + 10;
                            btn.Item.Height = oItem.Height;
                            btn.Item.Enabled = true;
                            btn.Caption = "EXIM Tracking";
                            #endregion EXIM TRACKING Button


                            /*************** FOR JOBWORK Addon ********************/
                            #region Jobwork In Button
                            SAPbouiCOM.Item oItem1 = (SAPbouiCOM.Item)oForm.Items.Item("2");  /// Existing Item on the form of Cancel Button
                            SAPbouiCOM.Button btn1 = (SAPbouiCOM.Button)oForm.Items.Add("btnJWI", SAPbouiCOM.BoFormItemTypes.it_BUTTON).Specific;
                            btn1.Item.Top = oItem1.Top - 25;
                            btn1.Item.Left = oItem1.Left + (oItem1.Width) + 17;
                            btn1.Item.Width = oItem1.Width + 10;
                            btn1.Item.Height = 20;
                            btn1.Item.Enabled = true;
                            btn1.Caption = "JobWork In";


                            /****** jobwork Addon Choose from list in item code */

                            oCFLs = oForm.ChooseFromLists;
                            SAPbouiCOM.ChooseFromList oCFL = null;
                            SAPbouiCOM.ChooseFromList oCFL1 = null;
                            SAPbouiCOM.ChooseFromListCreationParams oCFLCreationParams = null;
                            oCFLCreationParams = ((SAPbouiCOM.ChooseFromListCreationParams)(SBOMain.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)));

                            oCFLCreationParams.MultiSelection = false;
                            oCFLCreationParams.ObjectType = "4";
                            oCFLCreationParams.UniqueID = "CFL1";
                            oCFL = oCFLs.Add(oCFLCreationParams);
                            SAPbouiCOM.Matrix lineMatrix = oForm.Items.Item("39").Specific;
                            oCFL1 = oForm.ChooseFromLists.Item("CFL1");
                            SAPbouiCOM.Column oCol = lineMatrix.Columns.Item("U_ItemCode_CSJW");
                            oCol.ChooseFromListUID = "CFL1";
                            oCol.ChooseFromListAlias = "ItemCode";
                            /****** jobwork Addon Choose from list in item code */

                            #endregion Jobwork Out Button
                            /*************** FOR JOBWORK Addon ********************/

                        }
                        break;

                    case BoEventTypes.et_CLICK:

                        oForm = SBOMain.SBO_Application.Forms.ActiveForm;
                        if (pVal.BeforeAction == true)
                        {

                        }
                        if (pVal.BeforeAction == false)
                        {
                            if (pVal.ItemUID == "btnExTrk" && (oForm.Mode == BoFormMode.fm_OK_MODE))
                            {
                                string status = oForm.Items.Item("81").Specific.Value.ToString();

                                if (status != "1")
                                {
                                    SBOMain.SBO_Application.StatusBar.SetText("Sales order is not opened.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                }
                                else
                                {
                                    OutwardToEximTracking outEximTracking = new OutwardToEximTracking();
                                    //string DocNum = string.Empty; string Series = string.Empty;
                                    SAPbouiCOM.EditText CardCode = oForm.Items.Item("4").Specific;
                                    SAPbouiCOM.EditText CardName = oForm.Items.Item("54").Specific;
                                    SAPbouiCOM.EditText DocNum = oForm.Items.Item("8").Specific;
                                    SAPbouiCOM.ComboBox Series = oForm.Items.Item("88").Specific;

                                    string getDocEntry = "SELECT DocEntry FROM ORDR WHERE DocNum='" + DocNum.Value + "' AND Series='" + Series.Selected.Value + "' AND CardCode='" + CardCode.Value + "'";
                                    SAPbobsCOM.Recordset rec = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                    rec.DoQuery(getDocEntry);
                                    if (rec.RecordCount > 0)
                                    {
                                        string DocEntry = Convert.ToString(rec.Fields.Item("DocEntry").Value);

                                        string getDocEntry1 = "SELECT DocNum FROM dbo.[@EXET] WHERE U_exsonde = '" + DocEntry + "' AND U_exbc='" + CardCode.Value + "'";
                                        SAPbobsCOM.Recordset rec1 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                        rec1.DoQuery(getDocEntry1);
                                        if (rec1.RecordCount > 0)
                                        {
                                            outEximTracking.SODocEnt = DocEntry;
                                            outEximTracking.FromFrmName = "SalesOrderEXIST";
                                            bool plFormOpen = false;
                                            for (int i = 1; i < SBOMain.SBO_Application.Forms.Count; i++)
                                            {
                                                if (SBOMain.SBO_Application.Forms.Item(i).UniqueID == "frmETTrans")
                                                {
                                                    SBOMain.SBO_Application.Forms.Item(i).Select();
                                                    plFormOpen = true;
                                                }
                                            }
                                            if (!plFormOpen)
                                            {
                                                SBOMain.SBO_Application.Menus.Item("mnsmEXIM013").Activate();
                                            }
                                            clsExTrans oPrice = new clsExTrans(outEximTracking);
                                            oForm.Close();

                                            //SBOMain.SBO_Application.StatusBar.SetText("Sales order is linked with Exim Transaction no: '"+ Convert.ToString(rec1.Fields.Item("DocNum").Value) + "'.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                        }
                                        else
                                        {
                                            outEximTracking.BPName = CardName.Value;
                                            outEximTracking.BPCode = CardCode.Value;
                                            outEximTracking.SODocNo = DocNum.Value;
                                            outEximTracking.SODocEnt = DocEntry;
                                            outEximTracking.FromFrmName = "SalesOrder";
                                            outEximTracking.BPName = CardName.Value;

                                            SAPbouiCOM.Form oUDFForm = SBOMain.SBO_Application.Forms.Item(oForm.UDFFormUID);
                                            outEximTracking.Incoterm = oUDFForm.Items.Item("U_EXIMINCO").Specific.value;
                                            outEximTracking.PrecarriageBy = oUDFForm.Items.Item("U_EXIMPCGB").Specific.value;
                                            outEximTracking.PrecarrierBy = oUDFForm.Items.Item("U_EXIMPCRB").Specific.value;
                                            outEximTracking.Portofloading = oUDFForm.Items.Item("U_EXIMPOL").Specific.value;
                                            outEximTracking.Portlfdischarge = oUDFForm.Items.Item("U_EXIMPOD").Specific.Value;
                                            outEximTracking.Portofreceipt = oUDFForm.Items.Item("U_EXIMPOR").Specific.Value;
                                            outEximTracking.Finaldestination = oUDFForm.Items.Item("U_EXIMFD").Specific.Value;
                                            outEximTracking.Countryoforigin = oUDFForm.Items.Item("U_EXIMOC").Specific.Value;
                                            outEximTracking.DestinationCountry = oUDFForm.Items.Item("U_EXIMFDC").Specific.Value;

                                            bool plFormOpen = false;
                                            for (int i = 1; i < SBOMain.SBO_Application.Forms.Count; i++)
                                            {
                                                if (SBOMain.SBO_Application.Forms.Item(i).UniqueID == "frmETTrans")
                                                {
                                                    SBOMain.SBO_Application.Forms.Item(i).Select();
                                                    plFormOpen = true;
                                                }
                                            }
                                            if (!plFormOpen)
                                            {
                                                SBOMain.SBO_Application.Menus.Item("mnsmEXIM013").Activate();
                                            }
                                            clsExTrans oPrice = new clsExTrans(outEximTracking);
                                            oForm.Close();
                                        }
                                    }
                                }
                            }
                            //FOR JOBWORK ADDON
                            if (pVal.ItemUID == "btnJWI" && (oForm.Mode == BoFormMode.fm_OK_MODE))
                            {
                                string status = oForm.Items.Item("81").Specific.Value.ToString();
                                SAPbouiCOM.ComboBox DocType = oForm.Items.Item("3").Specific;
                                string dt = DocType.Selected.Value.ToString();
                                if (dt != "S")
                                {
                                    SBOMain.SBO_Application.StatusBar.SetText("Sales order should be only service type.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                }
                                else if (status == "3" || status == "4")
                                {
                                    SBOMain.SBO_Application.StatusBar.SetText("Sales order should not canceled or closed.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                }
                                else
                                {
                                    string abc = JDEntry;
                                    OutwardToJWO OutwardToJWO = new OutwardToJWO();

                                    string q1 = "SELECT DocEntry, U_SoDe FROM [dbo].[@JITR] WHERE U_SoDe='" + abc + "'";
                                    SAPbobsCOM.Recordset recq1 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                    recq1.DoQuery(q1);
                                    if (recq1.RecordCount > 0)
                                    {
                                        OutwardToJWO.FromFrmName = "SOExist";
                                        OutwardToJWO.DocEntry = Convert.ToString(recq1.Fields.Item("DocEntry").Value);
                                    }
                                    else
                                    {
                                        //string DocNum = string.Empty; string Series = string.Empty;
                                        SAPbouiCOM.EditText CardCode = oForm.Items.Item("4").Specific;
                                        SAPbouiCOM.EditText CardName = oForm.Items.Item("54").Specific;
                                        SAPbouiCOM.EditText DocNum = oForm.Items.Item("8").Specific;
                                        SAPbouiCOM.ComboBox Series = oForm.Items.Item("88").Specific;
                                        SAPbouiCOM.EditText VendorReference = oForm.Items.Item("14").Specific;

                                        string getDefWhs = "SELECT U_DefWhs FROM OCRD WHERE CardCode='" + CardCode.Value + "'";
                                        SAPbobsCOM.Recordset rec1 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                        rec1.DoQuery(getDefWhs);
                                        if (rec1.RecordCount > 0)
                                        {
                                            OutwardToJWO.DefWhs = Convert.ToString(rec1.Fields.Item("U_DefWhs").Value);
                                            rec1.MoveNext();
                                        }
                                        string getDocEntry = "SELECT DocEntry FROM ORDR WHERE DocNum='" + DocNum.Value + "' AND Series='" + Series.Selected.Value + "' AND CardCode='" + CardCode.Value + "'";
                                        SAPbobsCOM.Recordset rec = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                        rec.DoQuery(getDocEntry);
                                        if (rec.RecordCount > 0)
                                        {
                                            OutwardToJWO.DocEntry = Convert.ToString(rec.Fields.Item("DocEntry").Value);
                                            OutwardToJWO.BPName = CardName.Value;
                                            OutwardToJWO.BPCode = CardCode.Value;
                                            OutwardToJWO.DocNum = DocNum.Value;
                                            OutwardToJWO.VendorRef = VendorReference.Value;
                                            //OutwardToJWO.Series = Series.Value;
                                            OutwardToJWO.FromFrmName = "SalesOrder";

                                            List<OutwardToJWOItems> oItemList = new List<OutwardToJWOItems>();
                                            SAPbouiCOM.Matrix oMatrix = oForm.Items.Item("39").Specific;
                                            if (oMatrix.RowCount > 0)
                                            {
                                                for (int i = 1; i <= oMatrix.RowCount; i++)
                                                {
                                                    OutwardToJWOItems oItems = new OutwardToJWOItems();
                                                    if (!string.IsNullOrEmpty(Convert.ToString(((SAPbouiCOM.EditText)oMatrix.Columns.Item("U_ItemCode_CSJW").Cells.Item(i).Specific).Value)))
                                                    {
                                                        oItems.ItemCode = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("U_ItemCode_CSJW").Cells.Item(i).Specific).Value;
                                                        oItems.ItemName = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("U_Desc_CSJW").Cells.Item(i).Specific).Value;
                                                        oItems.Quantity = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("U_Qty_CSJW").Cells.Item(i).Specific).Value;
                                                        oItems.UnitPrice = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("U_Price_CSJW").Cells.Item(i).Specific).Value;
                                                        oItems.UOM = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("U_UOM_CSJW").Cells.Item(i).Specific).Value;
                                                        oItemList.Add(oItems);
                                                    }
                                                }
                                            }
                                            OutwardToJWO.lstItems = oItemList; 
                                        }
                                    }

                                    objCU.FormLoadAndActivate("frmJWIn", "mnsmJW002");
                                    clsJWIn oPriceNew = new clsJWIn(OutwardToJWO);
                                    oForm.Close();

                                }
                            }
                        }

                        break;

                    case BoEventTypes.et_GOT_FOCUS:

                        oForm = SBOMain.SBO_Application.Forms.ActiveForm;
                        if (pVal.BeforeAction == true)
                        {

                        }
                        if (pVal.BeforeAction == false)
                        {
                           
                        }
                        break;

                    case BoEventTypes.et_LOST_FOCUS:

                        oForm = SBOMain.SBO_Application.Forms.ActiveForm;
                        if (pVal.BeforeAction == true)
                        {

                        }
                        if (pVal.BeforeAction == false)
                        {
                            try
                            {
                                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("39").Specific;
                                if (pVal.ItemUID == "39" && pVal.ColUID == "U_ItemCode_CSJW")
                                {
                                    /*string getQuery = @"SELECT AcctCode FROM OACT WHERE ExportCode = 'JOBWORK-In'";
                                    SAPbobsCOM.Recordset rec;
                                    rec = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                    rec.DoQuery(getQuery);
                                    if (rec.RecordCount > 0)
                                    {
                                        while (!rec.EoF)
                                        {
                                            ((SAPbouiCOM.EditText)oMatrix.Columns.Item("2").Cells.Item(pVal.Row).Specific).Value = Convert.ToString(rec.Fields.Item("AcctCode").Value);
                                            rec.MoveNext();
                                        }
                                    }*/
                                }
                                if (pVal.ItemUID == "39" && (pVal.ColUID == "U_Qty_CSJW" || pVal.ColUID == "U_Price_CSJW"))
                                {
                                    double qty = Convert.ToDouble((oMatrix.Columns.Item("U_Qty_CSJW").Cells.Item(pVal.Row).Specific).Value);
                                    double price = Convert.ToDouble((oMatrix.Columns.Item("U_Price_CSJW").Cells.Item(pVal.Row).Specific).Value);
                                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("5").Cells.Item(pVal.Row).Specific).Value = Convert.ToString(qty * price);
                                }
                            }
                            catch (Exception ex)
                            {

                            }
                        }
                        break;
                }
            }
            catch (Exception ex)
            {

            }

            return BubbleEvent;
        }

        public void SBO_Application_FormDataEvent(ref BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                string DocEntry = "";
                SBOMain.GetDocEntryFromXml(BusinessObjectInfo.ObjectKey, ref DocEntry);
                JDEntry = DocEntry; 
            }
            catch (Exception ex)
            {

            }
        }
    }

    public class OutwardToLCMaster
    {
        public string lcno { get; set; }
    }
    public class OutwardToIncoMaster
    {
        public string inctno { get; set; }
    }
    public class OutwardToPortMaster
    {
        public string portcode { get; set; }
    }
    public class OutwardToNPMaster
    {
        public string processcode { get; set; }
    }
    public class OutwardToEXPMaster
    {
        public string itemcode { get; set; }
    }
    public class OutwardToSchemeMaster
    {
        public string schmeno { get; set; }
    }
    public class OutwardFromEximTracking
    {
        public string ScriptNo { get; set; }
    }
    public class OutwardToQC
    {
        public string ItemCode { get; set; }
        public string ItemName { get; set; }
        public double Qty { get; set; }
        public string Whs { get; set; }
        public string QCDocEntry { get; set; }
        public string DocNum { get; set; }
        public string DocEntry { get; set; }
        public string FormName { get; set; }
        public string BPCode { get; set; }
        public string BPName { get; set; }
        public string Batchno { get; set; }
        public string ItemGroup { get; set; }
        public string NumAtCard { get; set; }
        public string InOutNo { get; set; }
    }
    public class OutwardToEximTracking
    {
        public string BPCode { get; set; }
        public string BPName { get; set; }
        public string DocNum { get; set; }
        public string DocDate { get; set; }
        public string DocEntry { get; set; }
        public string FromFrmName { get; set; }
        public string DelDocNo { get; set; }
        public string DelDocEnt { get; set; }
        public string SODocNo { get; set; }
        public string SODocEnt { get; set; }

        public string Incoterm { get; set; }
        public string PrecarriageBy { get; set; }
        public string PrecarrierBy { get; set; }
        public string Portofloading { get; set; }
        public string Portlfdischarge { get; set; }
        public string Portofreceipt { get; set; }
        public string Finaldestination { get; set; }
        public string Countryoforigin { get; set; }
        public string DestinationCountry { get; set; }

    }
    public class OutwardToJWO
    {
        public string BPCode { get; set; }
        public string BPName { get; set; }
        public string DocNum { get; set; }
        public string DocDate { get; set; }
        public string DefWhs { get; set; }
        public string DocEntry { get; set; }
        public string FromFrmName { get; set; }
        public string VendorRef { get; set; }
        //public string Series { get; set; }
        public List<OutwardToJWOItems> lstItems { get; set; }
    }
    public class OutwardToJWOItems
    {
        public string ItemCode { get; set; }
        public string ItemName { get; set; }
        public string Quantity { get; set; } 
        public string UnitPrice { get; set; }
        public string UOM { get; set; }
    }
    public class OutwardToJWPM
    {
        public string NopCode{ get; set; }
    }
}
