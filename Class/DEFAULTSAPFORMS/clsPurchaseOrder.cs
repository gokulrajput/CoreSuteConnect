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
using CoreSuteConnect.Events;
using System.Data;
using System.Security.Cryptography;

namespace CoreSuteConnect.Class.DEFAULTSAPFORMS
{
    class clsPurchaseOrder
    {
        #region VariableDeclaration

        private SAPbouiCOM.Form oForm;
        CommonUtility objCU = new CommonUtility();

        SAPbouiCOM.ChooseFromList CFL_OITM;
         
        SAPbouiCOM.ChooseFromListCollection oCFLs = null;
        SAPbouiCOM.Conditions oCons = null;
        SAPbouiCOM.Condition oCon = null;
        public string JDEntry = null;

        #endregion VariableDeclaration

        public bool MenuEvent(ref SAPbouiCOM.MenuEvent pVal, string FormId)
        {
            bool BubbleEvent = true;
            try
            {

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

                                        string getQuery = @"SELECT AcctCode FROM OACT WHERE ExportCode = 'JOBWORK'";
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
                            btn.Item.Left = oItem.Left + oItem.Width + 7;
                            btn.Item.Width = oItem.Width + 10;
                            btn.Item.Height = oItem.Height;
                            btn.Item.Enabled = true;
                            btn.Caption = "EXIM Tracking";
                            #endregion EXIM TRACKING Button

                            
                            #region Jobwork Out Button
                            SAPbouiCOM.Item oItem1 = (SAPbouiCOM.Item)oForm.Items.Item("16");  /// Existing Item on the form of Cancel Button
                            SAPbouiCOM.Button btn1 = (SAPbouiCOM.Button)oForm.Items.Add("btnJWO", SAPbouiCOM.BoFormItemTypes.it_BUTTON).Specific;
                            btn1.Item.Top = oItem1.Top;
                            btn1.Item.Left = oItem1.Left + (oItem1.Width) + 17;
                            btn1.Item.Width = oItem1.Width + 10;
                            btn1.Item.Height = 20;
                            btn1.Item.Enabled = true;
                            btn1.Caption = "JobWork Out";

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
                                    SBOMain.SBO_Application.StatusBar.SetText("Purchase order is not opened.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                }
                                else
                                {
                                    OutwardToEximTracking outEximTracking = new OutwardToEximTracking();
                                    //string DocNum = string.Empty; string Series = string.Empty;
                                    SAPbouiCOM.EditText CardCode = oForm.Items.Item("4").Specific;
                                    SAPbouiCOM.EditText CardName = oForm.Items.Item("54").Specific;
                                    SAPbouiCOM.EditText DocNum = oForm.Items.Item("8").Specific;
                                    SAPbouiCOM.ComboBox Series = oForm.Items.Item("88").Specific;

                                    string getDocEntry = "SELECT DocEntry FROM OPOR WHERE DocNum='" + DocNum.Value + "' AND Series='" + Series.Selected.Value + "' AND CardCode='" + CardCode.Value + "'";
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
                                            outEximTracking.FromFrmName = "PurchaseOrderEXIST";

                                            objCU.FormLoadAndActivate("frmETTrans", "mnsmEXIM013");
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
                                            outEximTracking.FromFrmName = "PurchaseOrder";
                                            outEximTracking.BPName = CardName.Value;

                                            /*SAPbouiCOM.Form oUDFForm = SBOMain.SBO_Application.Forms.Item(oForm.UDFFormUID);
                                            outEximTracking.Incoterm = oUDFForm.Items.Item("U_EXIMINCO").Specific.value;
                                            outEximTracking.PrecarriageBy = oUDFForm.Items.Item("U_EXIMPCGB").Specific.value;
                                            outEximTracking.PrecarrierBy = oUDFForm.Items.Item("U_EXIMPCRB").Specific.value;
                                            outEximTracking.Portofloading = oUDFForm.Items.Item("U_EXIMPOL").Specific.value;
                                            outEximTracking.Portlfdischarge = oUDFForm.Items.Item("U_EXIMPOD").Specific.Value;
                                            outEximTracking.Portofreceipt = oUDFForm.Items.Item("U_EXIMPOR").Specific.Value;
                                            outEximTracking.Finaldestination = oUDFForm.Items.Item("U_EXIMFD").Specific.Value;
                                            outEximTracking.Countryoforigin = oUDFForm.Items.Item("U_EXIMOC").Specific.Value;
                                            outEximTracking.DestinationCountry = oUDFForm.Items.Item("U_EXIMFDC").Specific.Value;
*/
                                            objCU.FormLoadAndActivate("frmETTrans", "mnsmEXIM013");
                                            clsExTrans oPrice = new clsExTrans(outEximTracking);
                                            oForm.Close();
                                        }
                                    }
                                }
                            }
                             
                            //FOR JOBWORK ADDON
                            if (pVal.ItemUID == "btnJWO" && (oForm.Mode == BoFormMode.fm_OK_MODE))
                            {
                                string status = oForm.Items.Item("81").Specific.Value.ToString(); 
                                SAPbouiCOM.ComboBox DocType = oForm.Items.Item("3").Specific;
                                string dt =  DocType.Selected.Value.ToString();

                                if (dt != "S")
                                {
                                    SBOMain.SBO_Application.StatusBar.SetText("Purchase order should be only service type.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                }
                                else if (status == "3" || status == "4")
                                {
                                    SBOMain.SBO_Application.StatusBar.SetText("Purchase order should not canceled or closed.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                }
                                else
                                {

                                    string abc = JDEntry;
                                    OutwardToJWO OutwardToJWO = new OutwardToJWO();

                                    string q1 = "SELECT DocEntry,U_PoDe FROM [dbo].[@JOTR] WHERE U_PoDe='" + abc + "'";
                                    SAPbobsCOM.Recordset recq1 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                    recq1.DoQuery(q1);
                                    if (recq1.RecordCount > 0)
                                    {
                                        OutwardToJWO.FromFrmName = "POExist";
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
                                        string getDocEntry = "SELECT DocEntry FROM OPOR WHERE DocNum='" + DocNum.Value + "' AND Series='" + Series.Selected.Value + "' AND CardCode='" + CardCode.Value + "'";
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
                                            OutwardToJWO.FromFrmName = "PurchaseOrder";

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
                                    objCU.FormLoadAndActivate("frmJWout", "mnsmJW001");
                                    clsJWOut oPriceNew = new clsJWOut(OutwardToJWO);
                                    oForm.Close();
                                }
                            }
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
                                    /*string getQuery = @"SELECT AcctCode FROM OACT WHERE ExportCode = 'JOBWORK'";
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
                                    }  */
                                }
                                if (pVal.ItemUID == "39" && (pVal.ColUID == "U_Qty_CSJW" || pVal.ColUID == "U_Price_CSJW"))
                                {
                                    double qty = Convert.ToDouble((oMatrix.Columns.Item("U_Qty_CSJW").Cells.Item(pVal.Row).Specific).Value);
                                    double price = Convert.ToDouble((oMatrix.Columns.Item("U_Price_CSJW").Cells.Item(pVal.Row).Specific).Value); 
                                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("5").Cells.Item(pVal.Row).Specific).Value = Convert.ToString( qty * price );
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

                if (BusinessObjectInfo.ActionSuccess == true)
                {
                    if (!string.IsNullOrEmpty(SBOMain.TFromUID) && SBOMain.FromCnt != null)
                    {
                        
                        SBOMain.GetDocEntryFromXml(BusinessObjectInfo.ObjectKey, ref DocEntry); 
                        if (SBOMain.sForm == "LC")
                        {
                            SAPbouiCOM.Form LCTran = SBOMain.SBO_Application.Forms.GetForm(SBOMain.TFromUID, Convert.ToInt32(SBOMain.FromCnt));
                            SAPbouiCOM.Matrix oMatrix = LCTran.Items.Item("matLCEX").Specific;

                            string docnum = objCU.getDocNumFromDocKey("OPOR", DocEntry);
                            string currency = objCU.getCurrFromDocKey("OPOR", DocEntry);
                            double lineTotal = objCU.getLineTotalFromDocKey("POR1", DocEntry);
                            double rate = objCU.getRateFromDocKey("POR1", DocEntry);

                            ((SAPbouiCOM.EditText)oMatrix.Columns.Item("lc5bdn").Cells.Item(SBOMain.LineNo).Specific).Value = docnum;
                            ((SAPbouiCOM.EditText)oMatrix.Columns.Item("lc5bden").Cells.Item(SBOMain.LineNo).Specific).Value = DocEntry.ToString(); //oDataTable.GetValue("DocDate", 0).ToString();
                            ((SAPbouiCOM.EditText)oMatrix.Columns.Item("lc5pl").Cells.Item(SBOMain.LineNo).Specific).Value = lineTotal.ToString();
                            ((SAPbouiCOM.EditText)oMatrix.Columns.Item("lc5cur").Cells.Item(SBOMain.LineNo).Specific).Value = currency;
                            ((SAPbouiCOM.EditText)oMatrix.Columns.Item("lc5rt").Cells.Item(SBOMain.LineNo).Specific).Value = rate.ToString();
                        }
                        else if (SBOMain.sForm == "ET")
                        {
                            SAPbouiCOM.Form ExTrans = SBOMain.SBO_Application.Forms.GetForm(SBOMain.TFromUID, Convert.ToInt32(SBOMain.FromCnt));
                            SAPbouiCOM.Matrix oMatrix = ExTrans.Items.Item("matEXPAC").Specific;

                            string docnum = objCU.getDocNumFromDocKey("OPOR", DocEntry);
                            string currency = objCU.getCurrFromDocKey("OPOR", DocEntry);
                            double lineTotal = objCU.getLineTotalFromDocKey("POR1", DocEntry);
                            double rate = objCU.getRateFromDocKey("POR1", DocEntry);

                            ((SAPbouiCOM.EditText)oMatrix.Columns.Item("et3edn").Cells.Item(SBOMain.LineNo).Specific).Value = docnum;
                            ((SAPbouiCOM.EditText)oMatrix.Columns.Item("et3ede").Cells.Item(SBOMain.LineNo).Specific).Value = DocEntry.ToString(); //oDataTable.GetValue("DocDate", 0).ToString();
                            ((SAPbouiCOM.EditText)oMatrix.Columns.Item("et3pl").Cells.Item(SBOMain.LineNo).Specific).Value = lineTotal.ToString();
                            ((SAPbouiCOM.EditText)oMatrix.Columns.Item("et3cur").Cells.Item(SBOMain.LineNo).Specific).Value = currency;
                            ((SAPbouiCOM.EditText)oMatrix.Columns.Item("et3rt").Cells.Item(SBOMain.LineNo).Specific).Value = rate.ToString();
                        }

                        SBOMain.TFromUID = "";
                        SBOMain.FromCnt = null;
                        SBOMain.LineNo = null;
                        SBOMain.sForm = "";
                    }
                }
            }
            catch (Exception ex)
            {

            }
        }
    }


}
