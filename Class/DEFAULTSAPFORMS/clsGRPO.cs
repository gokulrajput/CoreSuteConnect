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
using CoreSuteConnect.Class.QC;
using System.Security.AccessControl;

namespace CoreSuteConnect.Class.DEFAULTSAPFORMS
{
    class clsGRPO
    {
        #region VariableDeclaration
        private SAPbouiCOM.Form oForm;
        public string cFormID = string.Empty;

        public static string getSalesForm = string.Empty;

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
                        }
                        break;
                        
                    case BoEventTypes.et_DOUBLE_CLICK:
                        oForm = SBOMain.SBO_Application.Forms.ActiveForm;
                        if (pVal.BeforeAction == true)
                        {
                            if (pVal.ItemUID == "38" && (oForm.Mode == BoFormMode.fm_OK_MODE))
                            {
                                SAPbouiCOM.EditText postingDate = (SAPbouiCOM.EditText)oForm.Items.Item("46").Specific;
                                DateTime d2 = Convert.ToDateTime(postingDate.String);
                                DateTime d1 = new DateTime(2022, 04, 01, 0, 0, 0);
                                int res = DateTime.Compare(d1, d2);
                                if (res == 1)
                                {
                                    SBOMain.SBO_Application.StatusBar.SetText("Select GRPO from current Finacial Year", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                    BubbleEvent = false;
                                }
                            }
                        } 
                        if (pVal.BeforeAction == false)
                        {
                            oForm = SBOMain.SBO_Application.Forms.ActiveForm;
                            if (pVal.ItemUID == "38" && (oForm.Mode == BoFormMode.fm_OK_MODE))
                            { 
                                string itemCode = oForm.Items.Item("1000056").Specific.Value;
                                string itemnm = oForm.Items.Item("1000064").Specific.Value;
                                double qty =  Convert.ToDouble(oForm.Items.Item("1000080").Specific.Value);
                                string whscode = oForm.Items.Item("1000098").Specific.Value;

                                oForm.Close(); // Row Form

                                oForm = SBOMain.SBO_Application.Forms.ActiveForm;
                                SAPbouiCOM.Matrix matItem = (SAPbouiCOM.Matrix)oForm.Items.Item("38").Specific;

                                OutwardToQC OutwardToQC = new OutwardToQC();
                                SAPbouiCOM.EditText CardCode = oForm.Items.Item("4").Specific;
                                SAPbouiCOM.EditText CardName = oForm.Items.Item("54").Specific;
                                SAPbouiCOM.EditText DocNum = oForm.Items.Item("8").Specific;
                                SAPbouiCOM.EditText NumAtCard = oForm.Items.Item("14").Specific;
                                SAPbouiCOM.ComboBox Series = oForm.Items.Item("88").Specific;

                                OutwardToQC.FormName = "GRPONew";
                                OutwardToQC.BPName = CardName.Value;
                                OutwardToQC.BPCode = CardCode.Value;
                                OutwardToQC.DocNum = DocNum.Value;
                                OutwardToQC.NumAtCard = NumAtCard.Value;

                                OutwardToQC.ItemCode = itemCode;
                                OutwardToQC.ItemName = itemnm;
                                OutwardToQC.Qty = qty;
                                OutwardToQC.Whs = whscode; 
                                OutwardToQC.ItemGroup = "";
                                OutwardToQC.Batchno = "";

                                string getDocEntry = "";
 
                                if (!string.IsNullOrEmpty(CardCode.Value))
                                {
                                    getDocEntry = "SELECT DocEntry FROM OPDN WHERE DocNum='" + DocNum.Value + "' AND Series='" + Series.Selected.Value + "' AND CardCode='" + CardCode.Value + "'";
                                }
                                else
                                {
                                    getDocEntry = "SELECT DocEntry FROM OPDN WHERE DocNum='" + DocNum.Value + "' AND Series='" + Series.Selected.Value + "' ";
                                } 
                                // string getDocEntry = "SELECT DocEntry FROM OPDN WHERE DocNum='" + DocNum.Value + "' AND Series='" + Series.Selected.Value + "' AND CardCode='" + CardCode.Value + "'";
                                SAPbobsCOM.Recordset rec = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                rec.DoQuery(getDocEntry);
                                if (rec.RecordCount > 0)
                                {
                                    string DocEntry = Convert.ToString(rec.Fields.Item("DocEntry").Value);
                                    OutwardToQC.DocEntry = DocEntry;  

                                     string getQC = "Select DocEntry from dbo.[@QCSR] where U_ItemCode = '" + itemCode + "'  AND";
                                     getQC = getQC + " U_DocType = 'GRPO' AND U_BaseDocEnt = '" + DocEntry + "' and U_BaseDocNo = '" + DocNum.Value + "' ";
                                    SAPbobsCOM.Recordset rec2 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                    rec2.DoQuery(getQC);
                                    if (rec2.RecordCount > 0)
                                    {
                                         OutwardToQC.QCDocEntry = Convert.ToString(rec2.Fields.Item("DocEntry").Value); 
                                         OutwardToQC.FormName = "GRPOExist";
                                    }
                                    else
                                    {
                                        string getBatch = "Select T0.BatchNum From IBT1 T0 Inner Join OPDN T1 On T0.BaseType=T1.ObjType ";
                                               getBatch = getBatch + " and T0.BaseEntry = T1.DocEntry Where T0.BaseType = 20 and T0.ItemCode = '"+ itemCode + "' AND T1.DocEntry = '"+ DocEntry + "'";
                                        SAPbobsCOM.Recordset rec3 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                        rec3.DoQuery(getBatch);
                                        if (rec3.RecordCount > 0)
                                        {
                                            OutwardToQC.Batchno = Convert.ToString(rec3.Fields.Item("BatchNum").Value); 
                                        }

                                        string getItemGrp = "select T1.ItmsGrpNam from OITM as T0 LEFT JOIN OITB AS T1 ON T1.ItmsGrpCod = T0.ItmsGrpCod where T0.ItemCode = '"+ itemCode + "'";
                                        SAPbobsCOM.Recordset rec4 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                        rec4.DoQuery(getItemGrp);
                                        if (rec4.RecordCount > 0)
                                        {
                                            OutwardToQC.ItemGroup = Convert.ToString(rec4.Fields.Item("ItmsGrpNam").Value);
                                        } 
                                    }

                                    objCU.FormLoadAndActivate("frmQCSample", "mnsmQC007"); 
                                    clsSampleRequest oSR = new clsSampleRequest(OutwardToQC);
                                    oForm.Close();
                                }
                                else
                                {
                                    SBOMain.SBO_Application.StatusBar.SetText("GRPO is not saved.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                }
                            }
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

                                /* if (status == "1")
                                 {
                                     SBOMain.SBO_Application.StatusBar.SetText("Delivery is not opened.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                 }
                                 else
                                 {*/

                                OutwardToEximTracking outEximTracking = new OutwardToEximTracking(); 
                                SAPbouiCOM.EditText CardCode = oForm.Items.Item("4").Specific;
                                SAPbouiCOM.EditText CardName = oForm.Items.Item("54").Specific;
                                SAPbouiCOM.EditText DocNum = oForm.Items.Item("8").Specific;
                                SAPbouiCOM.ComboBox Series = oForm.Items.Item("88").Specific;

                                string getDocEntry = "SELECT DocEntry FROM OPDN WHERE DocNum='" + DocNum.Value + "' AND Series='" + Series.Selected.Value + "' AND CardCode='" + CardCode.Value + "'";
                                SAPbobsCOM.Recordset rec = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                rec.DoQuery(getDocEntry);
                                if (rec.RecordCount > 0)
                                {
                                    string DocEntry = Convert.ToString(rec.Fields.Item("DocEntry").Value);

                                    string getDocEntry1 = "SELECT DocNum FROM dbo.[@EXET] WHERE U_exdnde = '" + DocEntry + "' AND U_exbc='" + CardCode.Value + "'";
                                    SAPbobsCOM.Recordset rec1 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                    rec1.DoQuery(getDocEntry1);
                                    if (rec1.RecordCount > 0)
                                    {
                                        outEximTracking.DelDocEnt = DocEntry;
                                        outEximTracking.FromFrmName = "GRPOEXIST";
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

                                        // SBOMain.SBO_Application.StatusBar.SetText("Delivery is linked with Exim Transaction no: '" + Convert.ToString(rec1.Fields.Item("DocNum").Value) + "'.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                    }
                                    else
                                    {
                                        outEximTracking.BPName = CardName.Value;
                                        outEximTracking.BPCode = CardCode.Value;
                                        outEximTracking.DelDocNo = DocNum.Value;
                                        outEximTracking.DelDocEnt = DocEntry;
                                        outEximTracking.FromFrmName = "GRPO"; 
                                         
                                        string getBaseEntry16 = "SELECT DISTINCT(BaseEntry) FROM PDN1 WHERE DocEntry = '" + DocEntry + "' AND BaseType = '22'";
                                        SAPbobsCOM.Recordset rec16 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                        rec16.DoQuery(getBaseEntry16);
                                        if (rec16.RecordCount > 0)
                                        {
                                            outEximTracking.SODocEnt = Convert.ToString(rec16.Fields.Item("BaseEntry").Value);

                                            string getBaseEntry20 = "SELECT DocNum FROM OPOR WHERE  DocEntry = '" + rec16.Fields.Item("BaseEntry").Value + "'";
                                            SAPbobsCOM.Recordset rec20 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                            rec20.DoQuery(getBaseEntry20);
                                            if (rec20.RecordCount > 0)
                                            {
                                                outEximTracking.SODocNo = Convert.ToString(rec20.Fields.Item("DocNum").Value);
                                            }
                                        } 

                                        objCU.FormLoadAndActivate("frmETTrans", "mnsmEXIM013"); 
                                        clsExTrans oPrice = new clsExTrans(outEximTracking);
                                        oForm.Close();
                                    }
                                }  
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
    }
}
