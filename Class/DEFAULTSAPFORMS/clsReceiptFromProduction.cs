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

namespace CoreSuteConnect.Class.DEFAULTSAPFORMS
{
    class clsReceiptFromProduction
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
                        oForm = SBOMain.SBO_Application.Forms.ActiveForm;

                        if (pVal.BeforeAction == true)
                        {

                        }
                        if (pVal.BeforeAction == false)
                        {

                        }
                        break;

                    case BoEventTypes.et_DOUBLE_CLICK:
                        oForm = SBOMain.SBO_Application.Forms.ActiveForm;
                        if (pVal.BeforeAction == true)
                        {
                            if (pVal.ItemUID == "13" && (oForm.Mode == BoFormMode.fm_OK_MODE))
                            {
                                SAPbouiCOM.EditText postingDate = (SAPbouiCOM.EditText)oForm.Items.Item("9").Specific;
                                DateTime d2 = Convert.ToDateTime(postingDate.String);
                                DateTime d1 = new DateTime(2022, 04, 01, 0, 0, 0);
                                int res = DateTime.Compare(d1, d2);
                                if (res == 1) {
                                    SBOMain.SBO_Application.StatusBar.SetText("Select Receipt From Production from current Finacial Year", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                    BubbleEvent = false;
                                }
                            } 
                        }
                        if (pVal.BeforeAction == false)
                        {
                            oForm = SBOMain.SBO_Application.Forms.ActiveForm;
                            if (pVal.ItemUID == "13" && (oForm.Mode == BoFormMode.fm_OK_MODE))
                            {
                                string itemCode = oForm.Items.Item("1000008").Specific.Value;
                                string itemnm = oForm.Items.Item("1000012").Specific.Value;
                                double qty = Convert.ToDouble(oForm.Items.Item("1000024").Specific.Value);
                                string whscode = oForm.Items.Item("1000036").Specific.Value;

                                oForm.Close(); // Row Form
                                oForm = SBOMain.SBO_Application.Forms.ActiveForm;
                                OutwardToQC OutwardToQC = new OutwardToQC();
                                SAPbouiCOM.EditText CardCode = oForm.Items.Item("U_BP_Name").Specific;
                                //SAPbouiCOM.EditText CardName = oForm.Items.Item("54").Specific;
                                SAPbouiCOM.EditText DocNum = oForm.Items.Item("7").Specific;
                                SAPbouiCOM.ComboBox Series = oForm.Items.Item("30").Specific;

                                OutwardToQC.FormName = "RFPNew";
                                OutwardToQC.BPName = "";
                                OutwardToQC.BPCode = CardCode.Value;
                                OutwardToQC.DocNum = DocNum.Value;

                                OutwardToQC.ItemCode = itemCode;
                                OutwardToQC.ItemName = itemnm;
                                OutwardToQC.Qty = qty;
                                OutwardToQC.Whs = whscode;
                                OutwardToQC.ItemGroup = "";
                                OutwardToQC.Batchno = "";

                                string getDocEntry = "";

                                if (!string.IsNullOrEmpty(CardCode.Value))
                                {
                                    getDocEntry = "SELECT DocEntry FROM OIGN WHERE DocNum='" + DocNum.Value + "' AND Series='" + Series.Selected.Value + "' AND U_BP_NAME='" + CardCode.Value + "'";
                                }
                                else
                                {
                                    getDocEntry = "SELECT DocEntry FROM OIGN WHERE DocNum='" + DocNum.Value + "' AND Series='" + Series.Selected.Value + "' ";
                                }
                                SAPbobsCOM.Recordset rec = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                rec.DoQuery(getDocEntry);
                                if (rec.RecordCount > 0)
                                {
                                    string DocEntry = Convert.ToString(rec.Fields.Item("DocEntry").Value);
                                    OutwardToQC.DocEntry = DocEntry;

                                    string getCardName = "Select CardName from OCRD where CardCode = '" + CardCode.Value + "'";
                                    SAPbobsCOM.Recordset rec1 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                    rec1.DoQuery(getCardName);
                                    if (rec1.RecordCount > 0)
                                    {
                                        OutwardToQC.BPName = Convert.ToString(rec1.Fields.Item("CardName").Value);
                                    }
                                      
                                    string getQC = "Select DocEntry from dbo.[@QCSR] where U_ItemCode = '" + itemCode + "'  AND";
                                    getQC = getQC + " U_DocType = 'RFP' AND U_BaseDocEnt = '" + DocEntry + "' and U_BaseDocNo = '" + DocNum.Value + "' ";
                                    SAPbobsCOM.Recordset rec2 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                    rec2.DoQuery(getQC);
                                    if (rec2.RecordCount > 0)
                                    {
                                        OutwardToQC.QCDocEntry = Convert.ToString(rec2.Fields.Item("DocEntry").Value);
                                        OutwardToQC.FormName = "RFPExist";
                                    }
                                    else
                                    {
                                        string getBatch = "Select T0.BatchNum From IBT1 T0 Inner Join OPDN T1 On T0.BaseType=T1.ObjType ";
                                        getBatch = getBatch + " and T0.BaseEntry = T1.DocEntry Where T0.BaseType = 20 and T0.ItemCode = '" + itemCode + "' AND T1.DocEntry = '" + DocEntry + "'";
                                        SAPbobsCOM.Recordset rec3 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                        rec3.DoQuery(getBatch);
                                        if (rec3.RecordCount > 0){
                                            OutwardToQC.Batchno = Convert.ToString(rec3.Fields.Item("BatchNum").Value);
                                        }
                                        string getItemGrp = "select T1.ItmsGrpNam from OITM as T0 LEFT JOIN OITB AS T1 ON T1.ItmsGrpCod = T0.ItmsGrpCod where T0.ItemCode = '" + itemCode + "'";
                                        SAPbobsCOM.Recordset rec4 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                        rec4.DoQuery(getItemGrp);
                                        if (rec4.RecordCount > 0){
                                            OutwardToQC.ItemGroup = Convert.ToString(rec4.Fields.Item("ItmsGrpNam").Value);
                                        }
                                    }

                                    objCU.FormLoadAndActivate("frmQCSample", "mnsmQC007"); 
                                    clsSampleRequest oSR = new clsSampleRequest(OutwardToQC);
                                    oForm.Close();
                                }
                                else
                                {
                                    SBOMain.SBO_Application.StatusBar.SetText("Goods receipt is not saved.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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