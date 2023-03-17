using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM;
using SAPbobsCOM;
using System.Threading.Tasks;
using CoreSuteConnect.Events;
using CoreSuteConnect.Class.DEFAULTSAPFORMS;
using System.Collections;
using System.Windows.Forms;
using CoreSuteConnect.Class.EXIM;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Button;
using System.IO;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Data;
using System.Diagnostics;

namespace CoreSuteConnect.Class.QC
{
    class clsSampleRequest
    {
        #region VariableDeclaration
        private SAPbouiCOM.Form oForm;
       
        SAPbouiCOM.ComboBox cb1, cb4;
        SAPbouiCOM.LinkedButton oLinkBaseDoc = null;

        SAPbouiCOM.EditText tDocNo;
        SAPbouiCOM.ChooseFromList CFL_GRPO, CFL_GI, CFL_GR;

        private SAPbouiCOM.ChooseFromList oCFL = null;
        private SAPbouiCOM.Conditions oConds = null;
        private SAPbouiCOM.Condition oCond = null;

        #endregion VariableDeclaration

        CommonUtility objCU = new CommonUtility();

        public clsSampleRequest(OutwardToQC outClass)
        { 
            try
            {
                if (outClass != null)
                {
                    oForm = SBOMain.SBO_Application.Forms.ActiveForm; 
                    if (outClass.FormName == "GRPONew")
                    {
                        assignFormValues(oForm, outClass, "GRPO"); 
                        setChooseFromListField(oForm, CFL_GRPO, "CFL_GRPO", tDocNo, "tDocNo", SAPbouiCOM.BoLinkedObject.lf_GoodsReceiptPO);
                    }
                    else if (outClass.FormName == "GRPOExist")
                    {
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                        oForm.Items.Item("tCode").Specific.value = outClass.QCDocEntry;
                        oForm.Items.Item("1").Click(); 
                        setChooseFromListField(oForm, CFL_GRPO, "CFL_GRPO", tDocNo, "tDocNo", SAPbouiCOM.BoLinkedObject.lf_GoodsReceiptPO);
                    }
                    else if (outClass.FormName == "GINew")
                    {
                        assignFormValues(oForm, outClass, "GI"); 
                        setChooseFromListField(oForm, CFL_GI, "CFL_GI", tDocNo, "tDocNo", SAPbouiCOM.BoLinkedObject.lf_GoodsIssue);
                    }
                    else if (outClass.FormName == "GIExist")
                    {
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                        oForm.Items.Item("tCode").Specific.value = outClass.QCDocEntry;
                        oForm.Items.Item("1").Click();
                        setChooseFromListField(oForm, CFL_GI, "CFL_GI", tDocNo, "tDocNo", SAPbouiCOM.BoLinkedObject.lf_GoodsIssue);
                    }
                    else if (outClass.FormName == "GRNew")
                    { 
                        assignFormValues(oForm, outClass, "GR"); 
                        setChooseFromListField(oForm, CFL_GR, "CFL_GR", tDocNo, "tDocNo", SAPbouiCOM.BoLinkedObject.lf_GoodsReceipt);
                    }
                    else if (outClass.FormName == "GRExist")
                    {
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                        oForm.Items.Item("tCode").Specific.value = outClass.QCDocEntry;
                        oForm.Items.Item("1").Click();
                        setChooseFromListField(oForm, CFL_GR, "CFL_GR", tDocNo, "tDocNo", SAPbouiCOM.BoLinkedObject.lf_GoodsReceipt);
                    }
                    else if (outClass.FormName == "RFPNew")
                    {
                        assignFormValues(oForm, outClass, "RFP");
                        setChooseFromListField(oForm, CFL_GR, "CFL_GR", tDocNo, "tDocNo", SAPbouiCOM.BoLinkedObject.lf_GoodsReceipt);
                    }
                    else if (outClass.FormName == "RFPExist")
                    {
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                        oForm.Items.Item("tCode").Specific.value = outClass.QCDocEntry;
                        oForm.Items.Item("1").Click();
                        setChooseFromListField(oForm, CFL_GR, "CFL_GR", tDocNo, "tDocNo", SAPbouiCOM.BoLinkedObject.lf_GoodsReceipt);
                    }
                }
            }
            catch (Exception ex)
            {
                // SBOMain.SBO_Application.MessageBox(ex.Message.ToString(), 1, "ok", "", "");
            }
            finally
            {

            }
        }

        public void assignFormValues(SAPbouiCOM.Form oForm, OutwardToQC outClass, string frm)
        {
            oForm.Freeze(true);
            cb1 = (SAPbouiCOM.ComboBox)oForm.Items.Item("tDocType").Specific;
            cb1.ExpandType = BoExpandType.et_DescriptionOnly;
            cb1.Select(frm);
            oForm.Items.Item("tDocEnt").Specific.value = outClass.DocEntry;
            oForm.Items.Item("tDocNo").Specific.value = outClass.DocNum;
            oForm.Items.Item("tItemCode").Specific.value = outClass.ItemCode;
            oForm.Items.Item("tDesc").Specific.value = outClass.ItemName;
            oForm.Items.Item("tBatchNo").Specific.value = outClass.Batchno;
            oForm.Items.Item("tWhs").Specific.value = outClass.Whs;
            oForm.Items.Item("tQty").Specific.value = outClass.Qty;  
            oForm.Items.Item("tItemGrp").Specific.value = outClass.ItemGroup; 
            oForm.Items.Item("tCardCode").Specific.value = outClass.BPCode;
            oForm.Items.Item("tCardName").Specific.value = outClass.BPName;
            oForm.Items.Item("tRefNo").Specific.value = outClass.NumAtCard;
            oForm.Items.Item("tInOutNo").Specific.value = outClass.InOutNo;
            oForm.Freeze(false); 
        }

        public bool MenuEvent(ref SAPbouiCOM.MenuEvent pVal, string FormId, string Type = "")
        {
            bool BubbleEvent = true;
            try
            {   
                oForm = SBOMain.SBO_Application.Forms.Item(FormId);
                if (pVal.BeforeAction == true)
                {

                }
                if (pVal.BeforeAction == false)
                {
                    SAPbouiCOM.Matrix matContent = (SAPbouiCOM.Matrix)oForm.Items.Item("matContent").Specific;
                    if ((oForm.Mode == BoFormMode.fm_ADD_MODE || oForm.Mode == BoFormMode.fm_UPDATE_MODE || Type == "ADDNEWFORM") && (Type != "DEL_ROW") && (Type != "ADD_ROW"))
                    { 
                        Form_Load_Components(oForm, "ADD"); 
                    }
                    
                    else if (Type == "navigation")
                    {
                        cb1 = oForm.Items.Item("tDocType").Specific;
                        string DocType = cb1.Selected.Value.ToString();

                        if (DocType == "GRPO")
                        {
                            setChooseFromListField(oForm, CFL_GRPO, "CFL_GRPO", tDocNo, "tDocNo", SAPbouiCOM.BoLinkedObject.lf_GoodsReceiptPO);
                        }
                        else if (DocType == "GI")
                        {
                            setChooseFromListField(oForm, CFL_GI, "CFL_GI", tDocNo, "tDocNo", SAPbouiCOM.BoLinkedObject.lf_GoodsIssue);
                        }
                        else if (DocType == "GR")
                        {
                            setChooseFromListField(oForm, CFL_GR, "CFL_GR", tDocNo, "tDocNo", SAPbouiCOM.BoLinkedObject.lf_GoodsReceipt);
                        }
                        else if (DocType == "RFP")
                        {
                            setChooseFromListField(oForm, CFL_GR, "CFL_GR", tDocNo, "tDocNo", SAPbouiCOM.BoLinkedObject.lf_GoodsReceipt);
                        }
                    }
                    else if (Type == "DEL_ROW")
                    {
                        DeleteMatrixBlankRow(matContent, "tParamCode");
                        ArrengeMatrixLineNum(matContent);
                    } 
                    else if (Type == "ADD_ROW")
                    {
                        ADDROWMain(matContent);
                    }
                    else if (Type == "previewLayout")
                    {
                         
                    }
                }

            }
            catch (Exception ex)
            {

                // SBOMain.SBO_Application.MessageBox(ex.Message.ToString(), 1, "ok", "", "");
            }
            finally
            {

            }

            return BubbleEvent;

        }
       /* private bool LoadCrViewer(CRAXDDRT.Report crxReport)
        {
            SAPbouiCOM.FormCreationParams SBOFormCreationParams;
            SAPbouiCOM.ActiveX SBOCRViewer;
            SAPbouiCOM.Form SBOForm;
            SAPbouiCOM.Item SBOItem;

            string strFormCount;

            SBOFormCreationParams = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
            SBOFormCreationParams.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Sizable;
            SBOFormCreationParams.FormType = "XX_INCRPT01";
            strFormCount = SBO_Application.Forms.Count.ToString;
            SBOFormCreationParams.UniqueID = "XX_INCRPT01" + strFormCount.PadLeft(3, "0");

            // Add new form
            SBOForm = SBO_Application.Forms.AddEx(SBOFormCreationParams);
            SBOForm.Left = 0;
            SBOForm.Top = 0;
            SBOForm.Width = SBO_Application.Desktop.Width;
            SBOForm.Height = SBO_Application.Desktop.Height;
            SBOForm.Title = "inCentea - Mapas";

            // Add CRViewer item
            SBOItem = SBOForm.Items.Add("XX_CR01", SAPbouiCOM.BoFormItemTypes.it_ACTIVE_X);
            SBOItem.Left = 0;
            SBOItem.Top = 0;
            SBOItem.Width = SBOForm.ClientWidth;
            SBOItem.Height = SBOForm.ClientHeight;

            // Create the new activeX control
            SBOCRViewer = SBOItem.Specific;

            SBOCRViewer.ClassID = "CrystalReports13.ActiveXReportViewer.1";

            var SBOCRViewerOBJ;
            SBOCRViewerOBJ = SBOCRViewer.Object;
            SBOCRViewerOBJ.EnablePrintButton = false;

            Process[] MyProcs;
            int i, ID;
            System.IntPtr a;

            // Try Send the handle of SAP window
            SBO_Application.Desktop.Title = "SBO under " + SBO_Application.Company.UserName;
            MyProcs = Process.GetProcessesByName("SAP Business One");
            for (i = 0; i <= MyProcs.Length - 1; i++)
            {
                if (MyProcs[i].MainWindowTitle == SBO_Application.Desktop.Title)
                {
                    ID = MyProcs[i].Id();
                    a = MyProcs[i].MainWindowHandle;
                    crxReport.SetDialogParentWindow(a.ToInt32());
                }
            }

            SBOCRViewerOBJ.ViewReport();

            SBOForm.Visible = true;

            return true;
        }*/
        private void DeleteMatrixBlankRow(SAPbouiCOM.Matrix oMatrix, string ColUID)
        {
            try
            {
                for (int i = oMatrix.VisualRowCount; i >= 1; i--)
                {
                    {
                        if (string.IsNullOrEmpty(((SAPbouiCOM.EditText)oMatrix.Columns.Item(ColUID).Cells.Item(i).Specific).Value))
                        {
                            oMatrix.DeleteRow(i);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
            }
        }
        public void ADDROWMain(SAPbouiCOM.Matrix oMatrix)
        {
            oMatrix.AddRow(1, SBOMain.RightClickLineNum);
            oMatrix.ClearRowData(SBOMain.RightClickLineNum + 1);
            ArrengeMatrixLineNum(oMatrix);
        }

        private void setChooseFromListField(SAPbouiCOM.Form oForm, SAPbouiCOM.ChooseFromList CFL, string clfname, SAPbouiCOM.EditText editext, string fieldname, SAPbouiCOM.BoLinkedObject linkobj)
        {  
            oLinkBaseDoc = oForm.Items.Item("lbBaseDoc").Specific;
            oLinkBaseDoc.LinkedObject = linkobj;
            CFL = oForm.ChooseFromLists.Item(clfname);
            editext = oForm.Items.Item(fieldname).Specific;
            editext.ChooseFromListUID = clfname;
            editext.ChooseFromListAlias = "DocNum";

            if(clfname != "CFL_GRPO")
            { 
                oForm.Items.Item("Item_3").Specific.Caption = "In/Out Ref No.";
            }else
            {
                oForm.Items.Item("Item_3").Specific.Caption = "Equi Ref No.";
            }
        }

        public bool ItemEvent(ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent, string FormId)
        {
            BubbleEvent = true;
            oForm = SBOMain.SBO_Application.Forms.Item(FormId);
            try
            {
                switch (pVal.EventType)
                { // KEY Down event for multiplication
                     
                    case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
                        if (pVal.BeforeAction == true)
                        {
                        }
                        if (pVal.BeforeAction == false)
                        {
                            if (pVal.ItemUID == "tDocType")
                            {
                                cb4 = oForm.Items.Item("tDocType").Specific;
                                string DocType = cb4.Selected.Value.ToString();
                                
                                if (DocType == "GRPO"){ 
                                    setChooseFromListField(oForm, CFL_GRPO, "CFL_GRPO", tDocNo, "tDocNo", SAPbouiCOM.BoLinkedObject.lf_GoodsReceiptPO);
                                }
                                else if (DocType == "GI"){
                                    setChooseFromListField(oForm, CFL_GI, "CFL_GI", tDocNo, "tDocNo", SAPbouiCOM.BoLinkedObject.lf_GoodsIssue);
                                }
                                else if (DocType == "GR"){
                                    setChooseFromListField(oForm, CFL_GR, "CFL_GR", tDocNo, "tDocNo", SAPbouiCOM.BoLinkedObject.lf_GoodsReceipt);
                                } 
                                else if (DocType == "RFP"){
                                    setChooseFromListField(oForm, CFL_GR, "CFL_GR", tDocNo, "tDocNo", SAPbouiCOM.BoLinkedObject.lf_GoodsReceipt);
                                }
                                oForm.Items.Item("tDocNo").Specific.value = "";
                                oForm.Items.Item("tDocEnt").Specific.value = ""; 
                            }
                        }
                        break;

                    case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
                        if (pVal.BeforeAction == false)
                        {
                            if (pVal.CharPressed == 9)
                            {
                                if (pVal.ItemUID == "matContent" && pVal.ColUID == "tParamCode")
                                {
                                    SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("matContent").Specific;
                                    AddMatrixRow(oMatrix, "tParamCode");
                                }
                            } 
                        }
                        break;

                    case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
                        // SAPbouiCOM.Form oForms = SBOMain.SBO_Application.Forms.ActiveForm;
                        if (pVal.BeforeAction == true)
                        {
                            if (pVal.ItemUID == "tCardCode")
                            {
                                CFLConditionBP("CFL_OCRD", pVal.ItemUID);
                            }
                            if (pVal.ItemUID == "tDocNo")
                            {
                                cb4 = oForm.Items.Item("tDocType").Specific;
                                string DocType = cb4.Selected.Value.ToString();
                                if (DocType == "GRPO")
                                {
                                    CFLConditionGRPO("CFL_GRPO", pVal.ItemUID);
                                }
                                if (DocType == "GR")
                                {
                                    CFLConditionGR("CFL_GR", pVal.ItemUID);
                                }
                                if (DocType == "GI")
                                {
                                    CFLConditionGI("CFL_GI", pVal.ItemUID);
                                }
                                if (DocType == "RFP")
                                {
                                    CFLConditionRFP("CFL_GR", pVal.ItemUID);
                                }
                            }
                            if (pVal.ItemUID == "tItemCode")
                            {
                                cb4 = oForm.Items.Item("tDocType").Specific;
                                string DocType = cb4.Selected.Value.ToString();
                                if (DocType == "GRPO")
                                {
                                    CFLConditionOITMGRPO("CFL_OITM", pVal.ItemUID);
                                }
                                else if (DocType == "GI")
                                {
                                    CFLConditionOITMGI("CFL_OITM", pVal.ItemUID);
                                }
                                else if (DocType == "GR")
                                {
                                    CFLConditionOITMGR("CFL_OITM", pVal.ItemUID);
                                }
                                else if (DocType == "RFP")
                                {
                                    CFLConditionOITMRFP("CFL_OITM", pVal.ItemUID);
                                }
                            } 
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
                                    if (pVal.ItemUID == "tCardCode")
                                    {
                                        oForm.Items.Item("tCardCode").Specific.value = oDataTable.GetValue("CardCode", 0).ToString();
                                        oForm.Items.Item("tCardName").Specific.value = oDataTable.GetValue("CardName", 0).ToString();
                                    }
                                    if (pVal.ItemUID == "tDocNo")
                                    {
                                        oForm.Items.Item("tDocNo").Specific.value = oDataTable.GetValue("DocNum", 0).ToString();
                                        oForm.Items.Item("tDocEnt").Specific.value = oDataTable.GetValue("DocEntry", 0).ToString();

                                        cb4 = oForm.Items.Item("tDocType").Specific;
                                        string DocType = cb4.Selected.Value.ToString();
                                        if (DocType == "GR" || DocType == "GI")
                                        {
                                            oForm.Items.Item("tInOutNo").Specific.value = oDataTable.GetValue("U_InOutRef", 0).ToString();
                                            oForm.Items.Item("tRefNo").Specific.value = oDataTable.GetValue("U_PartyRef", 0).ToString();
                                        }
                                    }
                                    if (pVal.ItemUID == "tItemCode")
                                    {
                                        oForm.Items.Item("tItemCode").Specific.value = oDataTable.GetValue("ItemCode", 0).ToString();
                                        oForm.Items.Item("tDesc").Specific.value = oDataTable.GetValue("ItemName", 0).ToString();

                                        cb4 = oForm.Items.Item("tDocType").Specific;
                                        string DocType = cb4.Selected.Value.ToString();
                                        if (DocType == "GRPO")
                                        {

                                            string getBatch = "Select T0.BatchNum From IBT1 T0 Inner Join OPDN T1 On T0.BaseType=T1.ObjType ";
                                            getBatch = getBatch + " and T0.BaseEntry = T1.DocEntry Where T0.BaseType = 20 and T0.ItemCode = '" + oDataTable.GetValue("ItemCode", 0).ToString() + "' AND T1.DocEntry = '" + oForm.Items.Item("tDocEnt").Specific.value + "'";
                                            SAPbobsCOM.Recordset rec3 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                            rec3.DoQuery(getBatch);
                                            if (rec3.RecordCount > 0)
                                            {
                                                oForm.Items.Item("tBatchNo").Specific.value = Convert.ToString(rec3.Fields.Item("BatchNum").Value);
                                            }
                                            string getItemGrp = "select T1.ItmsGrpNam from OITM as T0 LEFT JOIN OITB AS T1 ON T1.ItmsGrpCod = T0.ItmsGrpCod where T0.ItemCode = '" + oDataTable.GetValue("ItemCode", 0).ToString() + "'";
                                            SAPbobsCOM.Recordset rec4 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                            rec4.DoQuery(getItemGrp);
                                            if (rec4.RecordCount > 0)
                                            {
                                                oForm.Items.Item("tItemGrp").Specific.value = Convert.ToString(rec4.Fields.Item("ItmsGrpNam").Value);
                                            }
                                            string getwhsqty = "SELECT T0.Quantity, T0.WhsCode, T1.CardCode, T1.CardName, T1.NumAtCard from PDN1 AS T0 LEFt JOIN OPDN AS T1 ON T0.DocEntry = T1.DocEntry where T0.DocEntry =  '" + oForm.Items.Item("tDocEnt").Specific.value + "'";
                                            SAPbobsCOM.Recordset rec5 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                            rec5.DoQuery(getwhsqty);
                                            if (rec5.RecordCount > 0)
                                            {
                                                oForm.Items.Item("tQty").Specific.value = Convert.ToDouble(rec5.Fields.Item("Quantity").Value);
                                                oForm.Items.Item("tWhs").Specific.value = Convert.ToString(rec5.Fields.Item("WhsCode").Value); 
                                                oForm.Items.Item("tCardName").Specific.value = Convert.ToString(rec5.Fields.Item("CardName").Value);
                                                oForm.Items.Item("tRefNo").Specific.value = Convert.ToString(rec5.Fields.Item("NumAtCard").Value);
                                                oForm.Items.Item("tCardCode").Specific.value = Convert.ToString(rec5.Fields.Item("CardCode").Value);
                                            }
                                        }
                                        if (DocType == "GI")
                                        { 
                                            string getBatch = "Select T0.BatchNum From IBT1 T0 Inner Join OIGE T1 On T0.BaseType=T1.ObjType ";
                                            getBatch = getBatch + " and T0.BaseEntry = T1.DocEntry Where T0.BaseType = 60 and T0.ItemCode = '" + oDataTable.GetValue("ItemCode", 0).ToString() + "' AND T1.DocEntry = '" + oForm.Items.Item("tDocEnt").Specific.value + "'";
                                            SAPbobsCOM.Recordset rec3 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                            rec3.DoQuery(getBatch);
                                            if (rec3.RecordCount > 0)
                                            {
                                                oForm.Items.Item("tBatchNo").Specific.value = Convert.ToString(rec3.Fields.Item("BatchNum").Value);
                                            }
                                            string getItemGrp = "select T1.ItmsGrpNam from OITM as T0 LEFT JOIN OITB AS T1 ON T1.ItmsGrpCod = T0.ItmsGrpCod where T0.ItemCode = '" + oDataTable.GetValue("ItemCode", 0).ToString() + "'";
                                            SAPbobsCOM.Recordset rec4 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                            rec4.DoQuery(getItemGrp);
                                            if (rec4.RecordCount > 0)
                                            {
                                                oForm.Items.Item("tItemGrp").Specific.value = Convert.ToString(rec4.Fields.Item("ItmsGrpNam").Value);
                                            }
                                            string getwhsqty = " SELECT T0.ItemCode, T0.Quantity, T0.WhsCode,T1.U_PartyRef as 'NumAtCard',T1.U_BP_Name, (SELECT CardName from OCRD WHERE CardCode = T1.U_BP_Name) as 'CardName' from IGE1 AS T0 LEFT JOIN OIGE AS T1 ON T0.DocEntry = T1.DocEntry where T0.DocEntry = '" + oForm.Items.Item("tDocEnt").Specific.value + "'";
                                            SAPbobsCOM.Recordset rec5 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                            rec5.DoQuery(getwhsqty);
                                            if (rec5.RecordCount > 0)
                                            {
                                                oForm.Items.Item("tQty").Specific.value = Convert.ToDouble(rec5.Fields.Item("Quantity").Value);
                                                oForm.Items.Item("tWhs").Specific.value = Convert.ToString(rec5.Fields.Item("WhsCode").Value);
                                                oForm.Items.Item("tCardName").Specific.value = Convert.ToString(rec5.Fields.Item("CardName").Value);
                                                oForm.Items.Item("tRefNo").Specific.value = Convert.ToString(rec5.Fields.Item("NumAtCard").Value);
                                                oForm.Items.Item("tCardCode").Specific.value =  rec5.Fields.Item("U_BP_Name").Value.ToString();
                                            }
                                        }
                                        if (DocType == "GR" || DocType == "RFP")
                                        {
                                            string getBatch = "Select T0.BatchNum From IBT1 T0 Inner Join OIGN T1 On T0.BaseType=T1.ObjType ";
                                            getBatch = getBatch + " and T0.BaseEntry = T1.DocEntry Where T0.BaseType = 59 and T0.ItemCode = '" + oDataTable.GetValue("ItemCode", 0).ToString() + "' AND T1.DocEntry = '" + oForm.Items.Item("tDocEnt").Specific.value + "'";
                                            SAPbobsCOM.Recordset rec3 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                            rec3.DoQuery(getBatch);
                                            if (rec3.RecordCount > 0)
                                            {
                                                oForm.Items.Item("tBatchNo").Specific.value = Convert.ToString(rec3.Fields.Item("BatchNum").Value);
                                            }
                                            string getItemGrp = "select T1.ItmsGrpNam from OITM as T0 LEFT JOIN OITB AS T1 ON T1.ItmsGrpCod = T0.ItmsGrpCod where T0.ItemCode = '" + oDataTable.GetValue("ItemCode", 0).ToString() + "'";
                                            SAPbobsCOM.Recordset rec4 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                            rec4.DoQuery(getItemGrp);
                                            if (rec4.RecordCount > 0)
                                            {
                                                oForm.Items.Item("tItemGrp").Specific.value = Convert.ToString(rec4.Fields.Item("ItmsGrpNam").Value);
                                            }
                                            string getwhsqty = " SELECT T0.ItemCode, T0.Quantity, T0.WhsCode, T1.U_BP_Name, (SELECT CardName from OCRD WHERE CardCode = T1.U_BP_Name) as 'CardName' from IGN1 AS T0 LEFT JOIN OIGN AS T1 ON T0.DocEntry = T1.DocEntry where T0.DocEntry = '" + oForm.Items.Item("tDocEnt").Specific.value + "'";
                                           // string getwhsqty = "  SELECT Quantity,WhsCode from IGN1 where DocEntry = '" + oForm.Items.Item("tDocEnt").Specific.value + "'";
                                            SAPbobsCOM.Recordset rec5 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                            rec5.DoQuery(getwhsqty);
                                            if (rec5.RecordCount > 0)
                                            {
                                                oForm.Items.Item("tQty").Specific.value = Convert.ToDouble(rec5.Fields.Item("Quantity").Value);
                                                oForm.Items.Item("tWhs").Specific.value = Convert.ToString(rec5.Fields.Item("WhsCode").Value);
                                                oForm.Items.Item("tCardName").Specific.value = Convert.ToString(rec5.Fields.Item("CardName").Value); 
                                                oForm.Items.Item("tCardCode").Specific.value = rec5.Fields.Item("U_BP_Name").Value.ToString();
                                            }
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {

                                }
                            }
                        }
                        break;

                    case BoEventTypes.et_FORM_CLOSE:
                        if (pVal.BeforeAction == false)
                        {
                            oForm = null;
                            GC.Collect();
                            GC.WaitForPendingFinalizers();
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
                    case BoEventTypes.et_ITEM_PRESSED:
                        oForm = SBOMain.SBO_Application.Forms.ActiveForm;

                        if (pVal.BeforeAction == true){

                        } 
                        if (pVal.BeforeAction == false)
                        {
                            if (pVal.ItemUID == "1" && oForm.Mode == BoFormMode.fm_ADD_MODE)
                            {
                                Form_Load_Components(oForm, "ADD");
                            }
                            else if(pVal.ItemUID == "1" && oForm.Mode != BoFormMode.fm_ADD_MODE)
                            {
                                Form_Load_Components(oForm, "OK");
                            }
                        }
                        break; 
                     //default:
                }
            }
            catch (Exception ex)
            {
                // SBOMain.SBO_Application.MessageBox(ex.Message.ToString(), 1, "ok", "", "");

            }
            finally
            {
                 
            }  
            return BubbleEvent;
        }

        private void CFLConditionOITMGRPO(string CFLID, string ItemUID)
        {
            oCFL = oForm.ChooseFromLists.Item(CFLID);
            oConds = SBOMain.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions);

            oCond = oConds.Add();
            oCond.Alias = "validFor";
            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
            oCond.CondVal = "Y";
            oCFL.SetConditions(oConds);
            oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

            int abc = 0;
            string docentry = oForm.Items.Item("tDocEnt").Specific.value;
            string Query = "SELECT T1.ItemCode FROM OPDN AS T0  LEFT JOIN PDN1 AS T1 ON T0.DocEntry = T1.DocEntry WHERE T0.DocEntry = '" + docentry + "' ";
                   Query = Query + " and T1.ItemCode NOT In (SELECT U_ItemCode from dbo.[@QCSR] WHERE U_DocType = 'GRPO' AND U_BaseDocEnt = '" + docentry + "')";
            
            SAPbobsCOM.Recordset rec = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            rec.DoQuery(Query);
            if (rec.RecordCount > 0)
            {
                while (!rec.EoF)
                { 
                    if (abc != 0) { 
                        oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR;
                    } 
                    oCond = oConds.Add();
                    oCond.Alias = "ItemCode";
                    oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    oCond.CondVal =  rec.Fields.Item("ItemCode").Value.ToString(); 
                    oCFL.SetConditions(oConds);
                    abc++; 
                    rec.MoveNext();
                }
            }  

            oCFL.SetConditions(oConds);
            oCFL = null;
            oCond = null;
            oConds = null;
        }
        private void CFLConditionOITMGI(string CFLID, string ItemUID)
        {
            oCFL = oForm.ChooseFromLists.Item(CFLID);
            oConds = SBOMain.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions);

            oCond = oConds.Add();
            oCond.Alias = "validFor";
            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
            oCond.CondVal = "Y";
            oCFL.SetConditions(oConds);
            oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

            int abc = 0;

            string docentry = oForm.Items.Item("tDocEnt").Specific.value;
            string Query = "SELECT T1.ItemCode FROM OIGE AS T0 LEFT JOIN IGE1 AS T1 ON T0.DocEntry = T1.DocEntry WHERE T0.DocEntry = '" + docentry + "'";
                   Query = Query + " AND T1.ItemCode NOT In (SELECT U_ItemCode from dbo.[@QCSR] WHERE U_DocType = 'GI' AND U_BaseDocEnt =  '" + docentry + "')";

            SAPbobsCOM.Recordset rec = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            rec.DoQuery(Query);
            if (rec.RecordCount > 0)
            {
                while (!rec.EoF)
                {
                    if (abc != 0)
                    {
                        oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR;
                    } 
                    oCond = oConds.Add();
                    oCond.Alias = "ItemCode";
                    oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    oCond.CondVal = rec.Fields.Item("ItemCode").Value.ToString();
                    oCFL.SetConditions(oConds);
                    abc++;
                    rec.MoveNext();
                }
            } 
            oCFL.SetConditions(oConds);
            oCFL = null;
            oCond = null;
            oConds = null;
        }

        private void CFLConditionOITMGR(string CFLID, string ItemUID)
        {
            oCFL = oForm.ChooseFromLists.Item(CFLID);
            oConds = SBOMain.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions);

            oCond = oConds.Add();
            oCond.Alias = "validFor";
            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
            oCond.CondVal = "Y";
            oCFL.SetConditions(oConds);
            oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

            int abc = 0;

            string docentry = oForm.Items.Item("tDocEnt").Specific.value;

            string Que = "SELECT T1.ItemCode FROM OIGN AS T0 LEFT JOIN IGN1 AS T1 ON T0.DocEntry = T1.DocEntry WHERE T0.DocEntry = '" + docentry + "'";
                   Que = Que + " AND T1.ItemCode NOT In (SELECT U_ItemCode from dbo.[@QCSR] WHERE U_DocType = 'GR' AND U_BaseDocEnt = '" + docentry + "')";
            
            SAPbobsCOM.Recordset rec01 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            rec01.DoQuery(Que);
            if (rec01.RecordCount > 0)
            {
                while (!rec01.EoF)
                {
                    if (abc != 0)
                    {
                        oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR;
                    }
                    oCond = oConds.Add();
                    oCond.Alias = "ItemCode";
                    oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    oCond.CondVal = rec01.Fields.Item("ItemCode").Value.ToString();
                    oCFL.SetConditions(oConds);
                    abc++;
                    rec01.MoveNext();
                }
            } 
            oCFL.SetConditions(oConds);
            oCFL = null;
            oCond = null;
            oConds = null;
        }

        private void CFLConditionOITMRFP(string CFLID, string ItemUID)
        {
            oCFL = oForm.ChooseFromLists.Item(CFLID);
            oConds = SBOMain.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions);

            oCond = oConds.Add();
            oCond.Alias = "validFor";
            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
            oCond.CondVal = "Y";
            oCFL.SetConditions(oConds);

            string docentry = oForm.Items.Item("tDocEnt").Specific.value;

            string Que = "SELECT T1.ItemCode FROM OIGN AS T0 LEFT JOIN IGN1 AS T1 ON T0.DocEntry = T1.DocEntry WHERE T0.DocEntry = '" + docentry + "'";
            string Que1 = "SELECT U_ItemCode from dbo.[@QCSR] WHERE U_DocType = 'RFP' AND U_BaseDocEnt =  '" + docentry + "'";

            SAPbobsCOM.Recordset rec01 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            rec01.DoQuery(Que);
            if (rec01.RecordCount > 0)
            {
                while (!rec01.EoF)
                {
                    oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;
                    oCond = oConds.Add();
                    oCond.Alias = "ItemCode";
                    oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    oCond.CondVal = rec01.Fields.Item("ItemCode").Value.ToString();
                    oCFL.SetConditions(oConds);
                    rec01.MoveNext();
                }
            }
            oCFL.SetConditions(oConds);
            oCFL = null;
            oCond = null;
            oConds = null;
        }
  

        private void CFLConditionBP(string CFLID, string ItemUID)
        {
            oCFL = oForm.ChooseFromLists.Item(CFLID);
            oConds = SBOMain.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions);

            oCond = oConds.Add();
            oCond.Alias = "validFor";
            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
            oCond.CondVal = "Y";
            oCFL.SetConditions(oConds); 
            oCFL = null;
            oCond = null;
            oConds = null;
        }

        private void CFLConditionGRPO(string CFLID, string ItemUID)
        {
            oCFL = oForm.ChooseFromLists.Item(CFLID);
            oConds = SBOMain.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions);
              
            oCond = oConds.Add();
            oCond.Alias = "DocDate";
            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_GRATER_EQUAL;
            oCond.CondVal = "20220401";
            oCFL.SetConditions(oConds);

            if (!string.IsNullOrEmpty(oForm.Items.Item("tDocEnt").Specific.value))
            {
                oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;
                oCond = oConds.Add();
                oCond.Alias = "DocEntry";
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCond.CondVal = oForm.Items.Item("tDocEnt").Specific.value; 
            }
            else { 
                SAPbobsCOM.Recordset rec = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                string Query1 = "SELECT T0.DocEntry, T0.DocNum, (Select count(*) from PDN1 as T2 where T2.DocEntry = T0.DocEntry) As Total, ";
                Query1 = Query1 + " (Select count(*) from dbo.[@QCSR] as T3 where T3.U_DocType = 'GRPO' AND T3.U_BaseDocEnt = T0.DocEntry) As Total2 ";
                Query1 = Query1 + " FROM OPDN AS T0  where T0.DocDate > '20220401'";

                rec.DoQuery(Query1);
                if (rec.RecordCount > 0)
                {
                    while (!rec.EoF)
                    {
                        if(rec.Fields.Item("Total").Value == rec.Fields.Item("Total2").Value) { 
                            oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;
                            oCond = oConds.Add();
                            oCond.Alias = "DocEntry";
                            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_EQUAL;
                            oCond.CondVal = rec.Fields.Item("DocEntry").Value.ToString();
                        }
                        rec.MoveNext();
                    }
                }
            }
            oCFL.SetConditions(oConds);
            oCFL = null;
            oCond = null;
            oConds = null;
        }
        private void CFLConditionGI(string CFLID, string ItemUID)
        {
            oCFL = oForm.ChooseFromLists.Item(CFLID);
            oConds = SBOMain.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions);

            oCond = oConds.Add();
            oCond.Alias = "DocDate";
            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_GRATER_EQUAL;
            oCond.CondVal = "20220401";
            oCFL.SetConditions(oConds);

            oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;
            oCond = oConds.Add();
            oCond.Alias = "JrnlMemo";
            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
            oCond.CondVal = "Goods Issue";
            oCFL.SetConditions(oConds);

            if (!string.IsNullOrEmpty(oForm.Items.Item("tDocEnt").Specific.value))
            {
                oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;
                oCond = oConds.Add();
                oCond.Alias = "DocEntry";
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCond.CondVal = oForm.Items.Item("tDocEnt").Specific.value; 
            }
            else
            {
                SAPbobsCOM.Recordset rec = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                string Query1 = "SELECT T0.DocEntry, T0.DocNum, (Select count(*) from IGE1 as T2 where T2.DocEntry = T0.DocEntry) As Total, ";
                Query1 = Query1 + " (Select count(*) from dbo.[@QCSR] as T3 where T3.U_DocType = 'GI' AND T3.U_BaseDocEnt = T0.DocEntry) As Total2 ";
                Query1 = Query1 + " FROM OIGE AS T0  where T0.DocDate > '20220401'";

                rec.DoQuery(Query1);
                if (rec.RecordCount > 0)
                {
                    while (!rec.EoF)
                    {
                        if (rec.Fields.Item("Total").Value == rec.Fields.Item("Total2").Value)
                        {
                            oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;
                            oCond = oConds.Add();
                            oCond.Alias = "DocEntry";
                            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_EQUAL;
                            oCond.CondVal = rec.Fields.Item("DocEntry").Value.ToString();
                        }
                        rec.MoveNext();
                    }
                }
            }
            oCFL.SetConditions(oConds);
            oCFL = null;
            oCond = null;
            oConds = null;
        } 
        private void CFLConditionGR(string CFLID, string ItemUID)
        {
            oCFL = oForm.ChooseFromLists.Item(CFLID);
            oConds = SBOMain.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions);

            oCond = oConds.Add();
            oCond.Alias = "DocDate";
            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_GRATER_EQUAL;
            oCond.CondVal = "20220401";
            oCFL.SetConditions(oConds);

            oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;
            oCond = oConds.Add();
            oCond.Alias = "JrnlMemo";
            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
            oCond.CondVal = "Goods Receipt";
            oCFL.SetConditions(oConds);

            if (!string.IsNullOrEmpty(oForm.Items.Item("tDocEnt").Specific.value))
            {
                oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;
                oCond = oConds.Add();
                oCond.Alias = "DocEntry";
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCond.CondVal = oForm.Items.Item("tDocEnt").Specific.value;
                oCFL.SetConditions(oConds);
            }
            else
            {
                SAPbobsCOM.Recordset rec = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                string Query1 = "SELECT T0.DocEntry, T0.DocNum, (Select count(*) from IGN1 as T2 where T2.DocEntry = T0.DocEntry) As Total, ";
                Query1 = Query1 + " (Select count(*) from dbo.[@QCSR] as T3 where T3.U_DocType = 'GR' AND T3.U_BaseDocEnt = T0.DocEntry) As Total2 ";
                Query1 = Query1 + " FROM OIGN AS T0  where T0.DocDate > '20220401'";

                rec.DoQuery(Query1);
                if (rec.RecordCount > 0)
                {
                    while (!rec.EoF)
                    {
                        if (rec.Fields.Item("Total").Value == rec.Fields.Item("Total2").Value)
                        {
                            oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;
                            oCond = oConds.Add();
                            oCond.Alias = "DocEntry";
                            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_EQUAL;
                            oCond.CondVal = rec.Fields.Item("DocEntry").Value.ToString();
                            oCFL.SetConditions(oConds);
                        }
                        rec.MoveNext();
                    }
                }
            }

            oCFL.SetConditions(oConds);
            oCFL = null;
            oCond = null;
            oConds = null;
        }
        private void CFLConditionRFP(string CFLID, string ItemUID)
        {
            oCFL = oForm.ChooseFromLists.Item(CFLID);
            oConds = SBOMain.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions);

            oCond = oConds.Add();
            oCond.Alias = "DocDate";
            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_GRATER_EQUAL;
            oCond.CondVal = "20220401";
            oCFL.SetConditions(oConds); 

            oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;
            oCond = oConds.Add();
            oCond.Alias = "JrnlMemo";
            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
            oCond.CondVal = "Receipt from Production";
            oCFL.SetConditions(oConds);

            if (!string.IsNullOrEmpty(oForm.Items.Item("tDocEnt").Specific.value))
            {
                oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;
                oCond = oConds.Add();
                oCond.Alias = "DocEntry";
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCond.CondVal = oForm.Items.Item("tDocEnt").Specific.value;
                oCFL.SetConditions(oConds);
            }
            else
            {
                SAPbobsCOM.Recordset rec = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                string Query1 = "SELECT T0.DocEntry, T0.DocNum, (Select count(*) from IGN1 as T2 where T2.DocEntry = T0.DocEntry) As Total, ";
                Query1 = Query1 + " (Select count(*) from dbo.[@QCSR] as T3 where T3.U_DocType = 'RFP' AND T3.U_BaseDocEnt = T0.DocEntry) As Total2 ";
                Query1 = Query1 + " FROM OIGN AS T0  where T0.DocDate > '20220401'";

                rec.DoQuery(Query1);
                if (rec.RecordCount > 0)
                {
                    while (!rec.EoF)
                    {
                        if (rec.Fields.Item("Total").Value == rec.Fields.Item("Total2").Value)
                        {
                            oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;
                            oCond = oConds.Add();
                            oCond.Alias = "DocEntry";
                            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_EQUAL;
                            oCond.CondVal = rec.Fields.Item("DocEntry").Value.ToString();
                            oCFL.SetConditions(oConds);
                        }
                        rec.MoveNext();
                    }
                }
            } 
            oCFL.SetConditions(oConds);
            oCFL = null;
            oCond = null;
            oConds = null; 
        }
        

        private void AddMatrixRow(SAPbouiCOM.Matrix matrix, string ColUID)
        {
            if (matrix.RowCount == 0)
                matrix.AddRow();
            else
            {
                if (!string.IsNullOrEmpty(matrix.Columns.Item(ColUID).Cells.Item(matrix.RowCount).Specific.value))
                {
                    matrix.AddRow();
                    matrix.ClearRowData(matrix.RowCount);
                }

            }
            ArrengeMatrixLineNum(matrix);
        }
        private void ArrengeMatrixLineNum(SAPbouiCOM.Matrix matrix)
        {
            try
            {
                for (int i = 1; i <= matrix.RowCount; i++)
                {
                    matrix.Columns.Item("#").Cells.Item(i).Specific.value = i;
                }
            }
            catch (Exception ex)
            {

            }
        }
        public void Form_Load_Components(SAPbouiCOM.Form oForm, string mode)
        {
            oForm.Freeze(true);
            oForm.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized;
            SAPbouiCOM.EditText oEdit;
            string Table = "@QCSR";
            DateTime now = DateTime.Now;

            if (mode != "OK")
            {
                oEdit = oForm.Items.Item("tCode").Specific;
                objCU.GetNextDocNum(ref oEdit, ref Table);
                oForm.Items.Item("tDocNum").Specific.value = oForm.BusinessObject.GetNextSerialNumber("DocEntry", "QCSR");
                Events.Series.SeriesCombo("QCSR", "cSer");
                oForm.Items.Item("cSer").DisplayDesc = true;

                oForm.Items.Item("tDocDate").Specific.value = DateTime.Now.ToString("yyyyMMdd");

                cb1 = (SAPbouiCOM.ComboBox)oForm.Items.Item("tDocType").Specific;
                cb1.ExpandType = BoExpandType.et_DescriptionOnly;
                cb1.Select("GRPO");
                  
                setChooseFromListField(oForm, CFL_GRPO, "CFL_GRPO", tDocNo, "tDocNo", SAPbouiCOM.BoLinkedObject.lf_GoodsReceiptPO);

                /* SAPbouiCOM.ButtonCombo BC1 = (SAPbouiCOM.ButtonCombo)oForm.Items.Item("cmbCPYF").Specific;
                 BC1.ValidValues.Add("GRPO", "Receipt from PO");
                 BC1.ValidValues.Add("GI", "Goods Issue");
                 BC1.ValidValues.Add("GR", "Goods Receipt");
                 BC1.ValidValues.Add("RFP", "Receipt from Production");
                 BC1.ExpandType = BoExpandType.et_DescriptionOnly;*/

        SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("matContent").Specific;
                AddMatrixRow(oMatrix, "tParamCode");
            }
            oForm.Freeze(false);
            //throw new NotImplementedException();

        }
    }
}
