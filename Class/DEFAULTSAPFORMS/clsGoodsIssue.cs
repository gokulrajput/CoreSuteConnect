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
using System.Drawing.Drawing2D;
using CoreSuteConnect.Class.PRICELIST;

namespace CoreSuteConnect.Class.DEFAULTSAPFORMS
{
    class clsGoodsIssue
    {
        #region VariableDeclaration
        private SAPbouiCOM.Form oForm;
        public string cFormID = string.Empty;
        public string DocKey = null;
        public static string getSalesForm = string.Empty;

        CommonUtility objCU = new CommonUtility();

        #endregion VariableDeclaration

        public clsGoodsIssue(OutwardToGI outClass)
        {
            try
            {
                if (outClass != null)
                {
                    oForm = SBOMain.SBO_Application.Forms.ActiveForm;
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    string Q1 = @"select T1.CardCode,T1.Docnum, T1.DocDate  from OIGE where DocEntry = '" + outClass.DocEntry + "'";
                    SAPbobsCOM.Recordset r1;
                    r1 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    r1.DoQuery(Q1);
                    if (r1.RecordCount > 0)
                    {
                        oForm.Items.Item("7").Specific.value = r1.Fields.Item("DocNum").Value; 
                        DateTime tDocDate = Convert.ToDateTime(r1.Fields.Item("DocDate").Value);
                        oForm.Items.Item("9").Specific.value = tDocDate.ToString("yyyyMMdd");
                        oForm.Items.Item("1").Click();
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

        // FOR JOBWORK ADDON
        public void SBO_Application_FormDataEvent(ref BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (BusinessObjectInfo.ActionSuccess == true)
                {
                    // if (!string.IsNullOrEmpty(SBOMain.TFromUID) && SBOMain.FromCnt != null)
                    // {
                    oForm = SBOMain.SBO_Application.Forms.ActiveForm;
                    string DocEntry = "";
                    SBOMain.GetDocEntryFromXml(BusinessObjectInfo.ObjectKey, ref DocEntry);
                    DocKey = DocEntry;

                    string DocNum = oForm.Items.Item("7").Specific.value;
                    string DocDate = oForm.Items.Item("9").Specific.value;

                    SAPbouiCOM.Form oUDFForm = SBOMain.SBO_Application.Forms.Item(oForm.UDFFormUID);
                    string JOWid = oUDFForm.Items.Item("U_JWODe").Specific.value;

                    string act = null;
                    string glact = objCU.GetJobWorkOutAccount();
                    string Inglact = objCU.GetJobWorkInAccount();

                    SAPbouiCOM.Matrix matCMTR = (SAPbouiCOM.Matrix)oForm.Items.Item("13").Specific;
                    for (int i = 1; i <= matCMTR.RowCount; i++)
                    {
                        act = ((SAPbouiCOM.EditText)matCMTR.Columns.Item("59").Cells.Item(1).Specific).Value;
                    } 
                     
                    // For JOBWORK In Case
                    if (act == Inglact)
                    {
                        string codeno = Convert.ToString(objCU.getTableRecordCount("JOREL") + 1);

                        // Inserting data OTR3 table for linked Documents
                        string q1 = "Select * From [dbo].[@ITR3] where DocEntry = '" + JOWid + "' and U_TransType =  'Goods Issue' AND U_BaseDocEnt = '" + DocKey + "' ";
                        SAPbobsCOM.Recordset r1 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                        r1.DoQuery(q1);
                        if (r1.RecordCount == 0)
                        {
                            string q3 = "select LineId,VisOrder from [dbo].[@ITR3] where DocEntry = '" + JOWid + "' order BY LineId Desc";
                            SAPbobsCOM.Recordset r3 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                            r3.DoQuery(q3);

                            string lineId = Convert.ToString(r3.Fields.Item("LineId").Value + 1);
                            string Visorder = Convert.ToString(r3.Fields.Item("VisOrder").Value + 1);

                            string q2 = "INSERT INTO [dbo].[@ITR3] ([DocEntry],[LineId], [VisOrder],[Object],[U_TransType], [U_BaseDocEnt], [U_BaseDocNo], [U_DocDate]) ";
                            q2 = q2 + "VALUES ( '" + JOWid + "' , '" + lineId + "', '" + Visorder + "', 'JITR', 'Goods Issue' , '" + DocKey + "', '" + DocNum + "' , '" + DocDate + "')";
                            SAPbobsCOM.Recordset r2 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                            r2.DoQuery(q2);
                        }
                        // Inserting data OTR3 table for linked Documents

                        string qry1 = "Select * From [dbo].[@JOREL] where U_JWIId = '" + JOWid + "' and U_GI =  '" + DocKey + "'";
                        SAPbobsCOM.Recordset rec2 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                        rec2.DoQuery(qry1);
                        if (rec2.RecordCount == 0)
                        {
                            string qry4 = "Select Top 1 * From [dbo].[@JOREL] where U_JWIId = '" + JOWid + "' and U_GR IS NOT NULL Order BY U_GR ASC";
                            SAPbobsCOM.Recordset rec4 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                            rec4.DoQuery(qry4);
                            if (rec4.RecordCount > 0)
                            {
                                double issuedQty = 0;
                                string ItemCode = null;
                                string code = rec4.Fields.Item("Code").Value;

                                string updtQry = "UPDATE [dbo].[@JOREL] SET [U_GI] = '" + DocKey + "' WHERE  Code  = '" + code + "' ";
                                SAPbobsCOM.Recordset recGR = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                recGR.DoQuery(updtQry);

                                SAPbouiCOM.Matrix oMatrix = oForm.Items.Item("13").Specific;

                                for (int i = 1; i <= oMatrix.RowCount; i++)
                                {
                                    ItemCode = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("1").Cells.Item(i).Specific).Value;

                                    if (!string.IsNullOrEmpty(ItemCode))
                                    {
                                        issuedQty = Convert.ToDouble(((SAPbouiCOM.EditText)oMatrix.Columns.Item("9").Cells.Item(i).Specific).Value);

                                        double TQty = 0;
                                        double Iqty = 0;
                                        double FIqty = 0;
                                        double Balqty = 0;

                                        string qry3 = "select U_Quantity,U_RecQty from [dbo].[@ITR1] where DocEntry = '" + JOWid + "' and U_ItemCode = '" + ItemCode + "'";
                                        SAPbobsCOM.Recordset rec3 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                        rec3.DoQuery(qry3);
                                        if (rec3.RecordCount > 0)
                                        {
                                            TQty = Convert.ToDouble(rec3.Fields.Item("U_Quantity").Value);
                                            Iqty = Convert.ToDouble(rec3.Fields.Item("U_RecQty").Value);
                                            FIqty = Iqty + issuedQty;
                                            Balqty = TQty - FIqty;
                                            string UpdateQry = "UPDATE [dbo].[@ITR1] set U_RecQty = '" + FIqty + "', U_BalQty  = '" + Balqty + "' WHERE DocEntry = '" + JOWid + "' and U_ItemCode = '" + ItemCode + "' ";
                                            SAPbobsCOM.Recordset recUp = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                            recUp.DoQuery(UpdateQry);
                                        }
                                    }
                                }
                                rec4.MoveNext();
                            }
                            rec2.MoveNext();
                        }
                    }

                    // JOBWORK OUT Case

                    if (act == glact)
                    {
                        if (!string.IsNullOrEmpty(JOWid))
                        { 
                            // Inserting data OTR3 table for linked Documents
                            string q1 = "Select * From [dbo].[@OTR3] where DocEntry = '" + JOWid + "' and U_TransType =  'Goods Issue' AND U_BaseDocEnt = '" + DocKey + "' ";
                            SAPbobsCOM.Recordset r1 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                            r1.DoQuery(q1);
                            if (r1.RecordCount == 0)    
                            {
                                string q3 = "select LineId,VisOrder from [dbo].[@OTR3] where DocEntry = '" + JOWid + "' order BY LineId Desc";
                                SAPbobsCOM.Recordset r3 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                r3.DoQuery(q3);

                                string lineId = Convert.ToString(r3.Fields.Item("LineId").Value + 1);
                                string Visorder = Convert.ToString(r3.Fields.Item("VisOrder").Value + 1);

                                string q2 = "INSERT INTO [dbo].[@OTR3] ([DocEntry],[LineId], [VisOrder],[Object],[U_TransType], [U_BaseDocEnt], [U_BaseDocNo],  [U_DocDate]) ";
                                q2 = q2 + "VALUES ( '" + JOWid + "' , '" + lineId + "', '" + Visorder + "', 'JOTR', 'Goods Issue' , '" + DocKey + "', '" + DocNum + "' , '" + DocDate + "')";
                                SAPbobsCOM.Recordset r2 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                r2.DoQuery(q2);
                            }
                            // Inserting data OTR3 table for linked Documents

                            string codeno = Convert.ToString(objCU.getTableRecordCount("JOREL") + 1);
 
                            string qry1 = "Select * From [dbo].[@JOREL] where U_JWOId = '" + JOWid + "' and U_GI =  '" + DocKey + "'";
                            SAPbobsCOM.Recordset rec2 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                            rec2.DoQuery(qry1);
                            if (rec2.RecordCount == 0)
                            {
                                double issuedQty = 0;
                                string ItemCode = null;

                                string insertQry = "INSERT INTO [dbo].[@JOREL] ([Code],[Name], [U_JWOId] ,[U_GI]) VALUES ( '" + codeno + "' , '" + codeno + "' ,'" + JOWid + "', '" + DocKey + "')";
                                SAPbobsCOM.Recordset recIn = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                recIn.DoQuery(insertQry);

                                SAPbouiCOM.Matrix oMatrix = oForm.Items.Item("13").Specific;

                                for (int i = 1; i <= oMatrix.RowCount; i++)
                                {
                                    ItemCode = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("1").Cells.Item(i).Specific).Value;

                                    if (!string.IsNullOrEmpty(ItemCode))
                                    {
                                        issuedQty = Convert.ToDouble(((SAPbouiCOM.EditText)oMatrix.Columns.Item("9").Cells.Item(i).Specific).Value);

                                        double TQty = 0;
                                        double Iqty = 0;
                                        double FIqty = 0;
                                        double Balqty = 0;

                                        string qry3 = "select U_Quantity,U_IsueQty from [dbo].[@OTR2] where DocEntry = '" + JOWid + "' and U_ItemCode = '" + ItemCode + "'";
                                        SAPbobsCOM.Recordset rec3 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                        rec3.DoQuery(qry3);
                                        if (rec3.RecordCount > 0)
                                        {
                                            TQty = Convert.ToDouble(rec3.Fields.Item("U_Quantity").Value);
                                            Iqty = Convert.ToDouble(rec3.Fields.Item("U_IsueQty").Value);
                                            FIqty = Iqty + issuedQty;
                                            Balqty = TQty - FIqty;
                                            string UpdateQry = "UPDATE [dbo].[@OTR2] set U_IsueQty = '" + FIqty + "', U_BalQty  = '" + Balqty + "' WHERE DocEntry = '" + JOWid + "' and U_ItemCode = '" + ItemCode + "' ";
                                            SAPbobsCOM.Recordset recUp = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                            recUp.DoQuery(UpdateQry);
                                        }
                                    }
                                }
                                rec2.MoveNext();
                            }  
                        }
                    }
                }
            }
            catch (Exception ex)
            {

            }
        }
        // FOR JOBWORK ADDON

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
                                SAPbouiCOM.EditText postingDate = (SAPbouiCOM.EditText)oForm.Items.Item("38").Specific;
                                DateTime d2 = Convert.ToDateTime(postingDate.String);
                                DateTime d1 = new DateTime(2022, 04, 01, 0, 0, 0);
                                int res = DateTime.Compare(d1, d2);
                                if (res == 1)
                                {
                                    SBOMain.SBO_Application.StatusBar.SetText("Select Goods Issue from current Finacial Year", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                                SAPbouiCOM.EditText NumAtCard = oForm.Items.Item("U_PartyRef").Specific;
                                SAPbouiCOM.ComboBox Series = oForm.Items.Item("30").Specific;
                                SAPbouiCOM.EditText InOutNo = oForm.Items.Item("U_InOutRef").Specific;

                                OutwardToQC.FormName = "GINew";
                                OutwardToQC.BPName = "";
                                OutwardToQC.BPCode = CardCode.Value;
                                OutwardToQC.DocNum = DocNum.Value;

                                OutwardToQC.ItemCode = itemCode;
                                OutwardToQC.ItemName = itemnm;
                                OutwardToQC.Qty = qty;
                                OutwardToQC.Whs = whscode;
                                OutwardToQC.ItemGroup = "";
                                OutwardToQC.Batchno = "";
                                OutwardToQC.NumAtCard = NumAtCard.Value;
                                OutwardToQC.InOutNo = InOutNo.Value;

                                string getDocEntry = "";
                                if (!string.IsNullOrEmpty(CardCode.Value))
                                {
                                    getDocEntry = "SELECT DocEntry FROM OIGE WHERE DocNum='" + DocNum.Value + "' AND Series='" + Series.Selected.Value + "' AND U_BP_NAME='" + CardCode.Value + "'";
                                }
                                else
                                {
                                    getDocEntry = "SELECT DocEntry FROM OIGE WHERE DocNum='" + DocNum.Value + "' AND Series='" + Series.Selected.Value + "' ";
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
                                    getQC = getQC + " U_DocType = 'GI' AND U_BaseDocEnt = '" + DocEntry + "' and U_BaseDocNo = '" + DocNum.Value + "' ";
                                    SAPbobsCOM.Recordset rec2 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                    rec2.DoQuery(getQC);
                                    if (rec2.RecordCount > 0)
                                    {
                                        OutwardToQC.QCDocEntry = Convert.ToString(rec2.Fields.Item("DocEntry").Value);
                                        OutwardToQC.FormName = "GIExist";
                                    }
                                    else
                                    {
                                        string getBatch = "Select T0.BatchNum From IBT1 T0 Inner Join OIGE T1 On T0.BaseType=T1.ObjType ";
                                        getBatch = getBatch + " and T0.BaseEntry = T1.DocEntry Where T0.BaseType = 60 and T0.ItemCode = '" + itemCode + "' AND T1.DocEntry = '" + DocEntry + "'";
                                        SAPbobsCOM.Recordset rec3 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                        rec3.DoQuery(getBatch);
                                        if (rec3.RecordCount > 0)
                                        {
                                            OutwardToQC.Batchno = Convert.ToString(rec3.Fields.Item("BatchNum").Value);
                                        }
                                        string getItemGrp = "select T1.ItmsGrpNam from OITM as T0 LEFT JOIN OITB AS T1 ON T1.ItmsGrpCod = T0.ItmsGrpCod where T0.ItemCode = '" + itemCode + "'";
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
                                    SBOMain.SBO_Application.StatusBar.SetText("Goods Issue is not saved.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                            if (pVal.ItemUID == "1")
                            {
                                SAPbouiCOM.Form oUDFForm = SBOMain.SBO_Application.Forms.Item(oForm.UDFFormUID);
                                string JOWid = oUDFForm.Items.Item("U_JWODe").Specific.value;
                                if (!string.IsNullOrEmpty(JOWid))
                                {
                                    string docentry = DocKey;

                                    string cardcode = oForm.Items.Item("3").Specific.Value;
                                    string getDefWhs = "Select top 1 DocEntry, DocNum,CardCode from OIGE where U_JWODe = '" + JOWid + "' ORDER BY DocEntry DESC";
                                    SAPbobsCOM.Recordset rec1 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                    rec1.DoQuery(getDefWhs);
                                    if (rec1.RecordCount > 0)
                                    {
                                        string insertQry = "INSERT INTO JOREL set U_JWOId = '" + JOWid + "',U_JWOId = '" + Convert.ToString(rec1.Fields.Item("DocEntry").Value) + "' ";
                                        SAPbobsCOM.Recordset recIn = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                        recIn.DoQuery(insertQry);
                                        // ((SAPbouiCOM.EditText)oMatrix.Columns.Item("tFrmWhs").Cells.Item(pVal.Row).Specific).Value = Convert.ToString(rec1.Fields.Item("DfltWH").Value);
                                        rec1.MoveNext();
                                    }
                                }
                            }
                        }

                        break;

                    case BoEventTypes.et_ITEM_PRESSED:
                        oForm = SBOMain.SBO_Application.Forms.ActiveForm;
                        if (pVal.BeforeAction == true)
                        {

                            if (pVal.ItemUID == "1" && (oForm.Mode == BoFormMode.fm_ADD_MODE))
                            {
                                string act = null;
                                string OutAct = objCU.GetJobWorkOutAccount();
                                decimal qty = 0;
                                string itm = null; 
                                
                                SAPbouiCOM.Form oUDFForm = SBOMain.SBO_Application.Forms.Item(oForm.UDFFormUID);
                                string JOID = oUDFForm.Items.Item("U_JWODe").Specific.value;
                               // oUDFForm.Items.Item("U_JWODe").Specific.value = oHeader.JWODE;

                                string Qry1 = null;

                                SAPbouiCOM.Matrix matGR = (SAPbouiCOM.Matrix)oForm.Items.Item("13").Specific;
                                for (int i = 1; i <= matGR.RowCount; i++)
                                {
                                    act = Convert.ToString(((SAPbouiCOM.EditText)matGR.Columns.Item("59").Cells.Item(i).Specific).Value);
                                 
                                    if (act == OutAct)
                                    {
                                        qty = Convert.ToDecimal(((SAPbouiCOM.EditText)matGR.Columns.Item("9").Cells.Item(i).Specific).Value);
                                        itm = Convert.ToString(((SAPbouiCOM.EditText)matGR.Columns.Item("1").Cells.Item(i).Specific).Value);

                                        Qry1 = "  Select U_BalQty from dbo.[@OTR2] Where DocEntry = '"+ JOID + "' and U_ItemCode = '"+ itm + "' ";
                                        SAPbobsCOM.Recordset rec4 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                        rec4.DoQuery(Qry1);
                                        if (rec4.RecordCount > 0)
                                        {
                                              if (qty > Convert.ToDecimal(rec4.Fields.Item("U_BalQty").Value))
                                                {
                                                BubbleEvent = false;
                                                SBOMain.SBO_Application.StatusBar.SetText("Inventory Transfer Qty should be less or equal to Balance Quantity for Item : '"+ itm + "'", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                             }
                                        }  
                                    }
                                }    
                            }
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

    public class OutwardToGI
    {
        public string DocEntry { get; set; }
    }
}

