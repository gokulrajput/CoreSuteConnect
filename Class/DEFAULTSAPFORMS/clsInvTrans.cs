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
using System.Data;
using System.Drawing.Drawing2D;

namespace CoreSuteConnect.Class.DEFAULTSAPFORMS
{
    class clsInvTrans
    {
        #region VariableDeclaration
        private SAPbouiCOM.Form oForm;
        public string cFormID = string.Empty;

        public static string getSalesForm = string.Empty;

        string DocKey = null;

        CommonUtility objCU = new CommonUtility();

        #endregion VariableDeclaration

        public clsInvTrans(OutwardToInvTrans outClass)
        {
            try
            {
                if (outClass != null)
                {
                    oForm = SBOMain.SBO_Application.Forms.ActiveForm;
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    string Q1 = @"select T1.CardCode,T1.Docnum,T1.DocDate from OWTR as T1 where DocEntry = '" + outClass.DocEntry+"'";
                    SAPbobsCOM.Recordset r1;
                    r1 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    r1.DoQuery(Q1);
                    if(r1.RecordCount > 0)
                    {
                        oForm.Items.Item("3").Specific.value = r1.Fields.Item("CardCode").Value;
                        oForm.Items.Item("11").Specific.value = r1.Fields.Item("DocNum").Value;
                      
                        DateTime tDocDate = Convert.ToDateTime(r1.Fields.Item("DocDate").Value);
                        oForm.Items.Item("14").Specific.value = tDocDate.ToString("yyyyMMdd");

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
                    oForm = SBOMain.SBO_Application.Forms.ActiveForm;

                    string formmode = oForm.Mode.ToString();

                    string DocEntry = "";
                    SBOMain.GetDocEntryFromXml(BusinessObjectInfo.ObjectKey, ref DocEntry);
                    
                    DocKey = DocEntry;
                    string DocNum = oForm.Items.Item("11").Specific.value;
                    string DocDate = oForm.Items.Item("14").Specific.value;
                     
                    SAPbouiCOM.Form oUDFForm = SBOMain.SBO_Application.Forms.Item(oForm.UDFFormUID);
                    string JOWid = oUDFForm.Items.Item("U_JWODe").Specific.value;
                     
                     
                    if (!string.IsNullOrEmpty(JOWid))
                    {
                        string q1 = "Select * From [dbo].[@OTR3] where DocEntry = '" + JOWid + "' and U_TransType =  'Inventory Transfer' AND U_BaseDocEnt = '"+DocKey+"' ";
                        SAPbobsCOM.Recordset r1 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                        r1.DoQuery(q1);
                        if (r1.RecordCount == 0)
                        {
                            string q2 = "INSERT INTO [dbo].[@OTR3] ([DocEntry],[LineId], [VisOrder],[Object],[U_TransType], [U_BaseDocEnt], [U_BaseDocNo],[U_DocDate]) ";
                            q2 = q2 + "VALUES ( '" + JOWid + "' , 1,0, 'JOTR', 'Inventory Transfer' , '" + DocKey + "', '" + DocNum + "', '" + DocDate + "')";
                            SAPbobsCOM.Recordset r2 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                            r2.DoQuery(q2);
                        }
                     } 
                 } 
            }
            catch (Exception ex)
            {

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
                            /*if (pVal.ItemUID == "13" && (oForm.Mode == BoFormMode.fm_OK_MODE))
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
                            }*/
                        }
                        if (pVal.BeforeAction == false)
                        {
                            oForm = SBOMain.SBO_Application.Forms.ActiveForm;
                             
                        }
                        break;

                    case BoEventTypes.et_CLICK:

                        oForm = SBOMain.SBO_Application.Forms.ActiveForm;
                        if (pVal.BeforeAction == true)
                        {
                            if (pVal.ItemUID == "1")
                            {
                               
                            }
                        }
                        if (pVal.BeforeAction == false)
                        {
                            if (pVal.ItemUID == "1")
                            {
                                SAPbouiCOM.Form oUDFForm = SBOMain.SBO_Application.Forms.Item(oForm.UDFFormUID);
                                string JOWid = oUDFForm.Items.Item("U_JWODe").Specific.value;
                                if(!string.IsNullOrEmpty(JOWid))
                                {
                                    string cardcode = oForm.Items.Item("3").Specific.Value;

                                    string getDefWhs = "Select top 1 DocEntry, DocNum,CardCode from OWTR where CardCode = '" + oForm.Items.Item("3").Specific.Value + "' ORDER BY DocEntry DESC";
                                    SAPbobsCOM.Recordset rec1 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                    rec1.DoQuery(getDefWhs);
                                    if (rec1.RecordCount > 0)
                                    {
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
                                decimal qty = 0;
                                decimal bqty = 0;
                                string itm = null; 

                                SAPbouiCOM.Form oUDFForm = SBOMain.SBO_Application.Forms.Item(oForm.UDFFormUID);
                                string JOID = oUDFForm.Items.Item("U_JWODe").Specific.value; 
                                string Qry1 = null;
                                 
                                if (!string.IsNullOrEmpty(JOID))
                                {
                                    // CHECK ALREADY INVENTORY TRANSFER IS DONE OR NOT

                                    Qry1 = "Select count(*) as total from dbo.[@OTR3] where U_TransType = 'Inventory Transfer' AND DocEntry = '" + JOID + "' ";
                                    SAPbobsCOM.Recordset rec5 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                    rec5.DoQuery(Qry1);
                                    if (rec5.RecordCount > 0)
                                    {
                                        if (Convert.ToInt16(rec5.Fields.Item("total").Value) == 0)
                                        {
                                            BubbleEvent = false;
                                            SBOMain.SBO_Application.StatusBar.SetText("Inventory Transfer Already Done for Jobwork Out Doc No. : '"+ JOID + "'.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                        }
                                        else
                                        {
                                            // IF INVENTORY TRANSFER NOT DONE THEN CHECK Qty should not exceed
                                            SAPbouiCOM.Matrix matGR = (SAPbouiCOM.Matrix)oForm.Items.Item("23").Specific;
                                            for (int i = 1; i <= matGR.RowCount; i++)
                                            {
                                                qty = Convert.ToDecimal(((SAPbouiCOM.EditText)matGR.Columns.Item("10").Cells.Item(i).Specific).Value);
                                                itm = Convert.ToString(((SAPbouiCOM.EditText)matGR.Columns.Item("1").Cells.Item(i).Specific).Value);

                                                Qry1 = "Select U_Quantity from dbo.[@OTR2] Where DocEntry = '" + JOID + "' and U_ItemCode = '" + itm + "' ";
                                                SAPbobsCOM.Recordset rec4 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                                rec4.DoQuery(Qry1);
                                                if (rec4.RecordCount > 0)
                                                {
                                                    bqty = Convert.ToDecimal(rec4.Fields.Item("U_Quantity").Value);

                                                    if (qty > bqty)
                                                    {
                                                        BubbleEvent = false;
                                                        SBOMain.SBO_Application.StatusBar.SetText("Inventory Transfer Qty should be less or equal to Quantity for Item : '" + itm + "'", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    // CHECK ALREADY INVENTORY TRANSFER IS DONE OR NOT 
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
    
    public class OutwardToInvTrans
    {
        public string DocEntry { get; set; }
    }
}

