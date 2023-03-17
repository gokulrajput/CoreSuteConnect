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
using System.Configuration;
using CoreSuteConnect.Class.DEFAULTSAPFORMS;
using CoreSuteConnect.Class.EXIM;
using System.Collections.Specialized;
using System.Drawing.Drawing2D;
using System.Collections;

namespace CoreSuteConnect.Class.AUTOEMAIL
{
    class clsEmailAutomation
    {
        #region VariableDeclaration

        public static SAPbouiCOM.Application SBO_Application;

        private SAPbouiCOM.Form oForm;
        private SAPbouiCOM.ComboBox oComboBoxDocType;

        public string cFormID = string.Empty;
        string EmailFR, EmailCC, EmailSUB, EmailBDY, EmailRM, DepWise, SeWise, OwWise, FrmBPM, AAcnt, Defcnt, cntTran;
        string PDDocType, PDEmailFR, PDEmailCC, PDEmailSUB, PDEmailBDY, PDEmailRM, PDDepWise, PDSeWise, PDOwWise, PDFrmBPM, PDAAcnt, PDDefcnt, PDcntTran;
        
        string qry, qry1;

        private SAPbouiCOM.CheckBox oChk1, oChk2, oChk3, oChk4, oChk5, oChk6 , oChk7;
        private SAPbouiCOM.CheckBox PDoChk1, PDoChk2, PDoChk3, PDoChk4, PDoChk5, PDoChk6, PDoChk7;

        #endregion VariableDeclaration

        public bool MenuEvent(ref SAPbouiCOM.MenuEvent pVal, string FormId, string Type = "")
        {
            bool BubbleEvent = true;
            try
            {
                cFormID = FormId;
                oForm = SBOMain.SBO_Application.Forms.Item(FormId);

                if (pVal.BeforeAction == true)
                {  

                }
                if (pVal.BeforeAction == false)
                {
                    //if ((oForm.Mode == BoFormMode.fm_ADD_MODE || Type == "ADDNEWFORM") && Type != "DEL_ROW")
                    //{
                        Form_Load_Components(oForm);
                        setData(oForm);
                     //}
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

                    
                    case BoEventTypes.et_KEY_DOWN:
                        if (pVal.BeforeAction == true)
                        {
                            if (pVal.CharPressed == 9)
                            {
                            }
                        }
                        if (pVal.BeforeAction == false)
                        {
                            if (pVal.CharPressed == 9)
                            {
                                
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
                    case BoEventTypes.et_COMBO_SELECT:
                        if (pVal.BeforeAction == true)
                        {
                        }
                        if (pVal.BeforeAction == false)
                        {
                            if (pVal.ItemUID == "PDDocType")
                            {
                                string selctedExt = oForm.Items.Item("PDDocType").Specific.Selected.Value;
                            }
                        }
                        break;
                    case BoEventTypes.et_CLICK:
                        oForm = SBOMain.SBO_Application.Forms.ActiveForm;

                        if (pVal.BeforeAction == true)
                        {
                            if (pVal.ItemUID == "1")
                            {
                                if((oForm.Mode == BoFormMode.fm_UPDATE_MODE))
                                {
                                    /* if (!String.IsNullOrEmpty(oForm.Items.Item("exdnde").Specific.Value.ToString()))
                                     {
                                         string statusval = oForm.Items.Item("exls").Specific.value.ToString();
                                         if (statusval == "")
                                         {
                                             SBOMain.SBO_Application.StatusBar.SetText("Please select type Licence/Scheme.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                             BubbleEvent = false;
                                         }
                                     }*/

                                    /*oCheckbox = oForm.Items.Item("cschTPLP").Specific;
                                    if (oCheckbox.Checked && (oForm.Mode == BoFormMode.fm_UPDATE_MODE)
                                              && !string.IsNullOrEmpty(oForm.Items.Item("tschapinde").Specific.Value.ToString()))
                                    {
                                        SBOMain.SBO_Application.StatusBar.SetText("A/P invoice already linked in third party license purchase.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                        BubbleEvent = false;
                                    }*/
                                }
                            }
                        }
                        if (pVal.BeforeAction == false)
                        {
                            if (pVal.ItemUID == "1")
                            {
                                if ((oForm.Mode == BoFormMode.fm_UPDATE_MODE))
                                {   
                                     EmailFR = oForm.Items.Item("EmailFR").Specific.value.ToString();
                                     EmailCC = oForm.Items.Item("EmailCC").Specific.value.ToString();
                                     EmailSUB = oForm.Items.Item("EmailSUB").Specific.value.ToString();
                                     EmailBDY = oForm.Items.Item("EmailBDY").Specific.value.ToString();
                                     EmailRM = oForm.Items.Item("EmailRM").Specific.value.ToString();
                                     oChk1 = oForm.Items.Item("DepWise").Specific;
                                     oChk2 = oForm.Items.Item("SeWise").Specific;
                                     oChk3 = oForm.Items.Item("OwWise").Specific;
                                     oChk4 = oForm.Items.Item("FrmBPM").Specific;
                                     oChk5 = oForm.Items.Item("AAcnt").Specific;
                                     oChk6 = oForm.Items.Item("Defcnt").Specific;
                                     oChk7 = oForm.Items.Item("cntTran").Specific;

                                     if (oChk1.Checked) { DepWise = "Y"; } else { DepWise = "N"; }
                                     if (oChk2.Checked) { SeWise = "Y"; } else { SeWise = "N"; }
                                     if (oChk3.Checked) { OwWise = "Y"; } else { OwWise = "N"; }
                                     if (oChk4.Checked) { FrmBPM = "Y"; } else { FrmBPM = "N"; }
                                     if (oChk5.Checked) { AAcnt = "Y"; } else { AAcnt = "N"; }
                                     if (oChk6.Checked) { Defcnt = "Y"; } else { Defcnt = "N"; }
                                     if (oChk7.Checked) { cntTran = "Y"; } else { cntTran = "N"; }
                                    try
                                    {
                                        qry = "select * from dbo.[@AUEM]";
                                        SAPbobsCOM.Recordset rec = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                        rec.DoQuery(qry);
                                        if (rec.RecordCount == 0)
                                        {
                                            qry1 = "INSERT INTO[dbo].[@AUEM] ([Code], [Name], [U_EmailFR], [U_EmailCC], [U_EmailSUB],  [U_EmailBDY], [U_EmailRM], [U_DepWise], ";
                                            qry1 = qry1 + " [U_SeWise], [U_OwWise], [U_FrmBPM] ,[U_AAcnt] ,[U_Defcnt] ,[U_cntTran])  VALUES ";
                                            qry1 = qry1 + "  ('1','',  '" + EmailFR + "', '"+ EmailCC + "',  '"+ EmailSUB + "', '" + EmailBDY + "', '" + EmailRM + "',  '" + DepWise + "',  ";
                                            qry1 = qry1 + "             '"+ SeWise + "',  '"+ OwWise + "',   '"+ FrmBPM + "',   '"+ AAcnt + "',     '"+ Defcnt + "', '"+ cntTran + "' ) ";
                                            SAPbobsCOM.Recordset rec1 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                            rec1.DoQuery(qry1);
                                        } 
                                        else if (rec.RecordCount > 0)
                                        {
                                            qry1 = "UPDATE [dbo].[@AUEM] SET [U_EmailFR] = '" + EmailFR + "' , [U_EmailCC] = '" + EmailCC + "' ,[U_EmailSUB] ='" + EmailSUB + "'  ,[U_EmailBDY] ='" + EmailBDY + "' ";
                                            qry1 = qry1 + " ,[U_EmailRM] = '" + EmailRM + "' , [U_DepWise] = '" + DepWise + "' , [U_SeWise] = '" + SeWise + "', [U_OwWise] = '" + OwWise + "' ";
                                            qry1 = qry1 + " ,[U_FrmBPM] = '"+ FrmBPM + "' ,[U_AAcnt] = '"+ AAcnt +"', [U_Defcnt] = '"+ Defcnt +"', [U_cntTran] = '"+ cntTran +"' WHERE [Code] = 1";

                                            SAPbobsCOM.Recordset rec1 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                            rec1.DoQuery(qry1);
                                        }
                                    }
                                    catch(Exception ex)
                                    {
                                        SBOMain.SBO_Application.StatusBar.SetText(ex.Message.ToString(), BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                        BubbleEvent = false;
                                    }

                                    SAPbouiCOM.ComboBox cb3 = (SAPbouiCOM.ComboBox)oForm.Items.Item("PDDocType").Specific;

                                    PDDocType = oForm.Items.Item("PDDocType").Specific.Selected.Value;
                                    //PDDocType = oForm.Items.Item("PDDocType").Specific.value.ToString();
                                    PDEmailFR = oForm.Items.Item("PDEmailFR").Specific.value.ToString();
                                    PDEmailCC = oForm.Items.Item("PDEmailCC").Specific.value.ToString();
                                    PDEmailSUB = oForm.Items.Item("PDEmailSUB").Specific.value.ToString();
                                    PDEmailBDY = oForm.Items.Item("PDEmailBDY").Specific.value.ToString();
                                    PDEmailRM = oForm.Items.Item("PDEmailRM").Specific.value.ToString();
                                    PDoChk1 = oForm.Items.Item("PDDepWise").Specific;
                                    PDoChk2 = oForm.Items.Item("PDSeWise").Specific;
                                    PDoChk3 = oForm.Items.Item("PDOwWise").Specific;
                                    PDoChk4 = oForm.Items.Item("PDFrmBPM").Specific;
                                    PDoChk5 = oForm.Items.Item("PDAAcnt").Specific;
                                    PDoChk6 = oForm.Items.Item("PDDefcnt").Specific;
                                    PDoChk7 = oForm.Items.Item("PDcntTran").Specific;

                                    if (PDoChk1.Checked) { PDDepWise = "Y"; } else { PDDepWise = "N"; }
                                    if (PDoChk2.Checked) { PDSeWise = "Y"; } else { PDSeWise = "N"; }
                                    if (PDoChk3.Checked) { PDOwWise = "Y"; } else { PDOwWise = "N"; }
                                    if (PDoChk4.Checked) { PDFrmBPM = "Y"; } else { PDFrmBPM = "N"; }
                                    if (PDoChk5.Checked) { PDAAcnt = "Y"; } else { PDAAcnt = "N"; }
                                    if (PDoChk6.Checked) { PDDefcnt = "Y"; } else { PDDefcnt = "N"; }
                                    if (PDoChk7.Checked) { PDcntTran = "Y"; } else { PDcntTran = "N"; }
                                    try
                                    {
                                        qry = "select * from dbo.[@EMDW] Where U_PDDocType = '"+ PDDocType + "'";
                                        SAPbobsCOM.Recordset rec = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                        rec.DoQuery(qry);
                                        if (rec.RecordCount == 0)
                                        {
                                            qry1 = "INSERT INTO[dbo].[@EMDW] ([Name], [U_DocType], [U_EmailFR], [U_EmailCC], [U_EmailSUB],  [U_EmailBDY], [U_EmailRM], [U_DepWise], ";
                                            qry1 = qry1 + " [U_SeWise], [U_OwWise], [U_FrmBPM] ,[U_AAcnt] ,[U_Defcnt] ,[U_cntTran])  VALUES ";
                                            qry1 = qry1 + "  ('','"+ PDDocType + "' , '" + PDEmailFR + "', '" + PDEmailCC + "',  '" + PDEmailSUB + "', '" + PDEmailBDY + "', '" + PDEmailRM + "',  '" + PDDepWise + "',  ";
                                            qry1 = qry1 + "             '" + PDSeWise + "',  '" + PDOwWise + "',   '" + PDFrmBPM + "',   '" + PDAAcnt + "',     '" + PDDefcnt + "', '" + PDcntTran + "' ) ";
                                            SAPbobsCOM.Recordset rec1 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                            rec1.DoQuery(qry1);
                                        }
                                        else if (rec.RecordCount > 0)
                                        {
                                            qry1 = "UPDATE [dbo].[@EMDW] SET [U_EmailFR] = '" + PDEmailFR + "' , [U_EmailCC] = '" + PDEmailCC + "' ,[U_EmailSUB] ='" + PDEmailSUB + "'  ,[U_EmailBDY] ='" + PDEmailBDY + "' ";
                                            qry1 = qry1 + " ,[U_EmailRM] = '" + PDEmailRM + "' , [U_DepWise] = '" + PDDepWise + "' , [U_SeWise] = '" + PDSeWise + "', [U_OwWise] = '" + PDOwWise + "'";
                                            qry1 = qry1 + " ,[U_FrmBPM] = '" + PDFrmBPM + "' ,[U_AAcnt] = '" + PDAAcnt + "', [U_Defcnt] = '" + PDDefcnt + "', [U_cntTran] = '" + PDcntTran + "' WHERE [U_PDDocType] = '"+ PDDocType + "'";

                                            SAPbobsCOM.Recordset rec1 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                            rec1.DoQuery(qry1);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        SBOMain.SBO_Application.StatusBar.SetText(ex.Message.ToString(), BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                        BubbleEvent = false;
                                    }
                                }
                            }
                        }
                        break;

                    case BoEventTypes.et_FORM_DATA_LOAD:
                        oForm = SBOMain.SBO_Application.Forms.ActiveForm;
                        if (pVal.BeforeAction == false)
                        {
                            
                        }
                        break;

                    case BoEventTypes.et_ITEM_PRESSED:
                        oForm = SBOMain.SBO_Application.Forms.ActiveForm;
                        if (pVal.BeforeAction == true)
                        {
                            if (pVal.ItemUID == "1" && (oForm.Mode == BoFormMode.fm_ADD_MODE || oForm.Mode == BoFormMode.fm_UPDATE_MODE))
                            {
                                if (string.IsNullOrEmpty(oForm.Items.Item("exbc").Specific.value))
                                {
                                    BubbleEvent = false;
                                    SBOMain.SBO_Application.StatusBar.SetText("Please select business partner code", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                    oForm.Items.Item("exbc").Click();
                                }
                                else if (string.IsNullOrEmpty(oForm.Items.Item("exbn").Specific.value))
                                {
                                    BubbleEvent = false;
                                    SBOMain.SBO_Application.StatusBar.SetText("Please select business partner name", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                    oForm.Items.Item("exbn").Click();
                                }
                                else if (string.IsNullOrEmpty(oForm.Items.Item("exson").Specific.value))
                                {
                                    BubbleEvent = false;
                                    SBOMain.SBO_Application.StatusBar.SetText("Please select order", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                    oForm.Items.Item("exson").Click();
                                }
                                else
                                {
                                    SAPbouiCOM.ComboBox cb4 = oForm.Items.Item("exls").Specific;
                                    if (cb4.Selected == null)
                                    {
                                        BubbleEvent = false;
                                        SBOMain.SBO_Application.StatusBar.SetText("Please select Licence / Scheme.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                        oForm.Items.Item("exls").Click();
                                    }
                                }
                            }
                        }
                        if (pVal.BeforeAction == false)
                        {
                            if (pVal.ItemUID == "1" && oForm.Mode == BoFormMode.fm_ADD_MODE)
                            {
                                Form_Load_Components(oForm);
                                setData(oForm);
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
                /*if (oForm != null)
                    oForm.Freeze(false);*/
            }
            return BubbleEvent;
        }
        public void setData(SAPbouiCOM.Form oForm)
        {    
            SetValueGenForm(oForm);
            SetValueDocWiseForm(oForm,"SQ");
        }
        public void Form_Load_Components(SAPbouiCOM.Form oForm)
        {
            oForm.Items.Item("tab1").Visible = true;
            oForm.Items.Item("tab1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
            oForm.PaneLevel = 1;
            oComboBoxDocType = oForm.Items.Item("PDDocType").Specific;
            oComboBoxDocType.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;
            oComboBoxDocType.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
            oForm.Items.Item("tstatus").DisplayDesc = true;
        }
        public void SetValueGenForm(SAPbouiCOM.Form oForm)
        {
            string Query = "SELECT * from dbo.[@AUEM]";
            SAPbobsCOM.Recordset rec = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            rec.DoQuery(Query);
            if (rec.RecordCount > 0)
            {
                while (!rec.EoF)
                {
                    oForm.Items.Item("EmailFR").Specific.value = rec.Fields.Item("U_EmailFR").Value;
                    oForm.Items.Item("EmailCC").Specific.value = rec.Fields.Item("U_EmailCC").Value;
                    oForm.Items.Item("EmailBDY").Specific.value = rec.Fields.Item("U_EmailBDY").Value;
                    oForm.Items.Item("EmailSUB").Specific.value = rec.Fields.Item("U_EmailSUB").Value;
                    oForm.Items.Item("EmailRM").Specific.value = rec.Fields.Item("U_EmailRM").Value;
                    oForm.Items.Item("DepWise").Specific.Checked = true;
                    oForm.Items.Item("SeWise").Specific.Checked = true;
                    oForm.Items.Item("OwWise").Specific.Checked = true;
                    oForm.Items.Item("FrmBPM").Specific.Checked = true;
                    oForm.Items.Item("AAcnt").Specific.Checked = true;
                    oForm.Items.Item("Defcnt").Specific.Checked = true;
                    oForm.Items.Item("cntTran").Specific.Checked = true;
                    rec.MoveNext();
                }
            }
        }
        public void SetValueDocWiseForm(SAPbouiCOM.Form oForm, string DocType)
        {
            string Query = "SELECT * from dbo.[@EMDW] Where U_PDDocType = 'SQ'";
            SAPbobsCOM.Recordset rec = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            rec.DoQuery(Query);
            if (rec.RecordCount > 0)
            {
                while (!rec.EoF)
                {
                    // PDDocType 
                    oForm.Items.Item("PDEmailFR").Specific.value = rec.Fields.Item("U_EmailFR").Value;
                    oForm.Items.Item("PDEmailFR").Specific.value = rec.Fields.Item("U_EmailFR").Value;
                    oForm.Items.Item("PDEmailCC").Specific.value = rec.Fields.Item("U_EmailCC").Value;
                    oForm.Items.Item("PDEmailBDY").Specific.value = rec.Fields.Item("U_EmailBDY").Value;
                    oForm.Items.Item("PDEmailSUB").Specific.value = rec.Fields.Item("U_EmailSUB").Value;
                    oForm.Items.Item("PDEmailRM").Specific.value = rec.Fields.Item("U_EmailRM").Value;
                    oForm.Items.Item("PDDepWise").Specific.Checked = true;
                    oForm.Items.Item("PDSeWise").Specific.Checked = true;
                    oForm.Items.Item("PDOwWise").Specific.Checked = true;
                    oForm.Items.Item("PDFrmBPM").Specific.Checked = true;
                    oForm.Items.Item("PDAAcnt").Specific.Checked = true;
                    oForm.Items.Item("PDDefcnt").Specific.Checked = true;
                    oForm.Items.Item("PDcntTran").Specific.Checked = true;
                    rec.MoveNext();
                    //oForm.Items.Item("DepWise").Specific.checked = true;
                }
            }
        }
    }
}
