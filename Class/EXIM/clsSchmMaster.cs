using System;
using System.Collections.Generic;
using System.Linq;
using System.Collections;
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
using System.Collections.Specialized;
using CoreSuteConnect.Events;
using CoreSuteConnect.Class.DEFAULTSAPFORMS;

namespace CoreSuteConnect.Class.EXIM
{
    class clsSchmMaster
    {
        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        #region VariableDeclaration
        private SAPbouiCOM.Form oForm;
        private SAPbouiCOM.CheckBox oCheckbox;
        private SAPbouiCOM.ComboBox oComboBox;
        private SAPbouiCOM.ComboBox oComboBoxStatus;
         
        private SAPbouiCOM.Matrix oMatrix;
        public string cFormID = string.Empty;

        private SAPbouiCOM.ChooseFromList oCFL = null;
        private SAPbouiCOM.Conditions oConds = null;
        private SAPbouiCOM.Condition oCond = null;

        double qty = 0;
        double fqty = 0;
        double amt1 = 0;
        double amt2 = 0;

        OpenFileDialog OpenFileDialog = new OpenFileDialog();
        //string BrowseFilePath = string.Empty;

        ArrayList BrowseFilePath = new ArrayList();
        ArrayList ReplaceFilePath = new ArrayList();

        CommonUtility objCU = new CommonUtility();


        #endregion VariableDeclaration
        public clsSchmMaster(OutwardToSchemeMaster inClass)
        {
            if (inClass != null)
            {
                oForm = SBOMain.SBO_Application.Forms.ActiveForm;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                oForm.Items.Item("tschno").Specific.value = inClass.schmeno;
                oForm.Items.Item("1").Click();
            }
        }

        public bool MenuEvent(ref SAPbouiCOM.MenuEvent pVal, string FormId, string Type = "")
        {
            bool BubbleEvent = true;
            try
            {
                cFormID = FormId;
                oForm = SBOMain.SBO_Application.Forms.Item(FormId);

                if (pVal.BeforeAction == true)
                {
                    if (Type == "REMOVE")
                    {
                        String Q1 = "SELECT count(*) as total FROM dbo.[@XET10]  where U_ex10sc = '" + oForm.Items.Item("tschno").Specific.Value + "'";
                        SAPbobsCOM.Recordset rec1 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                        rec1.DoQuery(Q1);
                        int cnt1 = Convert.ToInt32(rec1.Fields.Item("total").Value);

                        if (cnt1 > 0)
                        {
                            SBOMain.SBO_Application.StatusBar.SetText("Remove operation not allowed because this Scheme is assigned in Exim Tracking", BoMessageTime.bmt_Long, BoStatusBarMessageType.smt_Error);
                            BubbleEvent = false;
                        } 
                    }
                }

                if (pVal.BeforeAction == false)
                { 

                    if ((oForm.Mode == BoFormMode.fm_ADD_MODE || oForm.Mode == BoFormMode.fm_UPDATE_MODE || Type == "ADDNEWFORM") && (Type != "DEL_ROW") && (Type != "ADD_ROW"))
                    {

                        Form_Load_Components(oForm, "ADD");
                        SchemeRateAndPercKGFields(oForm, "cschtype");
                        //oForm.Items.Item("tstatus").Specific.value = "Open";
                        //oForm.Items.Item("tstatus").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, (int)BoFormMode.fm_ADD_MODE, SAPbouiCOM.BoModeVisualBehavior.mvb_False);
                        oComboBox = oForm.Items.Item("cext").Specific;
                        oComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;
                        oComboBox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                        oForm.Items.Item("cext").DisplayDesc = true; 
                    }
                    if (Type == "navigation")
                    {
                        SchemeRateAndPercKGFields(oForm, "cschtype");
                        doAutoSummatEXOB(oForm);
                        doAutoSummatEXOB(oForm);
                    }
                    if (Type == "FIND")
                    {
                        SchemeRateAndPercKGFields(oForm, "cschtype");
                    }
                    else if (Type == "DEL_ROW" || Type == "ADD_ROW")
                    {
                        SAPbouiCOM.Matrix matEXOB = (SAPbouiCOM.Matrix)oForm.Items.Item("matEXOB").Specific;
                        SAPbouiCOM.Matrix matEXFL = (SAPbouiCOM.Matrix)oForm.Items.Item("matEXFL").Specific;
                        SAPbouiCOM.Matrix matEXIR = (SAPbouiCOM.Matrix)oForm.Items.Item("matEXIR").Specific;
                        SAPbouiCOM.Matrix matEXIU = (SAPbouiCOM.Matrix)oForm.Items.Item("matEXIU").Specific;
                        SAPbouiCOM.Matrix matHSN = (SAPbouiCOM.Matrix)oForm.Items.Item("matHSN").Specific;

                        if (Type == "ADD_ROW")
                        {
                            if (SBOMain.RightClickItemID == "matEXOB")
                            {
                                ADDROWMain(matEXOB);
                            }
                            else if (SBOMain.RightClickItemID == "matEXFL")
                            {
                                ADDROWMain(matEXFL);
                            }
                            else if (SBOMain.RightClickItemID == "matEXIR")
                            {
                                ADDROWMain(matEXIR);
                            }
                            else if (SBOMain.RightClickItemID == "matEXIU")
                            {
                                ADDROWMain(matEXIU);
                            }
                            else if (SBOMain.RightClickItemID == "matHSN")
                            {
                                ADDROWMain(matHSN);
                            }
                        }

                        if (Type == "DEL_ROW")
                        {
                            if (SBOMain.RightClickItemID == "matEXOB")
                            {
                                DeleteMatrixBlankRow(matEXOB, "titemcode");
                                ArrengeMatrixLineNum(matEXOB);
                            }
                            else if (SBOMain.RightClickItemID == "matEXIR")
                            {
                                DeleteMatrixBlankRow(matEXIR, "iritemcd");
                                ArrengeMatrixLineNum(matEXIR);
                            } 
                            else if (SBOMain.RightClickItemID == "matEXFL")
                            {    
                                DeleteMatrixBlankRow(matEXFL, "expexmno");
                                ArrengeMatrixLineNum(matEXFL);
                            } 
                            else if (SBOMain.RightClickItemID == "matEXIU")
                            {    
                                DeleteMatrixBlankRow(matEXIU, "impinvno");
                                ArrengeMatrixLineNum(matEXIU);
                            }
                            else if (SBOMain.RightClickItemID == "matHSN")
                            {    
                                DeleteMatrixBlankRow(matHSN, "hsncode");
                                ArrengeMatrixLineNum(matHSN);
                            }
                        }
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

        public bool ItemEvent(ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent, string FormId)
        {
            BubbleEvent = true;
            oForm = SBOMain.SBO_Application.Forms.Item(FormId);
            try
            {
                switch (pVal.EventType)
                { // KEY Down event for multiplication

                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
                        if (pVal.BeforeAction == false)
                        { 
                          // SchemeRateAndPercKGFields(oForm, "cschtype");
                        }
                        break;

                    case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
                        if (pVal.BeforeAction == true)
                        {
                        }
                        if (pVal.BeforeAction == false)
                        { 
                            if (pVal.ItemUID == "cschtype")
                            {
                                SchemeRateAndPercKGFields(oForm,"cschtype");
                                if (string.IsNullOrEmpty(oForm.Items.Item("tschsd").Specific.value) == false)
                                {
                                    string statusval = oForm.Items.Item("cschtype").Specific.value.ToString();
                                    if (statusval != "DBK" && statusval != "RoDTEP")
                                    {
                                        string selctedExt = oForm.Items.Item("cext").Specific.Selected.Value;
                                        SAPbouiCOM.EditText oDocDate1 = oForm.Items.Item("tschsd").Specific;
                                        DateTime lcsd = Convert.ToDateTime(oDocDate1.String);
                                        DateTime lcsd2 = objCU.Add_Year(lcsd);   // lcsd.AddYears(1);

                                        if (selctedExt == "-") // No Extension
                                        {
                                            lcsd2 = objCU.Add_Month(lcsd2); //lcsd2.AddMonths(6);
                                        }
                                        if (selctedExt == "1") // Extension 1
                                        {
                                            lcsd2 = objCU.Add_Year(lcsd2); //lcsd2.AddYears(1);
                                        }
                                        if (selctedExt == "2") // Extension 2
                                        {
                                            lcsd2 = objCU.Add_Year(lcsd2); //lcsd2.AddYears(1);
                                            lcsd2 = objCU.Add_Month(lcsd2);//lcsd2.AddMonths(6);
                                        }
                                        oForm.Items.Item("tschLEDE").Specific.Value = lcsd2.ToString("yyyyMMdd");
                                    }
                                }
                            }
                            if (pVal.ItemUID == "cext")
                            {  
                                if (string.IsNullOrEmpty(oForm.Items.Item("tschsd").Specific.value) == false)
                                {
                                    string statusval = oForm.Items.Item("cschtype").Specific.value.ToString();
                                    if (statusval != "DBK" && statusval != "RoDTEP")
                                    {
                                        string selctedExt = oForm.Items.Item("cext").Specific.Selected.Value;
                                        SAPbouiCOM.EditText oDocDate1 = oForm.Items.Item("tschsd").Specific;
                                        DateTime lcsd = Convert.ToDateTime(oDocDate1.String);
                                        DateTime lcsd2 = objCU.Add_Year(lcsd);   // lcsd.AddYears(1);

                                        if (selctedExt == "-") // No Extension
                                        {
                                            lcsd2 = objCU.Add_Month(lcsd2); //lcsd2.AddMonths(6);
                                        }
                                        if (selctedExt == "1") // Extension 1
                                        {
                                            lcsd2 = objCU.Add_Year(lcsd2); //lcsd2.AddYears(1);
                                        }
                                        if (selctedExt == "2") // Extension 2
                                        {
                                            lcsd2 = objCU.Add_Year(lcsd2); //lcsd2.AddYears(1);
                                            lcsd2 = objCU.Add_Month(lcsd2);//lcsd2.AddMonths(6);
                                        }
                                        oForm.Items.Item("tschLEDE").Specific.Value = lcsd2.ToString("yyyyMMdd");
                                    }
                                }
                            }
                        }
                        break;

                    case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
                        if (pVal.BeforeAction == false)
                        {
                            if (pVal.CharPressed == 9)
                            {
                                SAPbouiCOM.Matrix matEXOB = (SAPbouiCOM.Matrix)oForm.Items.Item("matEXOB").Specific;
                                SAPbouiCOM.Matrix matEXIR = (SAPbouiCOM.Matrix)oForm.Items.Item("matEXIR").Specific;

                                if (pVal.ItemUID == "matEXOB")
                                {
                                    if (pVal.ColUID == "tqty")
                                    {
                                        qty = Convert.ToDouble(((SAPbouiCOM.EditText)matEXOB.Columns.Item("tqty").Cells.Item(pVal.Row).Specific).Value);
                                        fqty = Convert.ToDouble(((SAPbouiCOM.EditText)matEXOB.Columns.Item("tflqty").Cells.Item(pVal.Row).Specific).Value);
                                        ((SAPbouiCOM.EditText)matEXOB.Columns.Item("rmqty").Cells.Item(pVal.Row).Specific).Value = Convert.ToString(qty - fqty);
                                        matEXOB.Columns.Item("tamtLC").Cells.Item(pVal.Row).Click();/// = Convert.ToString(qty - fqty);
                                        
                                        doAutoColSum(matEXOB, "rmqty");
                                    }
                                    if (pVal.ColUID == "tamtLC")
                                    {
                                        amt1 = Convert.ToDouble(((SAPbouiCOM.EditText)matEXOB.Columns.Item("tamtLC").Cells.Item(pVal.Row).Specific).Value);
                                        amt2 = Convert.ToDouble(((SAPbouiCOM.EditText)matEXOB.Columns.Item("tfllC").Cells.Item(pVal.Row).Specific).Value);
                                        ((SAPbouiCOM.EditText)matEXOB.Columns.Item("rmLC").Cells.Item(pVal.Row).Specific).Value = Convert.ToString(amt1 - amt2);
                                        matEXOB.Columns.Item("tamtFC").Cells.Item(pVal.Row).Click();
                                        doAutoColSum(matEXOB, "rmLC");
                                    }
                                    if (pVal.ColUID == "tamtFC")
                                    {
                                        amt1 = Convert.ToDouble(((SAPbouiCOM.EditText)matEXOB.Columns.Item("tamtFC").Cells.Item(pVal.Row).Specific).Value);
                                        amt2 = Convert.ToDouble(((SAPbouiCOM.EditText)matEXOB.Columns.Item("tflfC").Cells.Item(pVal.Row).Specific).Value);
                                        ((SAPbouiCOM.EditText)matEXOB.Columns.Item("rmFC").Cells.Item(pVal.Row).Specific).Value = Convert.ToString(amt1 - amt2);
                                       //matEXOB.Columns.Item("tamtFC").Cells.Item(pVal.Row).Click();
                                        doAutoColSum(matEXOB, "rmFC");
                                    }

                                    if (pVal.ColUID == "tqty" || pVal.ColUID == "tamtLC" || pVal.ColUID == "tamtFC" ||
                                        pVal.ColUID == "tflqty" || pVal.ColUID == "tfllC" || pVal.ColUID == "tflfC" ||
                                       pVal.ColUID == "rmqty" || pVal.ColUID == "rmLC" || pVal.ColUID == "rmFC")
                                    {
                                        doAutoColSum(matEXOB, pVal.ColUID);
                                    }
                                    if (pVal.ColUID == "titemcode")
                                    {
                                        AddMatrixRow(matEXOB, "titemcode");
                                    }
                                }
                                else if (pVal.ItemUID == "matEXFL" && pVal.ColUID == "texpinvno")
                                {
                                    SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("matEXFL").Specific;
                                    //AddMatrixRow(oMatrix, "texpinvno");
                                }
                                else if (pVal.ItemUID == "matEXIR")
                                {
                                    if (pVal.ColUID == "irqty" || pVal.ColUID == "iramtLC" || pVal.ColUID == "iramtFC" ||
                                        pVal.ColUID == "irutqty" || pVal.ColUID == "irutlC" || pVal.ColUID == "irutfC" ||
                                       pVal.ColUID == "irrmqty" || pVal.ColUID == "irrmLC" || pVal.ColUID == "irrmFC")
                                    {
                                        doAutoColSum(matEXIR, pVal.ColUID);
                                    }

                                    if (pVal.ColUID == "iritemcd")
                                    {
                                        AddMatrixRow(matEXIR, "iritemcd");
                                    }

                                }
                                else if (pVal.ItemUID == "matEXIU" && pVal.ColUID == "impinvno")
                                {
                                    SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("matEXIU").Specific;
                                    AddMatrixRow(oMatrix, "impinvno");
                                }
                                else if (pVal.ItemUID == "matHSN")
                                {
                                    SAPbouiCOM.Matrix matHSN = (SAPbouiCOM.Matrix)oForm.Items.Item("matHSN").Specific;
                                    if (pVal.ColUID == "hsncode")
                                    {
                                        string chapterid = ((SAPbouiCOM.EditText)matHSN.Columns.Item("hsncode").Cells.Item(pVal.Row).Specific).Value;
                                        string Query = "Select Dscription from OCHP where ChapterID = '" + chapterid + "'";
                                        SAPbobsCOM.Recordset rec = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                        rec.DoQuery(Query);
                                        string desc = rec.Fields.Item("Dscription").Value;
                                        ((SAPbouiCOM.EditText)matHSN.Columns.Item("hsnname").Cells.Item(pVal.Row).Specific).Value = desc;
                                        AddMatrixRow(matHSN, "hsncode");
                                    } 
                                }
                                if (pVal.ItemUID == "tschsd")
                                {

                                    if (string.IsNullOrEmpty(oForm.Items.Item("tschsd").Specific.value) == false)
                                    {
                                        string selctedExt = oForm.Items.Item("cext").Specific.Selected.Value;

                                        SAPbouiCOM.EditText oDocDate1 = oForm.Items.Item("tschsd").Specific;
                                        DateTime lcsd = Convert.ToDateTime(oDocDate1.String);
                                        DateTime lcsd1 = lcsd.AddYears(1);
                                        oForm.Items.Item("tschLEDI").Specific.Value = lcsd1.ToString("yyyyMMdd");

                                        DateTime lcsd2 = lcsd.AddYears(1);

                                        if (selctedExt == "-") // No Extension
                                        {
                                            lcsd2 = lcsd2.AddMonths(6);
                                        }
                                        if (selctedExt == "1") // Extension 1
                                        {
                                            lcsd2 = lcsd2.AddYears(1);
                                        }
                                        if (selctedExt == "2") // Extension 2
                                        {
                                            lcsd2 = lcsd2.AddYears(1);
                                            lcsd2 = lcsd2.AddMonths(6);
                                        }
                                        oForm.Items.Item("tschLEDE").Specific.Value = lcsd2.ToString("yyyyMMdd");
                                    }
                                }
                            }
                        }
                        break;

                    case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
                        // SAPbouiCOM.Form oForms = SBOMain.SBO_Application.Forms.ActiveForm;
                        if (pVal.BeforeAction == true)
                        {
                            if (pVal.ItemUID == "matEXOB" && pVal.ColUID == "titemcode")
                            {
                                CFLCondition("CFL_OITM", pVal.ItemUID);
                            }
                            if (pVal.ItemUID == "matEXIR" && pVal.ColUID == "iritemcd")
                            {
                                CFLCondition("CFL_OITMIR", pVal.ItemUID);
                            }

                            if (pVal.ItemUID == "tschvc")
                            {
                                CFLCondition("CFL_OCRD", pVal.ItemUID);
                            }
                            if (pVal.ItemUID == "tcustcd")
                            {
                                CFLCondition("CFL_OCCRD", pVal.ItemUID);
                            }
                            if (pVal.ItemUID == "tcustcd")
                            {
                                CFLCondition("CFL_OCCRD", pVal.ItemUID);
                            }
                            if (pVal.ItemUID == "tschapin")
                            {
                                string tpvendorcode = oForm.Items.Item("tschvc").Specific.value;
                                CFLConditionInv("CFL_OPCH", pVal.ItemUID, tpvendorcode);

                                //CFLCondition("CFL_OPCH", pVal.ItemUID);
                            }
                            if (pVal.ItemUID == "tarinv")
                            {
                                string tpvendorcode = oForm.Items.Item("tcustcd").Specific.value;
                                CFLConditionInv("CFL_OINV", pVal.ItemUID, tpvendorcode);

                                //CFLCondition("CFL_OINV", pVal.ItemUID);
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
                                    if (pVal.ItemUID == "matEXOB" && pVal.ColUID == "titemcode")
                                    {
                                        SAPbouiCOM.Matrix matrix = oForm.Items.Item("matEXOB").Specific;
                                        ((SAPbouiCOM.EditText)matrix.Columns.Item("titemcode").Cells.Item(pVal.Row).Specific).Value = oDataTable.GetValue("ItemCode", 0).ToString();
                                        ((SAPbouiCOM.EditText)matrix.Columns.Item("titemname").Cells.Item(pVal.Row).Specific).Value = oDataTable.GetValue("ItemName", 0).ToString();
                                        ((SAPbouiCOM.EditText)matrix.Columns.Item("tuom").Cells.Item(pVal.Row).Specific).Value = oDataTable.GetValue("InvntryUom", 0).ToString();
                                        AddMatrixRow(matrix, "titemcode");
                                    }
                                    if (pVal.ItemUID == "matEXIR" && pVal.ColUID == "iritemcd")
                                    {
                                        SAPbouiCOM.Matrix matrix = oForm.Items.Item("matEXIR").Specific;
                                        ((SAPbouiCOM.EditText)matrix.Columns.Item("iritemcd").Cells.Item(pVal.Row).Specific).Value = oDataTable.GetValue("ItemCode", 0).ToString();
                                        ((SAPbouiCOM.EditText)matrix.Columns.Item("iritemnm").Cells.Item(pVal.Row).Specific).Value = oDataTable.GetValue("ItemName", 0).ToString();
                                        ((SAPbouiCOM.EditText)matrix.Columns.Item("iruom").Cells.Item(pVal.Row).Specific).Value = oDataTable.GetValue("InvntryUom", 0).ToString();
                                        AddMatrixRow(matrix, "iritemcd"); 
                                    }
                                    if (pVal.ItemUID == "tschapin")
                                    {
                                        oForm.Items.Item("tschapin").Specific.value = oDataTable.GetValue("DocNum", 0).ToString();
                                        oForm.Items.Item("tschapinde").Specific.value = oDataTable.GetValue("DocEntry", 0).ToString();
                                    }
                                    if (pVal.ItemUID == "tarinv")
                                    {
                                        oForm.Items.Item("tarinv").Specific.value = oDataTable.GetValue("DocNum", 0).ToString();
                                        oForm.Items.Item("arinvde").Specific.value = oDataTable.GetValue("DocEntry", 0).ToString();
                                    }
                                    if (pVal.ItemUID == "texpcode")
                                    {
                                        oForm.Items.Item("texpcode").Specific.value = oDataTable.GetValue("CardCode", 0).ToString();
                                        oForm.Items.Item("tCARDNAME").Specific.value = oDataTable.GetValue("CardName", 0).ToString();
                                    }
                                    if (pVal.ItemUID == "tschvc")
                                    {
                                        oForm.Items.Item("tschvc").Specific.value = oDataTable.GetValue("CardCode", 0).ToString();
                                        oForm.Items.Item("tschvn").Specific.value = oDataTable.GetValue("CardName", 0).ToString();
                                    }
                                    if (pVal.ItemUID == "tcustcd")
                                    {
                                        oForm.Items.Item("tcustcd").Specific.value = oDataTable.GetValue("CardCode", 0).ToString();
                                        oForm.Items.Item("custnm").Specific.value = oDataTable.GetValue("CardName", 0).ToString();
                                    }
                                }
                                catch (Exception ex)
                                {
                                    //  SBOMain.SBO_Application.MessageBox(ex.Message.ToString(), 1, "ok", "", "");
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
                            // VALIDATIONS
                            if (pVal.ItemUID == "1")
                            {
                                if ((oForm.Mode == BoFormMode.fm_ADD_MODE) || (oForm.Mode == BoFormMode.fm_UPDATE_MODE))
                                {
                                    string statusval = oForm.Items.Item("cschtype").Specific.value.ToString();

                                    // BubbleEvent = objCU.IsNullOrEmpty(oForm, "cschtype", "select Licence / Scheme Type");
                                    // BubbleEvent = objCU.IsNullOrEmpty(oForm, "tschno", "Licence / Scheme Type");

                                    if (String.IsNullOrEmpty(oForm.Items.Item("cschtype").Specific.Value.ToString()))
                                    {
                                        SBOMain.SBO_Application.StatusBar.SetText("Please select Licence / Scheme Type", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                        BubbleEvent = false;
                                    }
                                    else if (String.IsNullOrEmpty(oForm.Items.Item("tschno").Specific.Value.ToString()))
                                    {
                                        SBOMain.SBO_Application.StatusBar.SetText("Please insert Scheme No", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                        BubbleEvent = false;
                                    }
                                    else if (String.IsNullOrEmpty(oForm.Items.Item("tschname").Specific.Value.ToString()))
                                    {
                                        SBOMain.SBO_Application.StatusBar.SetText("Please select Scheme Name", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                        BubbleEvent = false;
                                    }
                                    else if (statusval == "DBK" || statusval == "RoDTEP")
                                    {
                                        if (Convert.ToDouble(oForm.Items.Item("tschrate").Specific.Value) <= 0.00)
                                        {
                                            SBOMain.SBO_Application.StatusBar.SetText("Please add Scheme Rate", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                            BubbleEvent = false;
                                        }
                                        else if (String.IsNullOrEmpty(oForm.Items.Item("tschsd").Specific.Value.ToString()))
                                        {
                                            SBOMain.SBO_Application.StatusBar.SetText("Please add Scheme Start Date", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                            BubbleEvent = false;
                                        }
                                        else if (String.IsNullOrEmpty(oForm.Items.Item("tsched").Specific.Value.ToString()))
                                        {
                                            SBOMain.SBO_Application.StatusBar.SetText("Please add Scheme End Date", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                            BubbleEvent = false;
                                        }
                                    }
                                    else if (statusval == "Advanced" || statusval == "EPCG")
                                    {
                                        if (String.IsNullOrEmpty(oForm.Items.Item("tschsd").Specific.Value.ToString()))
                                        {
                                            SBOMain.SBO_Application.StatusBar.SetText("Please add Scheme Start Date", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                            BubbleEvent = false;
                                        }
                                        else if (String.IsNullOrEmpty(oForm.Items.Item("tschLEDI").Specific.Value.ToString()))
                                        {
                                            SBOMain.SBO_Application.StatusBar.SetText("Please add Licence Expiry Date for Import", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                            BubbleEvent = false;
                                        }
                                        else if (String.IsNullOrEmpty(oForm.Items.Item("tschLEDE").Specific.Value.ToString()))
                                        {
                                            SBOMain.SBO_Application.StatusBar.SetText("Please add Licence Expiry Date for Export", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                            BubbleEvent = false;
                                        }
                                    }

                                }
                            }

                            if (pVal.ItemUID == "cschTPLP")
                            {

                                oCheckbox = oForm.Items.Item("cschTPLP").Specific;
                                if (oCheckbox.Checked && (oForm.Mode == BoFormMode.fm_UPDATE_MODE)
                                          && !string.IsNullOrEmpty(oForm.Items.Item("tschapinde").Specific.Value.ToString()))
                                {
                                    SBOMain.SBO_Application.StatusBar.SetText("A/P invoice already linked in third party license purchase.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                    BubbleEvent = false;
                                }
                            }
                            if (pVal.ItemUID == "TPLS")
                            {
                                oCheckbox = oForm.Items.Item("TPLS").Specific;
                                if (oCheckbox.Checked && (oForm.Mode == BoFormMode.fm_UPDATE_MODE)
                                          && !string.IsNullOrEmpty(oForm.Items.Item("arinvde").Specific.Value.ToString()))
                                {
                                    SBOMain.SBO_Application.StatusBar.SetText("A/R invoice already linked in third party license sale.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                    BubbleEvent = false;
                                }
                            }

                        }
                        if (pVal.BeforeAction == false)
                        { 
                            if (pVal.ItemUID == "btnDIS")
                            {
                                SAPbouiCOM.Matrix omatLCLD = (SAPbouiCOM.Matrix)oForm.Items.Item("matATTACH").Specific;
                                int rowid = omatLCLD.GetNextSelectedRow(); // getnextselectedrow(0, roworder)

                                string a = omatLCLD.Columns.Item("trgtpath").Cells.Item(rowid).Specific.Value;
                                string b = omatLCLD.Columns.Item("filename").Cells.Item(rowid).Specific.Value;

                                string fullpath = a + b;
                                System.Diagnostics.Process.Start(@fullpath);
                            }

                            if (pVal.ItemUID == "1" && (oForm.Mode == BoFormMode.fm_FIND_MODE))
                            {
                                SchemeRateAndPercKGFields(oForm, "cschtype");
                            }
                                if (pVal.ItemUID == "1" && (oForm.Mode == BoFormMode.fm_ADD_MODE || oForm.Mode == BoFormMode.fm_UPDATE_MODE))
                            { 
                                SAPbouiCOM.Matrix matrix = oForm.Items.Item("matATTACH").Specific;

                                for (int i = 1; i <= matrix.RowCount; i++)
                                {
                                    string TargetPath = (matrix.Columns.Item("trgtpath").Cells.Item(i).Specific).Value;
                                    string FileName = (matrix.Columns.Item("filename").Cells.Item(i).Specific).Value;
                                    if (!File.Exists(TargetPath + FileName))
                                    {
                                        using (Stream s = File.Open(TargetPath + FileName, FileMode.Create))
                                        {
                                            using (StreamWriter sw = new StreamWriter(s))
                                            {
                                                sw.Write(BrowseFilePath);
                                            }
                                        }
                                    }
                                } 
                            }
                            
                            if (pVal.ItemUID == "matEXFL" && pVal.ColUID == "expexmno")
                            {
                                SAPbouiCOM.Matrix matrix = oForm.Items.Item("matEXFL").Specific;
                                string abc = (matrix.Columns.Item("expexmno").Cells.Item(pVal.Row).Specific).Value;
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
                                OutwardToEximTracking outEximTracking = new OutwardToEximTracking();

                                outEximTracking.DocEntry = abc;
                                outEximTracking.FromFrmName = "FindMode";

                                clsExTrans oPrice = new clsExTrans(outEximTracking);
                                //oForm.Close();
                            }

                            if (pVal.ItemUID == "btnRef") // Refresh Button
                            {
                                SAPbouiCOM.ComboBox cb4 = oForm.Items.Item("cschtype").Specific;
                                string val = cb4.Selected.Value.ToString();
                                if (val == "Advanced" || val == "EPCG")
                                {
                                    SAPbouiCOM.Matrix matrix = oForm.Items.Item("matEXFL").Specific;
                                    DeleteMatrixBlankRow(matrix, "expexmno");
                                    ArrengeMatrixLineNum(matrix);

                                    int i = 1;
                                    string liceneno = oForm.Items.Item("tschno").Specific.value;
                                    string Query = "SELECT T2.InvntryUom , T0.DocEntry, T0.U_exdt, T0.U_exinvno, T0.U_exinvnode, T1.U_ex4lfqty, T1.U_ex4lflc, T1.U_ex4lffc, T0.U_exdd, T0.U_exsbn,T0.U_exsbd FROM dbo.[@EXET] T0 LEFT JOIN dbo.[@XET4] AS T1 on ";
                                    Query = Query + " T0.DocEntry = T1.DocEntry LEFT JOIN OITM AS T2 on T1.U_ex4ic = T2.ItemCode WHERE T1.U_ex4ln = '" + liceneno + "' and T0.Status = 'O' and T0.U_exinvnode not in ( SELECT ";
                                    Query = Query + " T1.U_expinvde from dbo.[@EXSM] AS T0 LEFT JOIN dbo.[@XSM2] AS T1 On T0.Code = T1.Code where T1.U_expinvde IS NOT NULL AND T0.U_schtype = '" + val + "'";
                                    Query = Query + " AND T0.U_schno = '" + liceneno + "')";
                                    SAPbobsCOM.Recordset rec = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                    rec.DoQuery(Query);
                                    if (rec.RecordCount > 0)
                                    {
                                        while (!rec.EoF)
                                        {
                                            matrix.AddRow();
                                            (matrix.Columns.Item("#").Cells.Item(matrix.RowCount).Specific).Value = Convert.ToString(matrix.RowCount);
                                            (matrix.Columns.Item("expexmno").Cells.Item(matrix.RowCount).Specific).Value = rec.Fields.Item("DocEntry").Value;
                                            (matrix.Columns.Item("expinvde").Cells.Item(matrix.RowCount).Specific).Value = rec.Fields.Item("U_exinvnode").Value;
                                            (matrix.Columns.Item("texpinvno").Cells.Item(matrix.RowCount).Specific).Value = rec.Fields.Item("U_exinvno").Value;
                                            (matrix.Columns.Item("texpinvdt").Cells.Item(matrix.RowCount).Specific).Value = rec.Fields.Item("U_exdd").Value.ToString("yyyyMMdd");
                                            (matrix.Columns.Item("expbn").Cells.Item(matrix.RowCount).Specific).Value = rec.Fields.Item("U_exsbn").Value; //.ToString("yyyyMMdd");
                                            (matrix.Columns.Item("texpbd").Cells.Item(matrix.RowCount).Specific).Value = rec.Fields.Item("U_exsbd").Value.ToString("yyyyMMdd");
                                            (matrix.Columns.Item("expfq").Cells.Item(matrix.RowCount).Specific).Value = rec.Fields.Item("U_ex4lfqty").Value;
                                            (matrix.Columns.Item("UOM").Cells.Item(matrix.RowCount).Specific).Value = rec.Fields.Item("InvntryUom").Value;
                                            (matrix.Columns.Item("expfvFC").Cells.Item(matrix.RowCount).Specific).Value = rec.Fields.Item("U_ex4lffc").Value;
                                            (matrix.Columns.Item("expfvLC").Cells.Item(matrix.RowCount).Specific).Value = rec.Fields.Item("U_ex4lflc").Value;
                                            i++;
                                            rec.MoveNext();
                                        }
                                    }  
                                    
                                    /////// Export Obligation Grid

                                    SAPbouiCOM.Matrix matEXOB = oForm.Items.Item("matEXOB").Specific;
                                    if (matEXOB.RowCount > 0)
                                    {
                                        for (i = 1; i <= matEXOB.RowCount; i++)
                                        {
                                            string itemCode = (matEXOB.Columns.Item("titemcode").Cells.Item(i).Specific).Value;
                                            double tqty = Convert.ToDouble(matEXOB.Columns.Item("tqty").Cells.Item(i).Specific.Value);
                                            double tamtLC = Convert.ToDouble(matEXOB.Columns.Item("tamtLC").Cells.Item(i).Specific.Value);
                                            double tamtFC = Convert.ToDouble(matEXOB.Columns.Item("tamtFC").Cells.Item(i).Specific.Value);
                                            double tflqty = Convert.ToDouble(matEXOB.Columns.Item("tflqty").Cells.Item(i).Specific.Value);
                                            double tfllC = Convert.ToDouble(matEXOB.Columns.Item("tfllC").Cells.Item(i).Specific.Value);
                                            double tflfC = Convert.ToDouble(matEXOB.Columns.Item("tflfC").Cells.Item(i).Specific.Value);
                                            double rmqty = Convert.ToDouble(matEXOB.Columns.Item("rmqty").Cells.Item(i).Specific.Value);
                                            double rmLC = Convert.ToDouble(matEXOB.Columns.Item("rmLC").Cells.Item(i).Specific.Value);
                                            double rmFC = Convert.ToDouble(matEXOB.Columns.Item("rmFC").Cells.Item(i).Specific.Value);

                                            Query = "SELECT SUM(T1.U_ex4lfqty) AS 'Quantiry',  SUM(T1.U_ex4lffc) As 'FCAmt',  SUM(T1.U_ex4lflc) As 'LCAmt'  FROM dbo.[@EXET]  AS T0 LEFT JOIN dbo.[@XET4] AS T1 ON T0.DocEntry = T1.DocEntry";
                                            Query = Query + " WHERE T1.U_ex4ln = '" + liceneno + "' and T1.U_ex4ic = '" + itemCode + "'";

                                            SAPbobsCOM.Recordset rec1 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                            rec1.DoQuery(Query);
                                            if (rec1.RecordCount > 0)
                                            {
                                                double ex4lfqty = Convert.ToDouble(rec1.Fields.Item("Quantiry").Value);
                                                double ex4lffc = Convert.ToDouble(rec1.Fields.Item("FCAmt").Value);
                                                double ex4lflc = Convert.ToDouble(rec1.Fields.Item("LCAmt").Value);
                                                rmqty = tqty - ex4lfqty;// tflqty;
                                                rmLC = tamtLC - ex4lflc; // tfllC;
                                                rmFC = tamtFC - ex4lffc; // tflfC;

                                                (matEXOB.Columns.Item("tflqty").Cells.Item(i).Specific).Value = ex4lfqty; // tflqty;
                                                (matEXOB.Columns.Item("tfllC").Cells.Item(i).Specific).Value = ex4lflc; // tfllC;
                                                (matEXOB.Columns.Item("tflfC").Cells.Item(i).Specific).Value = ex4lffc; // tflfC;
                                            }

                                            (matEXOB.Columns.Item("rmqty").Cells.Item(i).Specific).Value = rmqty;
                                            (matEXOB.Columns.Item("rmLC").Cells.Item(i).Specific).Value = rmLC;
                                            (matEXOB.Columns.Item("rmFC").Cells.Item(i).Specific).Value = rmFC;

                                        }
                                    }


                                }
                            }

                            if (pVal.ItemUID == "tab2")
                            {
                                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("matEXFL").Specific;
                                doAutoSummatEXFL(oForm);
                            }
                            if (pVal.ItemUID == "tab3")
                            {
                                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("matEXIR").Specific;
                                AddMatrixRow(oMatrix, "iritemcd");
                                doAutoSummatEXIR(oForm);
                            }
                            if (pVal.ItemUID == "tab4")
                            {
                                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("matEXIU").Specific;
                                AddMatrixRow(oMatrix, "impinvno");
                                doAutoSummatEXIU(oForm);
                            }
                            if (pVal.ItemUID == "tab5")
                            {
                                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("matHSN").Specific;
                                AddMatrixRow(oMatrix, "hsncode");
                            }
                            if (pVal.ItemUID == "cschTPLP")
                            {
                                oCheckbox = oForm.Items.Item("cschTPLP").Specific;
                                if (oCheckbox.Checked)
                                {
                                    oForm.Freeze(true);
                                    oForm.Items.Item("tschvc").Specific.value = "";
                                    oForm.Items.Item("tschvn").Specific.value = "";
                                    oForm.Items.Item("tschapin").Specific.value = "";
                                    oForm.Items.Item("tschapinde").Specific.value = ""; 
                                    thirdPartyLicencePurchase(oForm, SAPbouiCOM.BoModeVisualBehavior.mvb_False);
                                    oForm.Freeze(false);

                                }
                                else
                                {
                                    oForm.Freeze(true);
                                    thirdPartyLicencePurchase(oForm, SAPbouiCOM.BoModeVisualBehavior.mvb_True);
                                    oForm.Freeze(false);
                                }
                            }
                            if (pVal.ItemUID == "TPLS")
                            {
                                oCheckbox = oForm.Items.Item("TPLS").Specific;
                                if (oCheckbox.Checked)
                                {
                                    oForm.Freeze(true);
                                    oForm.Items.Item("tcustcd").Specific.value = "";
                                    oForm.Items.Item("custnm").Specific.value = "";
                                    oForm.Items.Item("tarinv").Specific.value = "";
                                    oForm.Items.Item("arinvde").Specific.value = "";
                                    thirdPartyLicenceSales(oForm, SAPbouiCOM.BoModeVisualBehavior.mvb_False);
                                    oForm.Freeze(false);
                                }
                                else
                                {
                                    oForm.Freeze(true);
                                    thirdPartyLicenceSales(oForm, SAPbouiCOM.BoModeVisualBehavior.mvb_True);
                                    oForm.Items.Item("tcustcd").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, (int)BoFormMode.fm_ADD_MODE, SAPbouiCOM.BoModeVisualBehavior.mvb_True);
                                    oForm.Items.Item("custnm").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, (int)BoFormMode.fm_ADD_MODE, SAPbouiCOM.BoModeVisualBehavior.mvb_True);
                                    oForm.Items.Item("tarinv").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, (int)BoFormMode.fm_ADD_MODE, SAPbouiCOM.BoModeVisualBehavior.mvb_True);
                                    oForm.Items.Item("arinvde").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, (int)BoFormMode.fm_ADD_MODE, SAPbouiCOM.BoModeVisualBehavior.mvb_True);
                                    oForm.Items.Item("lblcustcd").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, (int)BoFormMode.fm_ADD_MODE, SAPbouiCOM.BoModeVisualBehavior.mvb_True);
                                    oForm.Items.Item("lblcustnm").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, (int)BoFormMode.fm_ADD_MODE, SAPbouiCOM.BoModeVisualBehavior.mvb_True);
                                    oForm.Items.Item("lblarinv").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, (int)BoFormMode.fm_ADD_MODE, SAPbouiCOM.BoModeVisualBehavior.mvb_True);
                                    oForm.Freeze(false);
                                }

                            }
                        }
                        break;
                    case BoEventTypes.et_ITEM_PRESSED:
                        oForm = SBOMain.SBO_Application.Forms.ActiveForm;

                        if (pVal.BeforeAction == true)
                        {

                        }
                        if (pVal.BeforeAction == false)
                        {
                            if (pVal.ItemUID == "1" && oForm.Mode == BoFormMode.fm_OK_MODE)
                            {
                                SchemeRateAndPercKGFields(oForm, "cschtype");
                            }
                            if (pVal.ItemUID == "1" && (oForm.Mode == BoFormMode.fm_ADD_MODE || oForm.Mode == BoFormMode.fm_UPDATE_MODE))
                            {
                                SAPbouiCOM.Matrix matrix = oForm.Items.Item("matLCATT").Specific;
                                for (int i = 0; i < BrowseFilePath.Count; i++)
                                {
                                    string a = Convert.ToString(BrowseFilePath[i]);
                                    string b = Convert.ToString(ReplaceFilePath[i]);
                                    File.Move(a, b);
                                }
                                BrowseFilePath.Clear();
                                ReplaceFilePath.Clear(); 
                                Form_Load_Components(oForm, "OK");
                            }
                            if (pVal.ItemUID == "btnATC")
                            {
                                OpenFile();
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
                //if (oForm != null)
                //  oForm.Freeze(false);
            }


            return BubbleEvent;
        }

        public void thirdPartyLicencePurchase(SAPbouiCOM.Form oForm, SAPbouiCOM.BoModeVisualBehavior Value)
        { 
            oForm.Items.Item("tschvc").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, (int)BoFormMode.fm_ADD_MODE, Value);
            oForm.Items.Item("tschvn").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, (int)BoFormMode.fm_ADD_MODE, Value);
            oForm.Items.Item("tschapin").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, (int)BoFormMode.fm_ADD_MODE, Value);
            oForm.Items.Item("tschapinde").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, (int)BoFormMode.fm_ADD_MODE, Value);
            oForm.Items.Item("lblschvc").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, (int)BoFormMode.fm_ADD_MODE, Value);
            oForm.Items.Item("lblschvn").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, (int)BoFormMode.fm_ADD_MODE, Value);
            oForm.Items.Item("lblschapin").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, (int)BoFormMode.fm_ADD_MODE, Value); 
        }
        public void thirdPartyLicenceSales(SAPbouiCOM.Form oForm, SAPbouiCOM.BoModeVisualBehavior Value)
        {
            oForm.Items.Item("tcustcd").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, (int)BoFormMode.fm_ADD_MODE, Value);
            oForm.Items.Item("custnm").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, (int)BoFormMode.fm_ADD_MODE, Value);
            oForm.Items.Item("tarinv").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, (int)BoFormMode.fm_ADD_MODE, Value);
            oForm.Items.Item("arinvde").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, (int)BoFormMode.fm_ADD_MODE, Value);
            oForm.Items.Item("lblcustcd").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, (int)BoFormMode.fm_ADD_MODE, Value);
            oForm.Items.Item("lblcustnm").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, (int)BoFormMode.fm_ADD_MODE, Value);
            oForm.Items.Item("lblarinv").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, (int)BoFormMode.fm_ADD_MODE, Value);

        }
        public void doAutoColSum(SAPbouiCOM.Matrix matrix, string ColumnName)
        {
            SAPbouiCOM.Column mCol = matrix.Columns.Item(ColumnName);
            mCol.RightJustified = true;
            mCol.ColumnSetting.SumType = BoColumnSumType.bst_Auto;
        }

        public void OpenFile()
        {
            try
            {
                System.Threading.Thread ShowFolderBrowserThread;
                ShowFolderBrowserThread = new System.Threading.Thread(ShowFolderBrowser);
                if (ShowFolderBrowserThread.ThreadState == System.Threading.ThreadState.Unstarted)
                {
                    ShowFolderBrowserThread.SetApartmentState(ApartmentState.STA);
                    ShowFolderBrowserThread.Start();
                }

                else if (ShowFolderBrowserThread.ThreadState == System.Threading.ThreadState.Stopped)
                {

                    ShowFolderBrowserThread.Start();
                    ShowFolderBrowserThread.Join();

                }

            }

            catch (Exception ex)
            {
                // SBOMain.SBO_Application.MessageBox(ex.Message.ToString(), 1, "ok", "", "");

            }

        }

        public void ShowFolderBrowser()
        {
            try
            {
                oForm = SBOMain.SBO_Application.Forms.Item(cFormID);
                NativeWindow nws = new NativeWindow();
                //System.Windows.Forms.Form nws = new Form();
                OpenFileDialog MyTest = new OpenFileDialog();
                Process[] MyProcs = null;
                string filename = null;
                 
                MyProcs = Process.GetProcessesByName("SAP Business One");
                nws.AssignHandle(System.Diagnostics.Process.GetProcessesByName("SAP Business One")[0].MainWindowHandle);

                IntPtr ptr = GetForegroundWindow();
                WindowWrapper oWindow = new WindowWrapper(ptr);

                if (MyTest.ShowDialog(oWindow) == System.Windows.Forms.DialogResult.OK)
                {
                    filename = MyTest.FileName;
                    SAPbouiCOM.Matrix matrix = oForm.Items.Item("matATTACH").Specific;

                    BrowseFilePath.Add(filename);
                    ReplaceFilePath.Add(SBOMain.Get_Attach_Folder_Path() + MyTest.SafeFileName);

                    matrix.AddRow();
                    (matrix.Columns.Item("#").Cells.Item(matrix.RowCount).Specific).Value = Convert.ToString(matrix.RowCount);
                    (matrix.Columns.Item("trgtpath").Cells.Item(matrix.RowCount).Specific).Value = SBOMain.Get_Attach_Folder_Path(); // rec.Fields.Item("DocEntry").Value;
                    (matrix.Columns.Item("filename").Cells.Item(matrix.RowCount).Specific).Value = MyTest.SafeFileName.ToString(); // rec.Fields.Item("U_exinvnode").Value;
                    (matrix.Columns.Item("atchdate").Cells.Item(matrix.RowCount).Specific).Value = DateTime.Today.ToString("yyyyMMdd"); ; // rec.Fields.Item("U_exinvno").Value;
                    (matrix.Columns.Item("fretext").Cells.Item(matrix.RowCount).Specific).Value = null;  // rec.Fields.Item("U_exdd").Value.ToString("yyyyMMdd");
                    (matrix.Columns.Item("cpytotd").Cells.Item(matrix.RowCount).Specific).Value = null;  // rec.Fields.Item("U_exsbn").Value; //.ToString("yyyyMMdd");

                    //oForm.Items.Item("tattach").Specific.value = filename;

                    System.Windows.Forms.Application.ExitThread();
                }
            }
            catch (Exception ex)
            {
                SBOMain.SBO_Application.MessageBox(ex.Message.ToString(), 1, "ok", "", "");
            }
        }

        private void SetCode()
        {
            oForm.Freeze(true);
            oForm.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized;
            string TableName = "EXSM";
            SBOMain.SetCode(oForm.UniqueID, TableName);
            oForm.Freeze(false);
            //throw new NotImplementedException();
        }

        #region MatrixSetLine

        public void ADDROWMain(SAPbouiCOM.Matrix oMatrix)
        {
            oMatrix.AddRow(1, SBOMain.RightClickLineNum);
            oMatrix.ClearRowData(SBOMain.RightClickLineNum + 1);
            ArrengeMatrixLineNum(oMatrix);
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
        private void ArrengeMatrixLineNum(SAPbouiCOM.Matrix matrix)
        {
            for (int i = 1; i <= matrix.RowCount; i++)
            {
                matrix.Columns.Item("#").Cells.Item(i).Specific.value = i;
            }
        }

        #endregion

        public void SchemeRateAndPercKGFields(SAPbouiCOM.Form oForm, string field)
        {
            string statusval = oForm.Items.Item(field).Specific.value.ToString();
            if (statusval == "DBK" || statusval == "RoDTEP")
            {
                oForm.Items.Item("tschrate").Enabled = true;
                oForm.Items.Item("tschratepk").Enabled = true;
                oForm.Items.Item("tschLEDI").Enabled = false;
                oForm.Items.Item("tschLEDE").Enabled = false;
            }
            if (statusval == "Advanced" || statusval == "EPCG")
            {
                oForm.Items.Item("tschrate").Enabled = false;
                oForm.Items.Item("tschratepk").Enabled = false;
                oForm.Items.Item("tschLEDI").Enabled = true;
                oForm.Items.Item("tschLEDE").Enabled = true;
            }
        }

        public void Form_Load_Components(SAPbouiCOM.Form oForm, string mode)
        {
            //oForm.Items.Item("tattach").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, (int)BoFormMode.fm_ADD_MODE, SAPbouiCOM.BoModeVisualBehavior.mvb_False);
            oForm.Items.Item("tab1").Visible = true;
            oForm.Items.Item("tab1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
            oForm.PaneLevel = 1;
            SetCode();

            oMatrix = oForm.Items.Item("matEXOB").Specific;
            AddMatrixRow(oMatrix, "titemcode");
            doAutoSummatEXOB(oForm);

            SAPbouiCOM.ComboBox cb1 = (SAPbouiCOM.ComboBox)oForm.Items.Item("cschtype").Specific;
            cb1.ExpandType = BoExpandType.et_DescriptionOnly;
               
            SAPbouiCOM.ComboBox cb3 = (SAPbouiCOM.ComboBox)oForm.Items.Item("cext").Specific;
            cb3.ExpandType = BoExpandType.et_DescriptionOnly;

            SAPbouiCOM.ComboBox cb4 = (SAPbouiCOM.ComboBox)oForm.Items.Item("tstatus").Specific;
            cb4.ExpandType = BoExpandType.et_DescriptionOnly;

            if (oForm.Mode == BoFormMode.fm_ADD_MODE || oForm.Mode == BoFormMode.fm_FIND_MODE)
            {
                string statusval = oForm.Items.Item("cschtype").Specific.value.ToString();
                if (statusval == "Advanced" || statusval == "EPCG")
                {
                    oForm.Items.Item("tschrate").Enabled = false;
                    oForm.Items.Item("tschratepk").Enabled = false;
                    oForm.Items.Item("tschLEDI").Enabled = true;
                    oForm.Items.Item("tschLEDE").Enabled = true;
                }
                else
                {
                    oForm.Items.Item("tschrate").Enabled = true;
                    oForm.Items.Item("tschratepk").Enabled = true;
                    oForm.Items.Item("tschLEDI").Enabled = false;
                    oForm.Items.Item("tschLEDE").Enabled = false;
                }
            } 

            oComboBoxStatus = oForm.Items.Item("tstatus").Specific;
            oComboBoxStatus.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;
            oComboBoxStatus.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
            oForm.Items.Item("tstatus").DisplayDesc = true;

        }
        public void doAutoSummatEXFL(SAPbouiCOM.Form oForm)
        {
            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("matEXFL").Specific;
            doAutoColSum(oMatrix, "expfq");
            doAutoColSum(oMatrix, "expfvFC");
            doAutoColSum(oMatrix, "expfvLC");
        }
        public void doAutoSummatEXIU(SAPbouiCOM.Form oForm)
        {
            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("matEXIU").Specific;
            AddMatrixRow(oMatrix, "impinvno");
            doAutoColSum(oMatrix, "impq");
            doAutoColSum(oMatrix, "impvFC");
            doAutoColSum(oMatrix, "impvLC");
        }
        public void doAutoSummatEXIR(SAPbouiCOM.Form oForm)
        {
            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("matEXIR").Specific;
            doAutoColSum(oMatrix, "irqty");
            doAutoColSum(oMatrix, "iramtLC");
            doAutoColSum(oMatrix, "iramtFC");
            doAutoColSum(oMatrix, "irutqty");
            doAutoColSum(oMatrix, "irutlC");
            doAutoColSum(oMatrix, "irutfC");
            doAutoColSum(oMatrix, "irrmqty");
            doAutoColSum(oMatrix, "irrmLC");
            doAutoColSum(oMatrix, "irrmFC");
        }
        private void CFLConditionInv(string CFLID, string ItemUID, string CardCode)
        {
            oCFL = oForm.ChooseFromLists.Item(CFLID);
            oConds = SBOMain.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions);
            if (CFLID == "CFL_OPCH" || CFLID == "CFL_OINV")
            {

                oCond = oConds.Add();
                oCond.Alias = "CardCode";
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCond.CondVal = CardCode;
                oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

                oCond = oConds.Add();
                oCond.Alias = "DocStatus";
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCond.CondVal = "O";
                oCFL.SetConditions(oConds);
            }
            oCFL = null;
            oCond = null;
            oConds = null;

        }
        private void CFLCondition(string CFLID, string ItemUID)
        {
            oCFL = oForm.ChooseFromLists.Item(CFLID);
            oConds = SBOMain.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions);

            if (CFLID == "CFL_OITM" || CFLID == "CFL_OITMIR")
            {
                oCond = oConds.Add();
                oCond.Alias = "InvntItem";
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCond.CondVal = "N";
                oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

                oCond = oConds.Add();
                oCond.Alias = "SellItem";
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCond.CondVal = "N";
                oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

                oCond = oConds.Add();
                oCond.Alias = "PrchseItem";
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCond.CondVal = "N";
                oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

                oCond = oConds.Add();
                oCond.Alias = "ItemClass";
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCond.CondVal = "2";
                oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

                oCond = oConds.Add();
                oCond.Alias = "ItemCode";
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_START;
                oCond.CondVal = "LIC";

                oCFL.SetConditions(oConds);

            }
            if (CFLID == "CFL_OCCRD")
            {

                oCond = oConds.Add();
                oCond.Alias = "CardType";
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCond.CondVal = "C";
                oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

                oCond = oConds.Add();
                oCond.Alias = "validFor";
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCond.CondVal = "Y";

                oCFL.SetConditions(oConds);

            }

            if (CFLID == "CFL_OCRD")
            {
                oCond = oConds.Add();
                oCond.Alias = "CardType";
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCond.CondVal = "S";

                oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

                oCond = oConds.Add();
                oCond.Alias = "validFor";
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCond.CondVal = "Y";

                oCFL.SetConditions(oConds);

            }

            oCFL = null;
            oCond = null;
            oConds = null;

        }
        private void DeleteMatrixAll(SAPbouiCOM.Matrix oMatrix)
        {
            try
            {
                for (int i = oMatrix.VisualRowCount; i >= 1; i--)
                {
                    oMatrix.DeleteRow(i);
                }
            }
            catch (Exception ex)
            {
                SBOMain.SBO_Application.MessageBox(ex.Message.ToString(), 1, "ok", "", "");
            }
        }
        public void doAutoSummatEXOB(SAPbouiCOM.Form oForm)
        {
            SAPbouiCOM.Matrix matEXOB = (SAPbouiCOM.Matrix)oForm.Items.Item("matEXOB").Specific;
            doAutoColSum(matEXOB, "tqty");
            doAutoColSum(matEXOB, "tamtLC");
            doAutoColSum(matEXOB, "tamtFC");
            doAutoColSum(matEXOB, "tflqty");
            doAutoColSum(matEXOB, "tfllC");
            doAutoColSum(matEXOB, "tflfC");
            doAutoColSum(matEXOB, "rmqty");
            doAutoColSum(matEXOB, "rmLC");
            doAutoColSum(matEXOB, "rmFC");
        }
    }
}
