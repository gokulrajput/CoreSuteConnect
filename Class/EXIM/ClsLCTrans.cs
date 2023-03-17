using System;
using System.Collections;
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
using System.Collections.Specialized;
using CoreSuteConnect.Events;
using CoreSuteConnect.Class.DEFAULTSAPFORMS;
 
namespace CoreSuteConnect.Class.EXIM
{
    class ClsLCTrans
    {
        [DllImport("user32.dll")]
        
        private static extern IntPtr GetForegroundWindow();

        #region VariableDeclaration
        private SAPbouiCOM.Form oForm;
        private SAPbouiCOM.Matrix oMatrix;
        public string cFormID = string.Empty;
         
        private SAPbouiCOM.ChooseFromList oCFL = null;
        private SAPbouiCOM.Conditions oConds = null;
        private SAPbouiCOM.Condition oCond = null;

        OpenFileDialog OpenFileDialog = new OpenFileDialog();
       
        ArrayList BrowseFilePath = new ArrayList();
        ArrayList ReplaceFilePath = new ArrayList();

        SAPbouiCOM.EditText lc5bdn;

        SAPbouiCOM.ChooseFromList CFL_PR, CFL_PQ, CFL_PO, CFL_PI;

        CommonUtility objCU = new CommonUtility();

        double rate = 0;
        //string currency = null;
        int rownum = 0;
        string rowyear = null;
        int rowmonth = 0;
       // int rowday = 0;

        #endregion VariableDeclaration
        public ClsLCTrans(OutwardToLCMaster inClass)
        {
            if (inClass != null)
            {
                oForm = SBOMain.SBO_Application.Forms.ActiveForm;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                oForm.Items.Item("lcln").Specific.value = inClass.lcno;
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
                        String Q1 = "SELECT count(*) as total  FROM dbo.[@EXET] where U_exlcno = '" + oForm.Items.Item("tCode").Specific.Value + "'";
                        SAPbobsCOM.Recordset rec1 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                        rec1.DoQuery(Q1);
                        int cnt1 = Convert.ToInt32(rec1.Fields.Item("total").Value);

                           if (cnt1 > 0)
                        {
                            SBOMain.SBO_Application.StatusBar.SetText("Remove operation not allowed because this LC is assigned in Exim Transaction", BoMessageTime.bmt_Long, BoStatusBarMessageType.smt_Error);
                            BubbleEvent = false;
                        }
                        
                    }
                }
                if (pVal.BeforeAction == false)
                {
                     
                    if ((oForm.Mode == BoFormMode.fm_ADD_MODE || oForm.Mode == BoFormMode.fm_UPDATE_MODE || Type == "ADDNEWFORM") && (Type != "DEL_ROW") && (Type != "ADD_ROW"))
                    {
                        Form_Load_Components(oForm, "Add");
                    }
                    if (Type == "navigation")
                    {   
                        doAutoSummatLCLD(oForm);
                        doAutoSummatLCEX(oForm);
                    }
                    else if (Type == "DEL_ROW" || Type == "ADD_ROW")
                    {
                        SAPbouiCOM.Matrix matLCLD = (SAPbouiCOM.Matrix)oForm.Items.Item("matLCLD").Specific;
                        SAPbouiCOM.Matrix matLCDOC = (SAPbouiCOM.Matrix)oForm.Items.Item("matLCDOC").Specific; 
                        SAPbouiCOM.Matrix matLCEX = (SAPbouiCOM.Matrix)oForm.Items.Item("matLCEX").Specific;
                        SAPbouiCOM.Matrix matLCAMED = (SAPbouiCOM.Matrix)oForm.Items.Item("matLCAMED").Specific;
                        SAPbouiCOM.Matrix matLCATT = (SAPbouiCOM.Matrix)oForm.Items.Item("matLCATT").Specific;

                        if (Type == "ADD_ROW")
                        {
                            if (SBOMain.RightClickItemID == "matLCLD"){
                                ADDROWMain(matLCLD);
                            } 
                            else if (SBOMain.RightClickItemID == "matLCDOC"){
                                ADDROWMain(matLCDOC);
                            } 
                            else if (SBOMain.RightClickItemID == "matLCEX"){
                                ADDROWMain(matLCEX);
                            } 
                            else if (SBOMain.RightClickItemID == "matLCAMED"){
                                ADDROWMain(matLCAMED);
                            }
                            else if (SBOMain.RightClickItemID == "matLCATT"){
                                ADDROWMain(matLCATT);
                            }
                        }
                        if (Type == "DEL_ROW")
                        {
                            if (SBOMain.RightClickItemID == "matLCLD")
                            {    
                                DeleteMatrixBlankRow(matLCLD, "lc2dde");
                                ArrengeMatrixLineNum(matLCLD);
                            }
                            else if (SBOMain.RightClickItemID == "matLCDOC")
                            {
                                DeleteMatrixBlankRow(matLCDOC, "lc4doc");
                                ArrengeMatrixLineNum(matLCDOC); 
                            }
                            else if (SBOMain.RightClickItemID == "matLCEX")
                            {
                                DeleteMatrixBlankRow(matLCEX, "lc5expt");
                                ArrengeMatrixLineNum(matLCEX); 
                            }
                            else if (SBOMain.RightClickItemID == "matLCAMED")
                            {
                                DeleteMatrixBlankRow(matLCAMED, "lc6amdn");
                                ArrengeMatrixLineNum(matLCAMED); 
                            }
                            else if (SBOMain.RightClickItemID == "matLCATT")
                            {
                                DeleteMatrixBlankRow(matLCATT, "lc7tp");
                                ArrengeMatrixLineNum(matLCATT); 
                            }
                        }
                    }
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

        public bool ItemEvent(ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent, string FormId)
        {
            BubbleEvent = true;
            oForm = SBOMain.SBO_Application.Forms.Item(FormId);
            try
            {
                switch (pVal.EventType)
                {   
                    case BoEventTypes.et_COMBO_SELECT:
                        oForm = SBOMain.SBO_Application.Forms.ActiveForm;
                        if (pVal.BeforeAction == false)
                        {
                            if (pVal.ItemUID == "matLCEX" && pVal.ColUID == "lc5yc")
                            {
                                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                                { 
                                    BubbleEvent = false;
                                    SBOMain.SBO_Application.StatusBar.SetText("Please save LC", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                                }
                                else if(oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                                {    
                                    Program.LCTransData.LcNo = oForm.Items.Item("lcln").Specific.Value;
                                    SAPbouiCOM.Matrix matLCEX = (SAPbouiCOM.Matrix)oForm.Items.Item("matLCEX").Specific;
                                    SAPbouiCOM.ComboBox cb4 = matLCEX.Columns.Item("lc5yc").Cells.Item(pVal.Row).Specific; // oForm.Items.Item("cmbCRCY").Specific;
                                    string cmbcal = cb4.Selected.Value.ToString();
                                    //SAPbouiCOM.LinkedButton oLinkExpdoc = matLCEX.Columns.Item("lc5bden").Cells.Item(pVal.Row).Specific;
                                     
                                    SBOMain.sForm = "LC";
                                    if (cmbcal == "PR")
                                    {    
                                        setChooseFromListField(oForm, CFL_PR, "CFL_PR", lc5bdn, "lc5bden", matLCEX , pVal.Row);
                                        SBOMain.LineNo = pVal.Row;
                                        SBOMain.TFromUID = oForm.UniqueID;
                                        SBOMain.FromCnt = oForm.TypeCount;
                                        //oLinkExpdoc.LinkedObjectType = "1470000113";
                                         
                                    }
                                    else if (cmbcal == "PQ")
                                    {
                                        setChooseFromListField(oForm, CFL_PQ, "CFL_PQ", lc5bdn, "lc5bden", matLCEX, pVal.Row);
                                        SBOMain.LineNo = pVal.Row;
                                        SBOMain.TFromUID = oForm.UniqueID;
                                        SBOMain.FromCnt = oForm.TypeCount;
                                       // oLinkExpdoc.LinkedObjectType = "540000006";
                                    }
                                    else if (cmbcal == "PO")
                                    {
                                        setChooseFromListField(oForm, CFL_PO, "CFL_PO", lc5bdn, "lc5bden", matLCEX, pVal.Row);
                                        SBOMain.LineNo = pVal.Row;
                                        SBOMain.TFromUID = oForm.UniqueID;
                                        SBOMain.FromCnt = oForm.TypeCount;
                                       // oLinkExpdoc.LinkedObjectType = "22";
                                    }
                                    else if (cmbcal == "PI")
                                    { 
                                        setChooseFromListField(oForm, CFL_PI, "CFL_PI", lc5bdn, "lc5bden", matLCEX, pVal.Row);
                                        SBOMain.LineNo = pVal.Row;
                                        SBOMain.TFromUID = oForm.UniqueID;
                                        SBOMain.FromCnt = oForm.TypeCount;
                                        //oLinkExpdoc.LinkedObjectType = "18";
                                    } 
                                }
                                // string docdate = oForm.Items.Item("lcod").Specific.value;
                            }
                            if (pVal.ItemUID == "cSer" && pVal.FormMode == 3)
                            {
                                oForm.Items.Item("tDocNum").Specific.Value = oForm.BusinessObject.GetNextSerialNumber(oForm.Items.Item("cSer").Specific.Value, "EXLR");
                            }
                            if (pVal.ItemUID == "cmbCRCY" && pVal.FormMode == 3)
                            {
                                SAPbouiCOM.ComboBox cb4 = oForm.Items.Item("cmbCRCY").Specific;
                                string currency = cb4.Selected.Value.ToString();
                                string docdate = oForm.Items.Item("lcod").Specific.value;

                                if (currency != "##" && currency != "INR")
                                { 
                                 if (!string.IsNullOrEmpty(docdate)) 
                                      {  
                                        //string docdate = oForm.Items.Item("lcod").Specific.value;
                                        string CurYear = docdate.Substring(0, 4);
                                        string CurMonth = docdate.Substring(4, 2);
                                        string CurDate = docdate.Substring(6, 2);
                                        string FromDateConvert = docdate.Substring(0, 4) + "-" + docdate.Substring(4, 2) + "-" + docdate.Substring(6, 2);
                                        DateTime daten = new DateTime(2020, Convert.ToInt16(CurMonth), 1);
                                        //rowmonth = daten.ToString("MMMM");
                                        rowmonth = Convert.ToInt16(CurMonth);
                                        rownum = Convert.ToInt16(CurDate);
                                        rowyear = CurYear; // Convert.ToInt16(CurYear); 

                                        string getQuery = @"SELECT Rate FROM ORTT WHERE Currency =  '" + currency + "' and RateDate = '" + FromDateConvert + "'";
                                        SAPbobsCOM.Recordset rec;
                                        rec = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                        rec.DoQuery(getQuery);
                                        if (rec.RecordCount > 0)
                                        {
                                            while (!rec.EoF)
                                            {
                                                rate = Convert.ToDouble(rec.Fields.Item("Rate").Value);
                                                rec.MoveNext();
                                            }
                                            if (rate == 0)
                                            {
                                                BubbleEvent = false;
                                                openExchangeRateForm(currency, rownum, rowmonth, rowyear); 
                                            }
                                            oForm.Items.Item("lc1exrt").Specific.value = rate.ToString();
                                        }
                                        else
                                        {
                                            if (currency != "INR")
                                            {
                                                BubbleEvent = false;
                                                openExchangeRateForm(currency, rownum, rowmonth, rowyear);

                                            }
                                            else
                                            {
                                                rate = 1.0;
                                            }
                                            oForm.Items.Item("lc1exrt").Specific.value = rate.ToString();
                                        }
                                     } 
                                    else
                                    {
                                        SBOMain.SBO_Application.StatusBar.SetText("Please select Opening date", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning); 
                                    }
                                }
                            }
                        }

                        break;
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
                        if (pVal.BeforeAction == true)
                        {

                        }
                        if (pVal.BeforeAction == false)
                        {
                            SAPbouiCOM.Matrix omatLCLD = (SAPbouiCOM.Matrix)oForm.Items.Item("matLCLD").Specific;
                            objCU.doAutoColSum(omatLCLD, "lc2dot");
                            objCU.doAutoColSum(omatLCLD, "lc2am");
                            objCU.doAutoColSum(omatLCLD, "lc2dq");
                            objCU.doAutoColSum(omatLCLD, "lc2aq");
                        }
                        break;
                    case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
                        if (pVal.BeforeAction == true)
                        {

                        }
                        if (pVal.BeforeAction == false)
                        {
                            if (pVal.CharPressed == 9)
                            {  

                                if (pVal.ItemUID == "matLCLD" && pVal.ColUID == "lc2am")
                                {
                                    double totalAmt = 0;
                                    double lcamt = Convert.ToDouble(oForm.Items.Item("lc1amt").Specific.Value);
                                    double lcexrt = Convert.ToDouble(oForm.Items.Item("lc1exrt").Specific.Value);

                                    SAPbouiCOM.Matrix omatLCLD = (SAPbouiCOM.Matrix)oForm.Items.Item("matLCLD").Specific;
                                    for (int i = 1; i <= omatLCLD.RowCount; i++)
                                    {
                                        totalAmt = totalAmt + Convert.ToDouble(((SAPbouiCOM.EditText)omatLCLD.Columns.Item("lc2am").Cells.Item(i).Specific).Value);
                                    }    
                                    oForm.Items.Item("lc1aufc").Specific.Value = totalAmt.ToString();
                                    oForm.Items.Item("lc1arfc").Specific.Value = (lcamt - totalAmt).ToString(); 
                                    oForm.Items.Item("lc1aulc").Specific.Value = (totalAmt * lcexrt).ToString();
                                    oForm.Items.Item("lc1arlc").Specific.Value = ((lcamt - totalAmt) * lcexrt ).ToString();

                                    objCU.doAutoColSum(omatLCLD, "lc2dot");
                                    objCU.doAutoColSum(omatLCLD, "lc2am");
                                    objCU.doAutoColSum(omatLCLD, "lc2dq");
                                    objCU.doAutoColSum(omatLCLD, "lc2aq");
                                }

                                if (pVal.ItemUID == "matLCDOC" && pVal.ColUID == "lc4doc")
                                {
                                    SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("matLCDOC").Specific;
                                    AddMatrixRow(oMatrix, "lc4doc");
                                }
                                if (pVal.ItemUID == "matLCDOC" && pVal.ColUID == "lc4cop")
                                {
                                    SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("matLCDOC").Specific;
                                    int a = Convert.ToInt16(oMatrix.Columns.Item("lc4org").Cells.Item(pVal.Row).Specific.Value);
                                    int b = Convert.ToInt16(oMatrix.Columns.Item("lc4cop").Cells.Item(pVal.Row).Specific.Value);

                                    oMatrix.Columns.Item("lc4tnoc").Cells.Item(pVal.Row).Specific.Value = (a + b).ToString();
                                }
                                if (pVal.ItemUID == "matLCLD" && pVal.ColUID == "lc2dq")
                                {
                                    SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("matLCLD").Specific;
                                    AddMatrixRow(oMatrix, "lc2dsn");
                                }
                                else if (pVal.ItemUID == "matLCDOC" && pVal.ColUID == "lc4doc")
                                {
                                    SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("matLCDOC").Specific;
                                    AddMatrixRow(oMatrix, "lc4doc");
                                }
                                else if (pVal.ItemUID == "matLCEX" && pVal.ColUID == "lc5expt")
                                {
                                    SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("matLCEX").Specific;
                                    AddMatrixRow(oMatrix, "lc5yc");
                                }
                                else if (pVal.ItemUID == "matLCEX" && pVal.ColUID == "lc5expt")
                                {
                                    /*SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("matLCEX").Specific;
                                    string expcode = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("lc5expt").Cells.Item(pVal.Row).Specific).Value;

                                    if (string.IsNullOrEmpty(expcode))
                                    {
                                        AddMatrixRow(oMatrix, "lc5expt");
                                        bool plFormOpen = false;
                                        for (int i = 1; i < SBOMain.SBO_Application.Forms.Count; i++)
                                        {
                                            if (SBOMain.SBO_Application.Forms.Item(i).UniqueID == "frmExpList")
                                            {
                                                SBOMain.SBO_Application.Forms.Item(i).Select();
                                                plFormOpen = true;
                                            }
                                        }
                                        if (!plFormOpen)
                                        {
                                            Program.ExExpData.EXExpMat = pVal.ItemUID;
                                            Program.ExExpData.EXExpMatRow = pVal.Row;
                                            Program.ExExpData.EXExpMatCol = "lc5expt";
                                            SBOMain.LoadFromXML("frmExpList", "EXIM");
                                            SBOMain.SBO_Application.Forms.Item("frmExpList").Select();

                                            var oForm1 = SBOMain.SBO_Application.Forms.ActiveForm;
                                            oForm1.DataSources.DataTables.Add("tab");
                                            SAPbouiCOM.Grid objGrid = oForm1.Items.Item("gridEXP").Specific;

                                            string Qry = "SELECT U_expcode as 'Expense Code',U_expname as 'Expense Name', U_expadname  as 'Expense Additional Name'FROM dbo.[@EXEM] where U_status = 1";
                                            oForm1.DataSources.DataTables.Item("tab").ExecuteQuery(Qry);
                                            objGrid.DataTable = oForm1.DataSources.DataTables.Item("tab");
                                        }
                                    }*/
                                }
                                else if (pVal.ItemUID == "matLCAMED" && pVal.ColUID == "lc6amdn")
                                {
                                    SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("matLCAMED").Specific;
                                    AddMatrixRow(oMatrix, "lc6amdn");
                                }
                                else if (pVal.ItemUID == "matLCATT" && pVal.ColUID == "lc7tp")
                                {
                                    SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("matLCATT").Specific;
                                    AddMatrixRow(oMatrix, "lc7tp");
                                }
                                else if (pVal.ItemUID == "lcsd" || pVal.ItemUID == "lcpedt")
                                {
                                    if ((string.IsNullOrEmpty(oForm.Items.Item("lcsd").Specific.value) == false) && (string.IsNullOrEmpty(oForm.Items.Item("lcpedt").Specific.value) == false))
                                    {
                                        SAPbouiCOM.EditText oDocDate1 = oForm.Items.Item("lcsd").Specific;
                                        DateTime lcsd = Convert.ToDateTime(oDocDate1.String);

                                        SAPbouiCOM.EditText oDocDate2 = oForm.Items.Item("lcpedt").Specific;
                                        DateTime lcpedt = Convert.ToDateTime(oDocDate2.String);

                                        TimeSpan age = lcpedt.Subtract(lcsd);
                                        Int32 diff = Convert.ToInt32(age.TotalDays);

                                        if (diff <= 0)
                                        {
                                            BubbleEvent = false;
                                            SBOMain.SBO_Application.StatusBar.SetText("Please Add Presention Expiry Date Greater than shipment Date", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                            oForm.Items.Item("lcpedt").Click();
                                        }
                                        else
                                        {
                                            oForm.Items.Item("lcnod").Specific.value = diff.ToString();
                                        }
                                    }
                                }
                            }
                        }
                        break;
                    case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
                        //SAPbouiCOM.Form oForms = SBOMain.SBO_Application.Forms.ActiveForm;
                        if (pVal.BeforeAction == true)
                        {
                            if (pVal.ItemUID == "matLCLD" && pVal.ColUID == "lc2dde")
                            {
                                SAPbouiCOM.ComboBox cb4 = oForm.Items.Item("cmbCRCY").Specific;
                                string currency = cb4.Selected.Value.ToString();
                                if(currency != "##")
                                {
                                    CFLConditionSO("CFL_SO", pVal.ItemUID, currency);
                                }
                                else
                                {
                                    SBOMain.SBO_Application.StatusBar.SetText("Please set currency in LC TERMS Tab", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                                } 
                            }


                            else if (pVal.ItemUID == "lc3pol")
                            {
                                CFLCondition("CFL_EXPM", pVal.ItemUID);
                            }
                            else if (pVal.ItemUID == "lc3pod")
                            {
                                CFLCondition("CFL_13", pVal.ItemUID);
                            }
                            else if (pVal.ItemUID == "lc3fd")
                            {
                                CFLCondition("CFL_14", pVal.ItemUID);
                            }

                            else if (pVal.ItemUID == "matLCEX" && pVal.ColUID == "lc5expt")
                            {
                                CFLConditionEXPType("CFL_15", pVal.ItemUID);
                            }
                            else if (pVal.ItemUID == "matLCDOC" && pVal.ColUID == "lc4doc")
                            {
                                CFLCondition("CFL_18", pVal.ItemUID);
                            } 

                            else if (pVal.ItemUID == "lcfc")
                            {
                                CFLCondition("CFL_OCRDC", pVal.ItemUID);
                            }
                            else if (pVal.ItemUID == "lc3slc")
                            {
                                CFLCondition("CFL_17", pVal.ItemUID);
                            }
                            /*else if (pVal.ItemUID == "lc3slc")
                            {
                                CFLCondition("CFL_OCRY1", pVal.ItemUID);
                            }*/
                            else if (pVal.ItemUID == "lcfv")
                            {
                                CFLCondition("CFL_OCRD", pVal.ItemUID);
                            }
                            else if (pVal.ItemUID == "lc1Abc")
                            {
                                NameValueCollection list1 = new NameValueCollection() { { "lc1Abc", "CFL_B1" }, { "lc1bbc", "CFL_B2" }, { "lc1ibc", "CFL_B3" }, { "lc1bc", "CFL_B4" } };
                                CFLCondition(list1[pVal.ItemUID], pVal.ItemUID);
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
                                    if (pVal.ItemUID == "matLCDOC" && pVal.ColUID == "lc4doc")
                                    {
                                        SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("matLCDOC").Specific;
                                        ((SAPbouiCOM.EditText)oMatrix.Columns.Item("lc4doc").Cells.Item(pVal.Row).Specific).Value = oDataTable.GetValue("U_doccode", 0).ToString();
                                        ((SAPbouiCOM.EditText)oMatrix.Columns.Item("lc4docnm").Cells.Item(pVal.Row).Specific).Value = oDataTable.GetValue("U_docname", 0).ToString();
                                        AddMatrixRow(oMatrix, "lc4doc");

                                    }
                                    else if (pVal.ItemUID == "lc3oc")
                                    {
                                        oForm.Items.Item("lc3oc").Specific.value = oDataTable.GetValue("Name", 0).ToString();
                                    }
                                    else if (pVal.ItemUID == "lc3slc")
                                    {
                                        oForm.Items.Item("lc3slc").Specific.value = oDataTable.GetValue("CardCode", 0).ToString();
                                        oForm.Items.Item("lc3sln").Specific.value = oDataTable.GetValue("CardName", 0).ToString();
                                    }
                                    else if (pVal.ItemUID == "lc3fdc")
                                    {
                                        oForm.Items.Item("lc3fdc").Specific.value = oDataTable.GetValue("Name", 0).ToString();
                                    }
                                    else if (pVal.ItemUID == "matLCEX" && pVal.ColUID == "lc5bden")
                                    {
                                        SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("matLCEX").Specific;
                                        try
                                        {
                                            SAPbouiCOM.ComboBox cb4 = oMatrix.Columns.Item("lc5yc").Cells.Item(pVal.Row).Specific; // oForm.Items.Item("cmbCRCY").Specific;
                                            string cmbcal = cb4.Selected.Value.ToString();
                                            double lineTotal = 0;
                                            double FClineTotal = 0;
                                            double rate = 0;
                                            string Currency = null;

                                            if (cmbcal == "PR")
                                            {
                                                lineTotal = objCU.getLineTotalFromDocKey("PRQ1", oDataTable.GetValue("DocEntry", 0).ToString());
                                                FClineTotal = objCU.getFCLineTotalFromDocKey("PRQ1", oDataTable.GetValue("DocEntry", 0).ToString());
                                                Currency = objCU.getCurrFromDocKey("PRQ1", oDataTable.GetValue("DocEntry", 0).ToString());
                                                rate = objCU.getRateFromDocKey("PRQ1", oDataTable.GetValue("DocEntry", 0).ToString());
                                            }
                                            else if (cmbcal == "PQ")
                                            {
                                                lineTotal = objCU.getLineTotalFromDocKey("PQT1", oDataTable.GetValue("DocEntry", 0).ToString());
                                                FClineTotal = objCU.getFCLineTotalFromDocKey("PQT1", oDataTable.GetValue("DocEntry", 0).ToString());
                                                Currency = objCU.getCurrFromDocKey("PQT1", oDataTable.GetValue("DocEntry", 0).ToString());
                                                rate = objCU.getRateFromDocKey("PQT1", oDataTable.GetValue("DocEntry", 0).ToString());
                                            }
                                            else if (cmbcal == "PO")
                                            {
                                                lineTotal = objCU.getLineTotalFromDocKey("POR1", oDataTable.GetValue("DocEntry", 0).ToString());
                                                FClineTotal = objCU.getFCLineTotalFromDocKey("POR1", oDataTable.GetValue("DocEntry", 0).ToString());
                                                Currency = objCU.getCurrFromDocKey("POR1", oDataTable.GetValue("DocEntry", 0).ToString());
                                                rate = objCU.getRateFromDocKey("POR1", oDataTable.GetValue("DocEntry", 0).ToString());

                                            }
                                            else if (cmbcal == "PI")
                                            {
                                                lineTotal = objCU.getLineTotalFromDocKey("PCH1", oDataTable.GetValue("DocEntry", 0).ToString());
                                                FClineTotal = objCU.getFCLineTotalFromDocKey("PCH1", oDataTable.GetValue("DocEntry", 0).ToString());
                                                Currency = objCU.getCurrFromDocKey("PCH1", oDataTable.GetValue("DocEntry", 0).ToString());
                                                rate = objCU.getRateFromDocKey("PCH1", oDataTable.GetValue("DocEntry", 0).ToString());
                                            }

                                            ((SAPbouiCOM.EditText)oMatrix.Columns.Item("lc5bdn").Cells.Item(pVal.Row).Specific).Value = oDataTable.GetValue("DocNum", 0).ToString();
                                            ((SAPbouiCOM.EditText)oMatrix.Columns.Item("lc5bden").Cells.Item(pVal.Row).Specific).Value = oDataTable.GetValue("DocEntry", 0).ToString();
                                            ((SAPbouiCOM.EditText)oMatrix.Columns.Item("lc5pl").Cells.Item(pVal.Row).Specific).Value = lineTotal.ToString();
                                            ((SAPbouiCOM.EditText)oMatrix.Columns.Item("lc5pf").Cells.Item(pVal.Row).Specific).Value = FClineTotal.ToString();
                                            ((SAPbouiCOM.EditText)oMatrix.Columns.Item("lc5cur").Cells.Item(pVal.Row).Specific).Value = Currency.ToString();
                                            ((SAPbouiCOM.EditText)oMatrix.Columns.Item("lc5rt").Cells.Item(pVal.Row).Specific).Value = rate.ToString();

                                        }
                                        catch (Exception ex)
                                        {                                            
                                        }
                                    }
                                    else if (pVal.ItemUID == "matLCLD" && pVal.ColUID == "lc2dde")
                                    {
                                        SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("matLCLD").Specific;
                                        ((SAPbouiCOM.EditText)oMatrix.Columns.Item("lc2dde").Cells.Item(pVal.Row).Specific).Value = oDataTable.GetValue("DocEntry", 0).ToString();
                                        ((SAPbouiCOM.EditText)oMatrix.Columns.Item("lc2dsn").Cells.Item(pVal.Row).Specific).Value = oDataTable.GetValue("DocNum", 0).ToString();
                                        DateTime dt = Convert.ToDateTime(oDataTable.GetValue("DocDate", 0).ToString());
                                        string abc = dt.ToString("yyyyMMdd");
                                        ((SAPbouiCOM.EditText)oMatrix.Columns.Item("lc2dt").Cells.Item(pVal.Row).Specific).Value =  dt.ToString("yyyyMMdd") ; //oDataTable.GetValue("DocDate", 0).ToString();
                                        ((SAPbouiCOM.EditText)oMatrix.Columns.Item("lc2bc").Cells.Item(pVal.Row).Specific).Value = oDataTable.GetValue("CardCode", 0).ToString();
                                        ((SAPbouiCOM.EditText)oMatrix.Columns.Item("lc2bn").Cells.Item(pVal.Row).Specific).Value = oDataTable.GetValue("CardName", 0).ToString();
                                        ((SAPbouiCOM.EditText)oMatrix.Columns.Item("lc2brn").Cells.Item(pVal.Row).Specific).Value = oDataTable.GetValue("NumAtCard", 0).ToString();
                                        ((SAPbouiCOM.EditText)oMatrix.Columns.Item("lc2cr").Cells.Item(pVal.Row).Specific).Value = oDataTable.GetValue("DocCur", 0).ToString();

                                        string getDocEntry = null;
                                        if (oDataTable.GetValue("DocCur", 0).ToString() != "INR")
                                        {
                                              getDocEntry = " select sum(Quantity) as 'Qty', sum(TotalFrgn) as 'linetotal'  from RDR1 where DocEntry = '" + oDataTable.GetValue("DocEntry", 0).ToString() + "'";
                                         }
                                        else { 
                                             getDocEntry = " select sum(Quantity) as 'Qty', sum(LineTotal) as 'linetotal'  from RDR1 where DocEntry = '" + oDataTable.GetValue("DocEntry", 0).ToString() + "'";
                                        }
                                        SAPbobsCOM.Recordset rec = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                        rec.DoQuery(getDocEntry);
                                        if (rec.RecordCount > 0)
                                        {
                                            ((SAPbouiCOM.EditText)oMatrix.Columns.Item("lc2dot").Cells.Item(pVal.Row).Specific).Value = Convert.ToString(rec.Fields.Item("linetotal").Value);
                                            ((SAPbouiCOM.EditText)oMatrix.Columns.Item("lc2dq").Cells.Item(pVal.Row).Specific).Value = Convert.ToString(rec.Fields.Item("Qty").Value); 
                                        }

                                        objCU.doAutoColSum(oMatrix, "lc2dot");
                                        objCU.doAutoColSum(oMatrix, "lc2am"); 
                                        objCU.doAutoColSum(oMatrix, "lc2dq");
                                        objCU.doAutoColSum(oMatrix, "lc2aq");
                                    }

                                    else if (pVal.ItemUID == "lcfc")
                                    {
                                        oForm.Items.Item("lcfc").Specific.value = oDataTable.GetValue("CardCode", 0).ToString();
                                    }
                                    else if (pVal.ItemUID == "lc3pol")
                                    {
                                        oForm.Items.Item("lc3pol").Specific.value = oDataTable.GetValue("U_portcode", 0).ToString();
                                    }
                                    else if (pVal.ItemUID == "lc3pod")
                                    {
                                        oForm.Items.Item("lc3pod").Specific.value = oDataTable.GetValue("U_portcode", 0).ToString();
                                    }
                                    else if (pVal.ItemUID == "lc3fd")
                                    {
                                        oForm.Items.Item("lc3fd").Specific.value = oDataTable.GetValue("U_portcode", 0).ToString();
                                    }
                                    else if (pVal.ItemUID == "matLCEX" && pVal.ColUID == "lc5expt")
                                    {
                                        SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("matLCEX").Specific;
                                        ((SAPbouiCOM.EditText)oMatrix.Columns.Item("lc5expt").Cells.Item(pVal.Row).Specific).Value = oDataTable.GetValue("U_expcode", 0).ToString();
                                        ((SAPbouiCOM.EditText)oMatrix.Columns.Item("lc5expnm").Cells.Item(pVal.Row).Specific).Value = oDataTable.GetValue("U_expname", 0).ToString();
                                        AddMatrixRow(oMatrix, "lc5expt");

                                    }
                                    else if (pVal.ItemUID == "lcfv")
                                    {     
                                      oForm.Items.Item("lcfv").Specific.value = oDataTable.GetValue("CardCode", 0).ToString();
                                    }
                                    else if (pVal.ItemUID == "lc1Abc")
                                    {
                                        oForm.Items.Item("lc1Abc").Specific.value = oDataTable.GetValue("BankCode", 0).ToString();
                                        oForm.Items.Item("lc1Abn").Specific.value = oDataTable.GetValue("BankName", 0).ToString();
                                        oForm.Items.Item("lc1Aban").Specific.value = oDataTable.GetValue("DfltAcct", 0).ToString();
                                        oForm.Items.Item("lc1Abb").Specific.value = oDataTable.GetValue("DfltBranch", 0).ToString();
                                        oForm.Items.Item("lc1Aswc").Specific.value = oDataTable.GetValue("SwiftNum", 0).ToString();
                                        oForm.Items.Item("lc1Aba").Specific.value = objCU.BankAddress(oDataTable.GetValue("BankCode", 0).ToString());   
                                    }
                                    else if (pVal.ItemUID == "lc1bbc")
                                    {
                                        oForm.Items.Item("lc1bbc").Specific.value = oDataTable.GetValue("BankCode", 0).ToString();
                                        oForm.Items.Item("lc1bbn").Specific.value = oDataTable.GetValue("BankName", 0).ToString();
                                        oForm.Items.Item("lc1bban").Specific.value = oDataTable.GetValue("DfltAcct", 0).ToString();
                                        oForm.Items.Item("lc1bbb").Specific.value = oDataTable.GetValue("DfltBranch", 0).ToString();
                                        oForm.Items.Item("lc1bswc").Specific.value = oDataTable.GetValue("SwiftNum", 0).ToString();
                                        oForm.Items.Item("lc1bba").Specific.value = objCU.BankAddress(oDataTable.GetValue("BankCode", 0).ToString());
                                    }
                                    else if (pVal.ItemUID == "lc1ibc")
                                    {
                                        oForm.Items.Item("lc1ibc").Specific.value = oDataTable.GetValue("BankCode", 0).ToString();
                                        oForm.Items.Item("lc1ibn").Specific.value = oDataTable.GetValue("BankName", 0).ToString();
                                        oForm.Items.Item("lc1iban").Specific.value = oDataTable.GetValue("DfltAcct", 0).ToString();
                                        oForm.Items.Item("lc1ibb").Specific.value = oDataTable.GetValue("DfltBranch", 0).ToString();
                                        oForm.Items.Item("lc1iswc").Specific.value = oDataTable.GetValue("SwiftNum", 0).ToString();
                                        oForm.Items.Item("lc1iba").Specific.value = objCU.BankAddress(oDataTable.GetValue("BankCode", 0).ToString());
                                    }
                                    else if (pVal.ItemUID == "lc1bc")
                                    {
                                        oForm.Items.Item("lc1bc").Specific.value = oDataTable.GetValue("BankCode", 0).ToString();
                                        oForm.Items.Item("lc1bn").Specific.value = oDataTable.GetValue("BankName", 0).ToString();
                                        oForm.Items.Item("lc1ban").Specific.value = oDataTable.GetValue("DfltAcct", 0).ToString();
                                     
                                    }
                                } catch (Exception ex)
                                {

                                }
                            }
                        }
                        break;
                    case BoEventTypes.et_FORM_CLOSE:
                        if (pVal.BeforeAction == true)
                        {

                        }
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
                            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("matLCLD").Specific;
                            //DeleteMatrixBlankRow(oMatrix);
                            if (pVal.ItemUID == "1")
                            {
                               //z oForm.Items.Item("tattach").Specific.value = SBOMain.Get_Attach_Folder_Path() + Path.GetFileName(BrowseFilePath);
                            }
                        }
                        if (pVal.BeforeAction == false)
                        {

                            if (pVal.ItemUID == "btnDIS")
                            {                               
                                SAPbouiCOM.Matrix omatLCLD = (SAPbouiCOM.Matrix)oForm.Items.Item("matLCATT").Specific;
                                int rowid = omatLCLD.GetNextSelectedRow(); // getnextselectedrow(0, roworder)
                               
                                string a =  omatLCLD.Columns.Item("lc7tp").Cells.Item(rowid).Specific.Value;
                                string b = omatLCLD.Columns.Item("lc7fn").Cells.Item(rowid).Specific.Value;
                                 
                                string fullpath = a + b;
                                System.Diagnostics.Process.Start(@fullpath);
                            }

                            if (pVal.ItemUID == "lbPOL" || pVal.ItemUID == "lbtPOD" || pVal.ItemUID == "lbtPOR")
                            {
                                string abc = null;
                                if (pVal.ItemUID == "lbPOL")
                                {
                                    abc = oForm.Items.Item("lc3pol").Specific.Value;
                                }
                                else if (pVal.ItemUID == "lbtPOD")
                                {
                                    abc = oForm.Items.Item("lc3pod").Specific.Value;
                                }
                                else if (pVal.ItemUID == "lbtPOR")
                                {
                                    abc = oForm.Items.Item("lc3fd").Specific.Value;
                                }

                                objCU.FormLoadAndActivate("frmPortMaster", "mnsmEXIM007"); 
                                OutwardToPortMaster inPortMaster = new OutwardToPortMaster();
                                inPortMaster.portcode = abc;
                                clsPortMaster oPort = new clsPortMaster(inPortMaster);
                                //oForm.Close();
                            }

                            if (pVal.ItemUID == "1" && (oForm.Mode == BoFormMode.fm_ADD_MODE || oForm.Mode == BoFormMode.fm_UPDATE_MODE))
                            {
                                /*SAPbouiCOM.Matrix matrix = oForm.Items.Item("matLCATT").Specific;

                                for (int i = 1; i <= matrix.RowCount; i++)
                                {
                                   *//* string TargetPath = (matrix.Columns.Item("lc7tp").Cells.Item(i).Specific).Value;
                                    string FileName = (matrix.Columns.Item("lc7fn").Cells.Item(i).Specific).Value;
                                    string ReplaceFilePath = SBOMain.Get_Attach_Folder_Path() + FileName; 
                                    File.Move(BrowseFilePath, ReplaceFilePath); *//*
                                }*/
                            }

                            if (pVal.ItemUID == "btnREFEXP")
                            {
                                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("matLCEX").Specific;
                                string ExpForm = null;
                                string FormDocEntry = null;
                                double PLCAmt = 0;
                                double PFCAmt = 0;
                                double LCAmt = 0;
                                double FCAmt = 0;

                                for (int i = 1; i <= oMatrix.RowCount; i++)
                                {
                                    ExpForm =  (oMatrix.Columns.Item("lc5yc").Cells.Item(i).Specific).Value;
                                    FormDocEntry = (oMatrix.Columns.Item("lc5bden").Cells.Item(i).Specific).Value;

                                    PLCAmt = Convert.ToDouble((oMatrix.Columns.Item("lc5pf").Cells.Item(i).Specific).Value);
                                    PFCAmt = Convert.ToDouble((oMatrix.Columns.Item("lc5pl").Cells.Item(i).Specific).Value);
                                    LCAmt = 0;
                                    FCAmt = 0;

                                    if (ExpForm == "PR")
                                    {
                                        LCAmt = 0;
                                        FCAmt = 0;
                                        // For PR -> PO -> All
                                        string getQuery3 = "Select DISTINCT(DocEntry) from POR1 Where BaseEntry = '" + FormDocEntry + "' AND BaseType = '1470000113'";
                                        SAPbobsCOM.Recordset rec3 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                        rec3.DoQuery(getQuery3);
                                        if (rec3.RecordCount > 0)
                                        {
                                            while (!rec3.EoF)
                                            {
                                                string getQuery4 = " SELECT sum(LineTotal) as LC,sum(TotalFrgn) As FC from PCH1 where BaseEntry = '" + rec3.Fields.Item("DocEntry").Value + "' and BaseType = '22'";
                                                SAPbobsCOM.Recordset rec4 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                                rec4.DoQuery(getQuery4);
                                                if (rec4.RecordCount > 0)
                                                {
                                                    LCAmt = LCAmt + Convert.ToDouble(rec4.Fields.Item("LC").Value);
                                                    FCAmt = FCAmt + Convert.ToDouble(rec4.Fields.Item("FC").Value);
                                                }

                                                // For Delivery then Invoice
                                                string getQuery9 = "Select DocEntry from PDN1 Where BaseEntry = '" + rec3.Fields.Item("DocEntry").Value + "' AND BaseType = '22'";
                                                SAPbobsCOM.Recordset rec9 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                                rec9.DoQuery(getQuery9);
                                                if (rec9.RecordCount > 0)
                                                {
                                                    while (!rec9.EoF)
                                                    {
                                                        string getQuery2 = " SELECT sum(LineTotal) as LC,sum(TotalFrgn) As FC from PCH1 where BaseEntry = '" + rec9.Fields.Item("DocEntry").Value + "' and BaseType = '20'";
                                                        SAPbobsCOM.Recordset rec2 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                                        rec2.DoQuery(getQuery2);
                                                        if (rec2.RecordCount > 0)
                                                        {
                                                            LCAmt = LCAmt + Convert.ToDouble(rec2.Fields.Item("LC").Value);
                                                            FCAmt = FCAmt + Convert.ToDouble(rec2.Fields.Item("FC").Value);
                                                        }
                                                        rec9.MoveNext();
                                                    }
                                                }

                                                rec3.MoveNext();
                                            }
                                        }

                                        // For PR -> PQ -> All
                                        string getQuery5 = "Select DISTINCT(DocEntry) from PQT1 Where BaseEntry = '" + FormDocEntry + "' AND BaseType = '1470000113'";
                                        SAPbobsCOM.Recordset rec5 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                        rec5.DoQuery(getQuery5);
                                        if (rec5.RecordCount > 0)
                                        {
                                            while (!rec5.EoF)
                                            {
                                                string getQuery = " SELECT sum(LineTotal) as LC,sum(TotalFrgn) As FC from PCH1 where BaseEntry = '" + FormDocEntry + "' and BaseType = '540000006'";
                                                SAPbobsCOM.Recordset rec = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                                rec.DoQuery(getQuery);
                                                if (rec.RecordCount > 0)
                                                {
                                                    LCAmt = Convert.ToDouble(rec.Fields.Item("LC").Value);
                                                    FCAmt = Convert.ToDouble(rec.Fields.Item("FC").Value);
                                                }

                                                // For PQ -> GRPO -> Invoice 
                                                getQuery = "Select DocEntry from PDN1 Where BaseEntry = '" + FormDocEntry + "' AND BaseType = '540000006'";
                                                SAPbobsCOM.Recordset rec1 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                                rec1.DoQuery(getQuery);
                                                if (rec1.RecordCount > 0)
                                                {
                                                    while (!rec1.EoF)
                                                    {
                                                        string getQuery2 = " SELECT sum(LineTotal) as LC,sum(TotalFrgn) As FC from PCH1 where BaseEntry = '" + rec1.Fields.Item("DocEntry").Value + "' and BaseType = '20'";
                                                        SAPbobsCOM.Recordset rec2 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                                        rec2.DoQuery(getQuery2);
                                                        if (rec2.RecordCount > 0)
                                                        {
                                                            LCAmt = LCAmt + Convert.ToDouble(rec2.Fields.Item("LC").Value);
                                                            FCAmt = FCAmt + Convert.ToDouble(rec2.Fields.Item("FC").Value);
                                                        }
                                                        rec1.MoveNext();
                                                    }
                                                }

                                                // For PQ -> PO -> GRPO -> Invoice 
                                                string getQuery6 = "Select DISTINCT(DocEntry) from POR1 Where BaseEntry = '" + FormDocEntry + "' AND BaseType = '540000006'";
                                                SAPbobsCOM.Recordset rec6 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                                rec6.DoQuery(getQuery6);
                                                if (rec6.RecordCount > 0)
                                                {
                                                    while (!rec6.EoF)
                                                    {
                                                        string getQuery4 = " SELECT sum(LineTotal) as LC,sum(TotalFrgn) As FC from PCH1 where BaseEntry = '" + rec6.Fields.Item("DocEntry").Value + "' and BaseType = '22'";
                                                        SAPbobsCOM.Recordset rec4 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                                        rec4.DoQuery(getQuery4);
                                                        if (rec4.RecordCount > 0)
                                                        {
                                                            LCAmt = LCAmt + Convert.ToDouble(rec4.Fields.Item("LC").Value);
                                                            FCAmt = FCAmt + Convert.ToDouble(rec4.Fields.Item("FC").Value);
                                                        }

                                                        // For Delivery then Invoice
                                                        string getQuery7 = "Select DocEntry from PDN1 Where BaseEntry = '" + rec6.Fields.Item("DocEntry").Value + "' AND BaseType = '22'";
                                                        SAPbobsCOM.Recordset rec7 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                                        rec7.DoQuery(getQuery7);
                                                        if (rec7.RecordCount > 0)
                                                        {
                                                            while (!rec7.EoF)
                                                            {

                                                                string getQuery2 = " SELECT sum(LineTotal) as LC,sum(TotalFrgn) As FC from PCH1 where BaseEntry = '" + rec7.Fields.Item("DocEntry").Value + "' and BaseType = '20'";
                                                                SAPbobsCOM.Recordset rec2 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                                                rec2.DoQuery(getQuery2);
                                                                if (rec2.RecordCount > 0)
                                                                {
                                                                    LCAmt = LCAmt + Convert.ToDouble(rec2.Fields.Item("LC").Value);
                                                                    FCAmt = FCAmt + Convert.ToDouble(rec2.Fields.Item("FC").Value);
                                                                }
                                                                rec7.MoveNext();
                                                            }
                                                        }

                                                        rec6.MoveNext();
                                                    }
                                                } 

                                                rec5.MoveNext();
                                            }
                                        }
                                    }
                                    if (ExpForm == "PQ")
                                    {
                                        LCAmt = 0;
                                        FCAmt = 0;
                                        // For Direct invoice from PQ
                                        string getQuery = " SELECT sum(LineTotal) as LC,sum(TotalFrgn) As FC from PCH1 where BaseEntry = '" + FormDocEntry + "' and BaseType = '540000006'";
                                        SAPbobsCOM.Recordset rec = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                        rec.DoQuery(getQuery);
                                        if (rec.RecordCount > 0)
                                        {
                                            LCAmt = Convert.ToDouble(rec.Fields.Item("LC").Value);
                                            FCAmt = Convert.ToDouble(rec.Fields.Item("FC").Value);
                                        }

                                        // For PQ -> GRPO -> Invoice 
                                        getQuery = "Select DocEntry from PDN1 Where BaseEntry = '" + FormDocEntry + "' AND BaseType = '540000006'";
                                        SAPbobsCOM.Recordset rec1 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                        rec1.DoQuery(getQuery);
                                        if (rec1.RecordCount > 0)
                                        {
                                            while (!rec1.EoF)
                                            {
                                                string getQuery2 = " SELECT sum(LineTotal) as LC,sum(TotalFrgn) As FC from PCH1 where BaseEntry = '" + rec1.Fields.Item("DocEntry").Value + "' and BaseType = '20'";
                                                SAPbobsCOM.Recordset rec2 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                                rec2.DoQuery(getQuery2);
                                                if (rec2.RecordCount > 0)
                                                {
                                                    LCAmt = LCAmt + Convert.ToDouble(rec2.Fields.Item("LC").Value);
                                                    FCAmt = FCAmt + Convert.ToDouble(rec2.Fields.Item("FC").Value);
                                                }
                                                rec1.MoveNext();
                                            }
                                        }

                                        // For PQ -> PO -> GRPO -> Invoice 
                                        string getQuery3 = "Select DISTINCT(DocEntry) from POR1 Where BaseEntry = '" + FormDocEntry + "' AND BaseType = '540000006'";
                                        SAPbobsCOM.Recordset rec3 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                        rec3.DoQuery(getQuery3);
                                        if (rec3.RecordCount > 0)
                                        {
                                            while (!rec3.EoF)
                                            {
                                                string getQuery4 = " SELECT sum(LineTotal) as LC,sum(TotalFrgn) As FC from PCH1 where BaseEntry = '" + rec3.Fields.Item("DocEntry").Value + "' and BaseType = '22'";
                                                SAPbobsCOM.Recordset rec4 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                                rec4.DoQuery(getQuery4);
                                                if (rec4.RecordCount > 0)
                                                {
                                                    LCAmt = LCAmt + Convert.ToDouble(rec4.Fields.Item("LC").Value);
                                                    FCAmt = FCAmt + Convert.ToDouble(rec4.Fields.Item("FC").Value);
                                                }

                                                // For Delivery then Invoice
                                                string getQuery5 = "Select DocEntry from PDN1 Where BaseEntry = '" + rec3.Fields.Item("DocEntry").Value + "' AND BaseType = '22'";
                                                SAPbobsCOM.Recordset rec5 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                                rec5.DoQuery(getQuery5);
                                                if (rec5.RecordCount > 0)
                                                {
                                                    while (!rec5.EoF)
                                                    {
                                                        
                                                        string getQuery2 = " SELECT sum(LineTotal) as LC,sum(TotalFrgn) As FC from PCH1 where BaseEntry = '" + rec5.Fields.Item("DocEntry").Value + "' and BaseType = '20'";
                                                        SAPbobsCOM.Recordset rec2 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                                        rec2.DoQuery(getQuery2);
                                                        if (rec2.RecordCount > 0)
                                                        {
                                                            LCAmt = LCAmt + Convert.ToDouble(rec2.Fields.Item("LC").Value);
                                                            FCAmt = FCAmt + Convert.ToDouble(rec2.Fields.Item("FC").Value);
                                                        }
                                                        rec5.MoveNext();
                                                    }
                                                }

                                                rec3.MoveNext();
                                            }
                                        }
                                    }
                                    if (ExpForm == "PO")
                                    { 
                                        // For Direct invoice from PO
                                        string getQuery = " SELECT sum(LineTotal) as LC,sum(TotalFrgn) As FC from PCH1 where BaseEntry = '" + FormDocEntry + "' and BaseType = '22'";
                                        SAPbobsCOM.Recordset rec = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                        rec.DoQuery(getQuery);
                                        if(rec.RecordCount > 0) { 
                                            LCAmt = Convert.ToDouble(rec.Fields.Item("LC").Value);
                                            FCAmt = Convert.ToDouble(rec.Fields.Item("FC").Value);
                                        } 

                                        // For Delivery then Invoice
                                        getQuery = "Select DocEntry from PDN1 Where BaseEntry = '" + FormDocEntry + "' AND BaseType = '22'";
                                        SAPbobsCOM.Recordset rec1 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                        rec1.DoQuery(getQuery);
                                        if (rec1.RecordCount > 0)
                                        {
                                            while (!rec1.EoF)
                                            {
                                                string getQuery2 = " SELECT sum(LineTotal) as LC,sum(TotalFrgn) As FC from PCH1 where BaseEntry = '" + rec1.Fields.Item("DocEntry").Value + "' and BaseType = '20'";
                                                SAPbobsCOM.Recordset rec2 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                                rec2.DoQuery(getQuery2);
                                                if (rec2.RecordCount > 0)
                                                {
                                                    LCAmt = LCAmt + Convert.ToDouble(rec2.Fields.Item("LC").Value);
                                                    FCAmt = FCAmt + Convert.ToDouble(rec2.Fields.Item("FC").Value);
                                                }
                                                rec1.MoveNext();
                                            }
                                           
                                        } 

                                    }
                                    if (ExpForm == "PI")
                                    { 
                                        /*
                                        double[] myNum = APINVgetFCLCFromAPINV(FormDocEntry);
                                        LCAmt = myNum[0];
                                        FCAmt = myNum[1];*/
                                         string getQuery = "SELECT sum(LineTotal) as LC,sum(TotalFrgn) As FC from PCH1 where DocEntry =  '" + FormDocEntry + "'";
                                        SAPbobsCOM.Recordset rec = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                        rec.DoQuery(getQuery);
                                        LCAmt = Convert.ToDouble(rec.Fields.Item("LC").Value);
                                        FCAmt = Convert.ToDouble(rec.Fields.Item("FC").Value);
                                    }
                                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("lc5al").Cells.Item(i).Specific).Value = LCAmt.ToString();
                                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("lc5af").Cells.Item(i).Specific).Value = FCAmt.ToString();
                                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("lc5dl").Cells.Item(i).Specific).Value = (PLCAmt - LCAmt).ToString();
                                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("lc5df").Cells.Item(i).Specific).Value = (PFCAmt - FCAmt).ToString(); 
                                }
                            }

                            if (pVal.ItemUID == "tab2")
                            {
                                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("matLCLD").Specific;
                                AddMatrixRow(oMatrix, "lc2dsn");
                            }
                            if (pVal.ItemUID == "tab4")
                            {
                                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("matLCDOC").Specific;
                                AddMatrixRow(oMatrix, "lc4doc");
                            }
                            if (pVal.ItemUID == "tab5")
                            {
                                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("matLCEX").Specific;
                                AddMatrixRow(oMatrix, "lc5expt");
                            }
                            if (pVal.ItemUID == "tab6")
                            {
                                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("matLCAMED").Specific;
                                AddMatrixRow(oMatrix, "lc6amdn");
                            }
                            if (pVal.ItemUID == "tab7")
                            {
                                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("matLCATT").Specific;
                                //AddMatrixRow(oMatrix, "lc7tp");
                            } 
                        }

                        break;
                    case BoEventTypes.et_ITEM_PRESSED:
                        oForm = SBOMain.SBO_Application.Forms.ActiveForm;
                        if (pVal.BeforeAction == true)
                        {
                            if (pVal.ItemUID == "1" && (oForm.Mode == BoFormMode.fm_ADD_MODE || oForm.Mode == BoFormMode.fm_UPDATE_MODE))
                            {
                                if (string.IsNullOrEmpty(oForm.Items.Item("lcsd").Specific.value))
                                {
                                    BubbleEvent = false;
                                    SBOMain.SBO_Application.StatusBar.SetText("Please Add Shipment Date", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                    oForm.Items.Item("lcsd").Click();
                                }
                                else if (string.IsNullOrEmpty(oForm.Items.Item("lcpedt").Specific.value))
                                {
                                    BubbleEvent = false;
                                    SBOMain.SBO_Application.StatusBar.SetText("Please Add Presention Expiry Date", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                    oForm.Items.Item("lcpedt").Click();
                                }
                                else if (string.IsNullOrEmpty(oForm.Items.Item("lcsd").Specific.value) == false)
                                {
                                    SAPbouiCOM.EditText oDocDate1 = oForm.Items.Item("lcsd").Specific;
                                    DateTime lcsd = Convert.ToDateTime(oDocDate1.String);

                                    SAPbouiCOM.EditText oDocDate2 = oForm.Items.Item("lcpedt").Specific;
                                    DateTime lcpedt = Convert.ToDateTime(oDocDate2.String);

                                    TimeSpan age = lcpedt.Subtract(lcsd);
                                    Int32 diff = Convert.ToInt32(age.TotalDays);

                                    if (diff <= 0)
                                    {
                                        BubbleEvent = false;
                                        SBOMain.SBO_Application.StatusBar.SetText("Please Add Presention Expiry Date Greater than shipment Date", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                        oForm.Items.Item("lcpedt").Click();
                                    }
                                }
                            }
                        }
                        if (pVal.BeforeAction == false)
                        {  
                            if (pVal.ItemUID == "1" && (oForm.Mode == BoFormMode.fm_ADD_MODE || oForm.Mode == BoFormMode.fm_UPDATE_MODE))
                            {
                               SAPbouiCOM.Matrix matrix = oForm.Items.Item("matLCATT").Specific; 
                                for (int i = 0; i < BrowseFilePath.Count; i++)
                                {
                                    string a = Convert.ToString(BrowseFilePath[i]);
                                    string b = Convert.ToString(ReplaceFilePath[i]);
                                    File.Move(a,b);
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


            }
            finally
            {
              /*  if (oForm != null)
                    oForm.Freeze(false);*/
            } 
            return BubbleEvent;
        }
         
        private void setChooseFromListField(SAPbouiCOM.Form oForm, SAPbouiCOM.ChooseFromList CFL, string clfname, SAPbouiCOM.EditText editext, string fieldname, SAPbouiCOM.Matrix matLCEX, int rowno)
        {
            CFL = oForm.ChooseFromLists.Item(clfname);
            SAPbouiCOM.Column oCol = matLCEX.Columns.Item("lc5bden");
            oCol.ChooseFromListUID = clfname;
            oCol.ChooseFromListAlias = "DocEntry";
            
        }

        public void openExchangeRateForm(string currency, int rownum, int rowmonth, string rowyear)
        {
            SBOMain.SBO_Application.StatusBar.SetText("Please Add Exchange Rate for customer's Currency.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            SBOMain.SBO_Application.Menus.Item("3333").Activate();
            SAPbouiCOM.Form oForms = SBOMain.SBO_Application.Forms.ActiveForm;
            SAPbouiCOM.Matrix exratematrix = (SAPbouiCOM.Matrix)oForms.Items.Item("4").Specific;

            SAPbouiCOM.ComboBox cb1 = (SAPbouiCOM.ComboBox)oForms.Items.Item("13").Specific;
            cb1.Select(rowmonth.ToString());

            SAPbouiCOM.ComboBox cb2 = (SAPbouiCOM.ComboBox)oForms.Items.Item("12").Specific;
            cb2.Select(rowyear);

            int matcol = exratematrix.Columns.Count;
            string coltitle = null;
            string colname = null;
            for (int i = 0; i < matcol; i++)
            {
                colname = "V_" + i.ToString();
                coltitle = exratematrix.Columns.Item(colname).Title.ToString();
                if (coltitle == currency)
                {
                    i = matcol;
                    exratematrix.Columns.Item(colname).Cells.Item(rownum).Click();
                }
            }
        }
        private void CFLConditionEXPType(string CFLID, string ItemUID)
        {
            oCFL = oForm.ChooseFromLists.Item(CFLID);
            oConds = SBOMain.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions);

            if (CFLID == "CFL_15")
            {
                oCond = oConds.Add();
                oCond.Alias = "U_status";
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCond.CondVal = "1";
                oCFL.SetConditions(oConds);
            }
            oCFL = null;
            oCond = null;
            oConds = null;

        }

        public void Form_Load_Components(SAPbouiCOM.Form oForm, string mode)
        {  
            //oForm.Freeze(true);

            oForm.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized;
            SAPbouiCOM.EditText oEdit;
            string Table = "@EXLR";
            DateTime now = DateTime.Now;
            if (mode != "OK")
            {
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                oEdit = oForm.Items.Item("tCode").Specific;
                objCU.GetNextDocNum(ref oEdit, ref Table);
                oForm.Items.Item("tDocNum").Specific.value = oForm.BusinessObject.GetNextSerialNumber("DocEntry", "EXLR");
                Events.Series.SeriesCombo("EXLR", "cSer");
                oForm.Items.Item("cSer").DisplayDesc = true;

                oForm.Items.Item("tab1").Visible = true;
                oForm.Items.Item("tab1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                oForm.PaneLevel = 1;

                oForm.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized;

                //oForm.Items.Item("exdd").Specific.value = DateTime.Now.ToString("yyyyMMdd");

                SAPbouiCOM.ComboBox cb = (SAPbouiCOM.ComboBox)oForm.Items.Item("lcsta").Specific;
                cb.ExpandType = BoExpandType.et_DescriptionOnly;
                cb.Select("O");

                /******/
                SAPbouiCOM.ComboBox cb4 = (SAPbouiCOM.ComboBox)oForm.Items.Item("cmbCRCY").Specific;

                string getQuery = @"Select CurrCode,CurrName From OCRN";
                string QueryItemCode = string.Empty;

                SAPbobsCOM.Recordset rec;
                rec = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                rec.DoQuery(getQuery);

                if (rec.RecordCount > 0)
                {
                    while (!rec.EoF)
                    {
                        cb4.ValidValues.Add(Convert.ToString(rec.Fields.Item("CurrCode").Value), Convert.ToString(rec.Fields.Item("CurrName").Value));
                        rec.MoveNext();
                    }
                }
                cb4.ValidValues.Add("##", "All Currencies");
                cb4.ExpandType = BoExpandType.et_DescriptionOnly;
                cb4.Select("INR");
                /******/ 

                /**SHIPMENT DETAILS Radio Buttons **/
                SAPbouiCOM.OptionBtn tallowed = (SAPbouiCOM.OptionBtn)oForm.Items.Item("lc3all").Specific;
                SAPbouiCOM.OptionBtn tnallowed = (SAPbouiCOM.OptionBtn)oForm.Items.Item("lc3nall").Specific;
                tnallowed.GroupWith("lc3all");
                tallowed.Selected = true;

                SAPbouiCOM.OptionBtn tlc3pslc = (SAPbouiCOM.OptionBtn)oForm.Items.Item("lc3pslc").Specific;
                SAPbouiCOM.OptionBtn tlc3psln = (SAPbouiCOM.OptionBtn)oForm.Items.Item("lc3psln").Specific;
                tlc3psln.GroupWith("lc3pslc");
                tlc3pslc.Selected = true; 
                /****/

                SAPbouiCOM.ComboBox cb1 = (SAPbouiCOM.ComboBox)oForm.Items.Item("clctype").Specific;
                cb1.ExpandType = BoExpandType.et_DescriptionOnly;
                cb1.Select("E");

                oMatrix = oForm.Items.Item("matLCLD").Specific;
                AddMatrixRow(oMatrix, "lc2dsn");
            } 
            //oForm.Freeze(false);
        }
        private void CFLConditionPort(string CFLID, string ItemUID)
        {
            oCFL = oForm.ChooseFromLists.Item(CFLID);
            oConds = SBOMain.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions);

            if (CFLID == "CFL_EXPM" || CFLID == "CFL_13" || CFLID == "CFL_14")
            {
                oCond = oConds.Add();
                oCond.Alias = "U_status";
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCond.CondVal = "1";
                oCFL.SetConditions(oConds);
            }
            oCFL = null;
            oCond = null;
            oConds = null;
        }
        #region MatrixSetLine
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
            for (int i = 1; i <= matrix.RowCount; i++)
            {
                matrix.Columns.Item("#").Cells.Item(i).Specific.value = i;
            }
        }
        public void ADDROWMain(SAPbouiCOM.Matrix oMatrix)
        {
            oMatrix.AddRow(1, SBOMain.RightClickLineNum);
            oMatrix.ClearRowData(SBOMain.RightClickLineNum + 1);
            ArrengeMatrixLineNum(oMatrix);
        }
        #endregion
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
                    SAPbouiCOM.Matrix matrix = oForm.Items.Item("matLCATT").Specific;

                    BrowseFilePath.Add(filename);
                    ReplaceFilePath.Add(SBOMain.Get_Attach_Folder_Path() + MyTest.SafeFileName);
                     
                    matrix.AddRow();
                    (matrix.Columns.Item("#").Cells.Item(matrix.RowCount).Specific).Value = Convert.ToString(matrix.RowCount);
                    (matrix.Columns.Item("lc7tp").Cells.Item(matrix.RowCount).Specific).Value = SBOMain.Get_Attach_Folder_Path(); // rec.Fields.Item("DocEntry").Value;
                    (matrix.Columns.Item("lc7fn").Cells.Item(matrix.RowCount).Specific).Value = MyTest.SafeFileName.ToString(); // rec.Fields.Item("U_exinvnode").Value;
                    (matrix.Columns.Item("lc7ad").Cells.Item(matrix.RowCount).Specific).Value = DateTime.Today.ToString("yyyyMMdd"); ; // rec.Fields.Item("U_exinvno").Value;
                    (matrix.Columns.Item("lc7ft").Cells.Item(matrix.RowCount).Specific).Value = null;  // rec.Fields.Item("U_exdd").Value.ToString("yyyyMMdd");
                    (matrix.Columns.Item("lc7cttd").Cells.Item(matrix.RowCount).Specific).Value = null;  // rec.Fields.Item("U_exsbn").Value; //.ToString("yyyyMMdd");

                    //oForm.Items.Item("tattach").Specific.value = filename;

                    System.Windows.Forms.Application.ExitThread();
                }
            }
            catch (Exception ex)
            {
                SBOMain.SBO_Application.MessageBox(ex.Message.ToString(), 1, "ok", "", "");
            }
        }
        private void CFLConditionExp(string CFLID, string ItemUID)
        {
            oCFL = oForm.ChooseFromLists.Item(CFLID);
            oConds = SBOMain.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions);

            if (CFLID == "CFL_PR" || CFLID == "CFL_PQ" || CFLID == "CFL_PR" || CFLID == "CFL_PI")
            {
                   
               /* oCond = oConds.Add();
                oCond.Alias = "validFor";
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCond.CondVal = "Y";
                oCFL.SetConditions(oConds);*/
            }
            oCFL = null;
            oCond = null;
            oConds = null;

        }
        private void CFLCondition(string CFLID, string ItemUID)
        {
            oCFL = oForm.ChooseFromLists.Item(CFLID);
            oConds = SBOMain.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions);

            if (CFLID == "CFL_18")
            { 
                oCond = oConds.Add();
                oCond.Alias = "U_status";
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCond.CondVal = "1"; 
                oCFL.SetConditions(oConds);


            }
            if (CFLID == "CFL_OCRDC")
            {
                if (ItemUID == "lcfc")
                {
                    oCond = oConds.Add();
                    oCond.Alias = "CardType";
                    oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    oCond.CondVal = "C";
                    oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;
                } 
                oCond = oConds.Add();
                oCond.Alias = "validFor";
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCond.CondVal = "Y";

                oCFL.SetConditions(oConds);

            }

            if (CFLID == "CFL_OCRD" || CFLID == "CFL_17")
            {  
                if (ItemUID == "lcfv" || ItemUID == "lc3slc")
                {
                    oCond = oConds.Add();
                    oCond.Alias = "CardType";
                    oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    oCond.CondVal = "S";

                    oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;
                }
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
        private void CFLConditionSO(string CFLID, string ItemUID, string currency)
        {
            oCFL = oForm.ChooseFromLists.Item(CFLID);
            oConds = SBOMain.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions);
             
            if (CFLID == "CFL_SO")
            {     
                oCond = oConds.Add();
                oCond.Alias = "DocCur";
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCond.CondVal = currency;
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
        private void DeleteMatrixBlankRow(SAPbouiCOM.Matrix oMatrix)
        {
            try
            {
                for (int i = oMatrix.VisualRowCount; i >= 1; i--)
                {
                    {
                        if (string.IsNullOrEmpty(((SAPbouiCOM.EditText)oMatrix.Columns.Item("lc2dsn").Cells.Item(i).Specific).Value))
                            oMatrix.DeleteRow(i);
                    }
                }

            }
            catch
            {
            }
        }
        public void doAutoSummatLCLD(SAPbouiCOM.Form oForm)
        {
            SAPbouiCOM.Matrix matLCLD = (SAPbouiCOM.Matrix)oForm.Items.Item("matLCLD").Specific;
            doAutoColSum(matLCLD, "lc2dot");
            doAutoColSum(matLCLD, "lc2am");
            doAutoColSum(matLCLD, "lc2dq");
            doAutoColSum(matLCLD, "lc2aq");
        }
        public void doAutoSummatLCEX(SAPbouiCOM.Form oForm)
        {
            SAPbouiCOM.Matrix matLCEX = (SAPbouiCOM.Matrix)oForm.Items.Item("matLCEX").Specific;
            doAutoColSum(matLCEX, "lc5pf");
            doAutoColSum(matLCEX, "lc5pl");
            doAutoColSum(matLCEX, "lc5af");
            doAutoColSum(matLCEX, "lc5al");
            doAutoColSum(matLCEX, "lc5df");
            doAutoColSum(matLCEX, "lc5dl");
        }
        public void doAutoColSum(SAPbouiCOM.Matrix matrix, string ColumnName)
        {
            SAPbouiCOM.Column mCol = matrix.Columns.Item(ColumnName);
            mCol.RightJustified = true;
            mCol.ColumnSetting.SumType = BoColumnSumType.bst_Auto;
        }

    }
}
