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
using System.Collections.Specialized;
using CoreSuteConnect.Events;
using CoreSuteConnect.Class.DEFAULTSAPFORMS;

namespace CoreSuteConnect.Class.EXIM
{
    class clsSCTrans
    {
        #region VariableDeclaration

        private SAPbouiCOM.Form oForm;
        public string cFormID = string.Empty;

        SAPbouiCOM.EditText exson, exdn, exinvno;
        SAPbouiCOM.ChooseFromList CFL_OINV;

        SAPbouiCOM.ComboBox cb, cb1;
        string schemeType = null;
        string Query = null;

        private SAPbouiCOM.ChooseFromList oCFL = null;
        private SAPbouiCOM.Conditions oConds = null;
        private SAPbouiCOM.Condition oCond = null;

        CommonUtility objCU = new CommonUtility();

        #endregion VariableDeclaration
          
        public clsSCTrans(List<ETTransList> outClass, OutwardFromEximTracking inClass=null)
        {
            try {
                if (inClass != null)
                {
                    if(inClass.ScriptNo != "nofind")
                        { 
                            oForm = SBOMain.SBO_Application.Forms.ActiveForm;
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                            oForm.Items.Item("schsn").Specific.value = inClass.ScriptNo;
                            oForm.Items.Item("1").Click();
                        }
                }
                if (outClass != null)
                {
                    oForm = SBOMain.SBO_Application.Forms.ActiveForm;
                    SAPbouiCOM.Matrix matSTDET = oForm.Items.Item("matSTDET").Specific;
                    double totalApplied = 0;

                    for (int i = 0; i < outClass.Count; i++)
                    {
                        matSTDET.AddRow();
                        ((SAPbouiCOM.EditText)matSTDET.Columns.Item("#").Cells.Item(matSTDET.RowCount).Specific).Value = Convert.ToString(i + 1);
                        ((SAPbouiCOM.EditText)matSTDET.Columns.Item("schexno").Cells.Item(matSTDET.RowCount).Specific).Value = Convert.ToString(outClass[i].eximtrackingno);
                        ((SAPbouiCOM.EditText)matSTDET.Columns.Item("schpcpn").Cells.Item(matSTDET.RowCount).Specific).Value =  outClass[i].portcode  + " " +  outClass[i].portname;
                        ((SAPbouiCOM.EditText)matSTDET.Columns.Item("schinvno").Cells.Item(matSTDET.RowCount).Specific).Value = Convert.ToString(outClass[i].invoiceno);
                        ((SAPbouiCOM.EditText)matSTDET.Columns.Item("schinvde").Cells.Item(matSTDET.RowCount).Specific).Value = Convert.ToString(outClass[i].invoicedocentry);
                        ((SAPbouiCOM.EditText)matSTDET.Columns.Item("schinvdt").Cells.Item(matSTDET.RowCount).Specific).Value = outClass[i].invoicedate.ToString("yyyyMMdd");
                        ((SAPbouiCOM.EditText)matSTDET.Columns.Item("schblno").Cells.Item(matSTDET.RowCount).Specific).Value = Convert.ToString(outClass[i].shippingbillno);
                        ((SAPbouiCOM.EditText)matSTDET.Columns.Item("schbldt").Cells.Item(matSTDET.RowCount).Specific).Value = outClass[i].shippingbilldate.ToString("yyyyMMdd");
                        ((SAPbouiCOM.EditText)matSTDET.Columns.Item("schfob").Cells.Item(matSTDET.RowCount).Specific).Value = Convert.ToString(outClass[i].totalFOB);
                        ((SAPbouiCOM.EditText)matSTDET.Columns.Item("schapamt").Cells.Item(matSTDET.RowCount).Specific).Value = Convert.ToString(outClass[i].appliedamount);

                        totalApplied = totalApplied + outClass[i].appliedamount;

                      //  ((SAPbouiCOM.EditText)matSTDET.Columns.Item("schrcamt").Cells.Item(matSTDET.RowCount).Specific).Value = Convert.ToString(outClass[i].receiveamount);
                    }
                    
                    if (totalApplied > 0) { 
                        oForm.Items.Item("schaa").Specific.Value = Convert.ToString(totalApplied);
                    }
                    doAutoColSum(matSTDET, "schfob");
                    doAutoColSum(matSTDET, "schapamt");
                    doAutoColSum(matSTDET, "schrcamt");

                }
            }
            catch (Exception ex)
            {

            }

        }

        public void doAutoColSum(SAPbouiCOM.Matrix matrix, string ColumnName)
        {
            SAPbouiCOM.Column mCol = matrix.Columns.Item(ColumnName);
            mCol.RightJustified = true;
            mCol.ColumnSetting.SumType = BoColumnSumType.bst_Auto;
        }

        public bool MenuEvent(ref SAPbouiCOM.MenuEvent pVal, string FormId, string Type = "")
        {
            bool BubbleEvent = true;
            try
            {
                cFormID = FormId;
                oForm = SBOMain.SBO_Application.Forms.Item(FormId);
                /*oForm.EnableMenu("1292", true);//Add Row 
                oForm.EnableMenu("1293", true);//Delete Row*/
                if (pVal.BeforeAction == false)
                {
                    if ((oForm.Mode == BoFormMode.fm_ADD_MODE || oForm.Mode == BoFormMode.fm_UPDATE_MODE || Type == "ADDNEWFORM") && (Type != "DEL_ROW") && (Type != "ADD_ROW"))
                    {
                        Form_Load_Components(oForm,"ADD");
                    }
                    if (Type == "navigation")
                    {
                         
                    }
                    else if (Type == "DEL_ROW" || Type == "ADD_ROW")
                    {
                        SAPbouiCOM.Matrix matSTDET = (SAPbouiCOM.Matrix)oForm.Items.Item("matSTDET").Specific;
                        SAPbouiCOM.Matrix matDIT = (SAPbouiCOM.Matrix)oForm.Items.Item("matDIT").Specific;
                       
                        if (Type == "ADD_ROW")
                        {
                            if (SBOMain.RightClickItemID == "matSTDET") {
                                ADDROWMain(matSTDET);
                            }
                            else if (SBOMain.RightClickItemID == "matDIT") {
                                ADDROWMain(matDIT);
                            } 
                        }
                        if (Type == "DEL_ROW")
                        {
                            if (SBOMain.RightClickItemID == "matSTDET")
                            { 
                                DeleteMatrixBlankRow(matSTDET, "schexno");
                                ArrengeMatrixLineNum(matSTDET); 
                            }
                            else if (SBOMain.RightClickItemID == "matDIT")
                            {
                                DeleteMatrixBlankRow(matDIT, "impexno");
                                ArrengeMatrixLineNum(matDIT);
                                
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
        public void ADDROWMain(SAPbouiCOM.Matrix oMatrix)
        {
            oMatrix.AddRow(1, SBOMain.RightClickLineNum);
            oMatrix.ClearRowData(SBOMain.RightClickLineNum + 1);
            ArrengeMatrixLineNum(oMatrix);
        }
        private void ArrengeMatrixLineNum(SAPbouiCOM.Matrix matrix)
        {
            for (int i = 1; i <= matrix.RowCount; i++)
            {
                matrix.Columns.Item("#").Cells.Item(i).Specific.value = i;
            }
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
                    case BoEventTypes.et_FORM_CLOSE:
                        if (pVal.BeforeAction == false)
                        {
                            oForm = null;
                            GC.Collect();
                            GC.WaitForPendingFinalizers();
                        }
                        break;
                    case BoEventTypes.et_COMBO_SELECT:
                        oForm = SBOMain.SBO_Application.Forms.ActiveForm;
                        if (pVal.BeforeAction == false)
                        {
                            if (pVal.ItemUID == "cSer" && pVal.FormMode == 3)
                            {
                                oForm.Items.Item("tDocNum").Specific.Value = oForm.BusinessObject.GetNextSerialNumber(oForm.Items.Item("cSer").Specific.Value, "EXRU");
                            }
                        }
                        break;
                    case BoEventTypes.et_FORM_LOAD:
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
                        if (pVal.BeforeAction == true)
                        {
                            if (pVal.ItemUID == "1" && (oForm.Mode == BoFormMode.fm_ADD_MODE || oForm.Mode == BoFormMode.fm_UPDATE_MODE))
                            {

                               // bool JE = crateJurnalEntry();

                                if (string.IsNullOrEmpty(oForm.Items.Item("sched").Specific.value))
                                {
                                    BubbleEvent = false;
                                    SBOMain.SBO_Application.StatusBar.SetText("Please enter Scheme expire date.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                    oForm.Items.Item("sched").Click();
                                }

                                cb1 = oForm.Items.Item("schtype").Specific;
                                if (string.IsNullOrEmpty(oForm.Items.Item("schsn").Specific.value))
                                {
                                    if (cb1.Selected.Value.ToString() == "DBK")
                                    {
                                        BubbleEvent = false;
                                        SBOMain.SBO_Application.StatusBar.SetText("Script No is Mandatory when script type is DBK.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                        oForm.Items.Item("schsn").Click();
                                    }
                                }
                                if(BubbleEvent == true)
                                { 
                                    SAPbouiCOM.Matrix matQt = (SAPbouiCOM.Matrix)oForm.Items.Item("matSTDET").Specific;
                                    string DocNum = Convert.ToString(oForm.Items.Item("tDocNum").Specific.Value);
                                      
                                    if (matQt.RowCount > 0)
                                    {
                                        for (int i = 1; i <= matQt.RowCount; i++)
                                        {
                                            string docentr = ((SAPbouiCOM.EditText)matQt.Columns.Item("schinvde").Cells.Item(i).Specific).Value;
                                            string recAmt = ((SAPbouiCOM.EditText)matQt.Columns.Item("schrcamt").Cells.Item(i).Specific).Value;

                                           bool JEEntry =  crateJurnalEntry();

                                            if (cb1.Selected.Value.ToString() == "DBK")
                                            { 
                                                string getQuery1 = @"Update dbo.[@XET6] Set U_received = '" + recAmt + "' , U_dbkno = '" + oForm.Items.Item("schsn").Specific.Value + "' Where DocEntry = ( SELECT DISTINCT(T0.DocEntry)  ";
                                                getQuery1 = getQuery1 + " FROM dbo.[@EXET] AS T0 LEFT JOIN dbo.[@XET10] AS T1 ON T0.DocEntry = T1.docEntry WHERE U_ex10sc IS NOT NULL AND U_exinvnode = '" + docentr + "' )";

                                                SAPbobsCOM.Recordset rec1;
                                                rec1 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                                rec1.DoQuery(getQuery1);
                                            }
                                            else if (cb1.Selected.Value.ToString() == "RoDTEP")
                                            { 
                                                string getQuery1 = @"Update dbo.[@XET12] Set U_rodprec = '" + recAmt + "' , U_rodno = '" + oForm.Items.Item("schsn").Specific.Value + "' Where DocEntry = ( SELECT DISTINCT(T0.DocEntry)  ";
                                                getQuery1 = getQuery1 + " FROM dbo.[@EXET] AS T0 LEFT JOIN dbo.[@XET10] AS T1 ON T0.DocEntry = T1.docEntry WHERE U_ex10sc IS NOT NULL AND U_exinvnode = '" + docentr + "' )";

                                                SAPbobsCOM.Recordset rec1;
                                                rec1 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                                rec1.DoQuery(getQuery1);
                                            }  
                                        }
                                    }
                                }
                            }
                        }
                        if (pVal.BeforeAction == false)
                        {
                            if (pVal.ItemUID == "1" && oForm.Mode == BoFormMode.fm_ADD_MODE)
                            {
                                Form_Load_Components(oForm, "OK");
                            }
                        }
                        break;
                    case BoEventTypes.et_KEY_DOWN:
                        if (pVal.BeforeAction == false)
                        {
                            if (pVal.CharPressed == 9)
                            {
                                oForm = SBOMain.SBO_Application.Forms.ActiveForm;
                                if (pVal.ItemUID == "matSTDET" && pVal.ColUID == "schrcamt")
                                {
                                    SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("matSTDET").Specific;
                                    double totalrem = 0;
                                    for (int i = 1; i <= oMatrix.RowCount; i++)
                                    {
                                        totalrem = totalrem + Convert.ToDouble((oMatrix.Columns.Item("schrcamt").Cells.Item(i).Specific).Value);
                                    }
                                    oForm.Items.Item("schra").Specific.Value = Convert.ToString(totalrem);
                                    oForm.Items.Item("schrma").Specific.Value = Convert.ToString(totalrem - Convert.ToDouble(oForm.Items.Item("schua").Specific.Value));
                                }
                            }
                        } 
                        break;
                    case BoEventTypes.et_CLICK:
                        oForm = SBOMain.SBO_Application.Forms.ActiveForm;
                        if (pVal.BeforeAction == true)
                        {
                            if (pVal.ItemUID == "btnPFET")
                            {
                                cb1 = oForm.Items.Item("schtype").Specific;
                                if (cb1.Selected.Value.ToString() == "-")
                                {
                                    SBOMain.SBO_Application.StatusBar.SetText("Please Select Scheme Type!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                    BubbleEvent = false;
                                }
                            }
                        }
                        if (pVal.BeforeAction == false)
                        {
                            if (pVal.ItemUID == "matSTDET" && pVal.ColUID == "schexno")
                            {
                                SAPbouiCOM.Matrix matrix = oForm.Items.Item("matSTDET").Specific;
                                string abc = (matrix.Columns.Item("schexno").Cells.Item(pVal.Row).Specific).Value;

                                objCU.FormLoadAndActivate("frmETTrans", "mnsmEXIM013");  
                                OutwardToEximTracking outEximTracking = new OutwardToEximTracking(); 
                                outEximTracking.DocEntry = abc;
                                outEximTracking.FromFrmName = "FindMode";

                                clsExTrans oPrice = new clsExTrans(outEximTracking);
                                //oForm.Close();
                            }

                            if (pVal.ItemUID == "matDIT" && pVal.ColUID == "impexno")
                            {
                                SAPbouiCOM.Matrix matrix = oForm.Items.Item("matDIT").Specific;
                                string abc = (matrix.Columns.Item("impexno").Cells.Item(pVal.Row).Specific).Value;

                                objCU.FormLoadAndActivate("frmETTrans", "mnsmEXIM013"); 
                                OutwardToEximTracking outEximTracking = new OutwardToEximTracking();

                                outEximTracking.DocEntry = abc;
                                outEximTracking.FromFrmName = "FindMode";

                                clsExTrans oPrice = new clsExTrans(outEximTracking);
                                //oForm.Close();
                            }
                            if (pVal.ItemUID == "btnREF")
                            {
                                string scrno = oForm.Items.Item("schsn").Specific.value;
                                string getDocEntry = " select sum(T1.U_ex13sua) as 'Utilized' from dbo.[@EXET] AS T0 LEFT JOIN dbo.[@xet13] ";
                                       getDocEntry = getDocEntry + " AS T1 on T0.DocEntry = T1.DocEntry where T1.U_ex13sn = '" + scrno+"'";
                                SAPbobsCOM.Recordset rec = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                rec.DoQuery(getDocEntry);
                                if (rec.RecordCount > 0)
                                {
                                    oForm.Items.Item("schua").Specific.value = Convert.ToString(rec.Fields.Item("Utilized").Value);
                                }
                                oForm.Items.Item("schrma").Specific.value = Convert.ToString(Convert.ToDouble(oForm.Items.Item("schra").Specific.value)-rec.Fields.Item("Utilized").Value);

                                ///
                                oForm.Items.Item("tab1").Visible = true;
                                oForm.Items.Item("tab1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                oForm.PaneLevel = 1;

                                SAPbouiCOM.Matrix matrix = oForm.Items.Item("matDIT").Specific;

                                DeleteMatrixAllRowmatDIT(matrix);
                                getDocEntry = " select  T0.DocEntry as 'EximNo', T0.U_exinvno, T0.U_exinvnode, T0.*, T1.* from dbo.[@EXET] AS T0 LEFT JOIN dbo.[@xet13] ";
                                getDocEntry = getDocEntry + " AS T1 on T0.DocEntry = T1.DocEntry where T1.U_ex13sn = '" + scrno + "'";
                                SAPbobsCOM.Recordset rec1 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                rec1.DoQuery(getDocEntry);
                                if (rec1.RecordCount > 0)
                                {
                                    while (!rec1.EoF)
                                    {
                                        matrix.AddRow();
                                        (matrix.Columns.Item("#").Cells.Item(matrix.RowCount).Specific).Value = Convert.ToString(matrix.RowCount);
                                        (matrix.Columns.Item("impexno").Cells.Item(matrix.RowCount).Specific).Value = rec1.Fields.Item("EximNo").Value;
                                        (matrix.Columns.Item("impinvno").Cells.Item(matrix.RowCount).Specific).Value = rec1.Fields.Item("U_exinvno").Value;
                                        (matrix.Columns.Item("impinvde").Cells.Item(matrix.RowCount).Specific).Value = rec1.Fields.Item("U_exinvnode").Value;
                                        (matrix.Columns.Item("impinvdt").Cells.Item(matrix.RowCount).Specific).Value = rec1.Fields.Item("U_exdd").Value.ToString("yyyyMMdd");
                                        (matrix.Columns.Item("impbeno").Cells.Item(matrix.RowCount).Specific).Value = rec1.Fields.Item("U_exsbn").Value; //.ToString("yyyyMMdd");
                                        (matrix.Columns.Item("impbedt").Cells.Item(matrix.RowCount).Specific).Value = rec1.Fields.Item("U_exsbd").Value.ToString("yyyyMMdd");
                                        
                                        (matrix.Columns.Item("impAV").Cells.Item(matrix.RowCount).Specific).Value = rec1.Fields.Item("U_ex13bv").Value;
                                        (matrix.Columns.Item("impUA").Cells.Item(matrix.RowCount).Specific).Value = rec1.Fields.Item("U_ex13sua").Value; 
                                        rec1.MoveNext();
                                    }
                                }
                                DeleteMatrixBlankRowmatDIT(matrix);
                            }

                            if (pVal.ItemUID == "btnPFET")
                               {
                                bool plFormOpen = false;
                                for (int i = 1; i < SBOMain.SBO_Application.Forms.Count; i++)
                                {
                                    if (SBOMain.SBO_Application.Forms.Item(i).UniqueID == "frmETTransList")
                                    {
                                        SBOMain.SBO_Application.Forms.Item(i).Select();
                                        plFormOpen = true;
                                    }
                                }
                                if (!plFormOpen)
                                {
                                    SBOMain.LoadFromXML("frmETTransList", "EXIM");
                                    SBOMain.SBO_Application.Forms.Item("frmETTransList").Select();

                                    var oForm1 = SBOMain.SBO_Application.Forms.ActiveForm;
                                    oForm1.DataSources.DataTables.Add("tab");
                                    SAPbouiCOM.Grid objGrid = oForm1.Items.Item("grid").Specific;

                                    SAPbouiCOM.ComboBox cb4 = oForm.Items.Item("schtype").Specific;
                                    string type = cb4.Selected.Value.ToString();
                                    string getDocEntry1;
                                    if (type == "DBK")
                                    {       
                                      getDocEntry1 = " SELECT 'N' AS CHK, T0.DocEntry as 'Exim Tracking No', T2.U_ex1pol as 'PortCode', (SELECT Distinct(U_portname) FROM dbo.[@EXPM] where U_portcode = T2.U_ex1pol)  as 'PortName', ";
                                        getDocEntry1 = getDocEntry1 + " T0.U_exinvno as 'Invoice no', T0.U_exinvnode as 'Invoice DocEntry', T3.DocDate as 'Invoice Date',  T0.U_exsbn as 'Shipping bill no', ";
                                        getDocEntry1 = getDocEntry1 + " T0.U_exsbd as 'Shipping bill Date',   sum(T1.U_ex10fbv) as 'Total FOB',  sum(T1.U_ex10fv) as 'Applied Amount' ";
                                        getDocEntry1 = getDocEntry1 + " FROM dbo.[@EXET] AS T0   LEFT JOIN dbo.[@XET10] AS T1 ON T0.DocEntry = T1.DocEntry ";
                                        getDocEntry1 = getDocEntry1 + " LEFT JOIN dbo.[@XET1] AS T2 ON T0.DocEntry = T2.DocEntry LEFT JOIN OINV AS T3 ON T3.DocEntry =T0.U_exinvnode  WHERE T0.DocEntry Not in (SELECT U_schexno FROM dbo.[@XRU1] where U_schexno is not null ) AND  T1.U_ex10sv is not null   AND T0.U_exson is not null AND ";
                                        getDocEntry1 = getDocEntry1 + " T1.U_ex10fbv is not null Group By  T0.DocEntry, T0.U_exinvno, T0.U_exinvnode, T0.U_exsbn, T0.U_exsbd , T2.U_ex1pol ,T3.DocDate ";
                                    }
                                    else { 
                                      getDocEntry1 = " SELECT 'N' AS CHK, T0.DocEntry as 'Exim Tracking No', T2.U_ex1pol as 'PortCode', (SELECT Distinct(U_portname) FROM dbo.[@EXPM] where U_portcode = T2.U_ex1pol)  as 'PortName', ";
                                        getDocEntry1 = getDocEntry1 + " T0.U_exinvno as 'Invoice no', T0.U_exinvnode as 'Invoice DocEntry', T3.DocDate as 'Invoice Date',  T0.U_exsbn as 'Shipping bill no', ";
                                        getDocEntry1 = getDocEntry1 + " T0.U_exsbd as 'Shipping bill Date',   sum(T1.U_ex11fbv) as 'Total FOB',  sum(T1.U_ex11fv) as 'Applied Amount' ";
                                        getDocEntry1 = getDocEntry1 + "   FROM dbo.[@EXET] AS T0   LEFT JOIN dbo.[@XET11] AS T1 ON T0.DocEntry = T1.DocEntry ";
                                        getDocEntry1 = getDocEntry1 + " LEFT JOIN dbo.[@XET1] AS T2 ON T0.DocEntry = T2.DocEntry LEFT JOIN OINV AS T3 ON T3.DocEntry =T0.U_exinvnode  WHERE T0.DocEntry Not in (SELECT U_schexno FROM dbo.[@XRU1] where U_schexno is not null ) AND T1.U_ex11sv is not null   AND T0.U_exson is not null AND ";
                                        getDocEntry1 = getDocEntry1 + " T1.U_ex11fbv is not null Group By  T0.DocEntry, T0.U_exinvno, T0.U_exinvnode, T0.U_exsbn, T0.U_exsbd , T2.U_ex1pol ,T3.DocDate ";
                                    }

                                    oForm1.DataSources.DataTables.Item("tab").ExecuteQuery(getDocEntry1);
                                    objGrid.DataTable = oForm1.DataSources.DataTables.Item("tab");
                                    objGrid.Columns.Item(0).Type = BoGridColumnType.gct_CheckBox;
                                    objGrid.Columns.Item(0).TitleObject.Caption = "Select"; 
                                }
                            }
                        }
                        break;
                }
            }
            catch (Exception ex)
            {

            }
            finally
            {
                /*if (oForm != null)
                    oForm.Freeze(false);*/
            }
            return BubbleEvent;
        }

        public bool crateJurnalEntry()
        {
            string ErrorStr = "";
            bool ReturnValue = false;
            SAPbobsCOM.JournalEntries doc = null;
            try
            {
                if (SBOMain.oCompany.Connected)
                { 
                    doc = (SAPbobsCOM.JournalEntries)SBOMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);
                    
                    string ref2 = "ref2";
                    //string ref1 = Description;
                    doc.TaxDate = DateTime.Today;
                    doc.ReferenceDate = DateTime.Today;
                    doc.Reference2 = "Reference2";
                    doc.DueDate = DateTime.Today;  

                    doc.Memo = "Test JE"; 
                    /*foreach (JournalEntryItem item in oJournalEntry.items)
                    {*/
                        //doc.Lines.FCCurrency = "INR";  
                        doc.Lines.AccountCode = "L104010001";  
                        doc.Lines.Credit = Math.Abs(100);  
                        doc.Lines.Reference2 = ref2;
                        doc.Lines.Reference1 = "Reference1";
                        //string lineDetail = ref2 + "|" + "Amazon fees" + "|" + item.Description + "|" + doc.DueDate.ToString("yyyyMMdd") + "|" + item.Amount.ToString();
                       // doc.Lines.UserFields.Fields.Item("U_amz_Det").Value = lineDetail;
                        doc.Lines.Add();

                    /*}*/

                    /*if (JeTotal != 0.00)
                    {*/
                       // doc.Lines.FCCurrency = "INR";
                        doc.Lines.AccountCode = "A201010002";
                        //doc.Lines.ShortName = oJournalEntry.CardCode;
                        doc.Lines.Debit = Math.Abs(100); 
                        doc.Lines.Reference2 = ref2;
                        doc.Lines.Add();
                    /*}*/
                    int Result = doc.Add();


                    if (Result != 0)
                    {
                        ReturnValue = false;
                        //SBOMain.oCompany.GetLastError(out lErrCode, out sErrMsg);
                       // ErrorStr = sErrMsg;
                    }
                    else
                    {
                        ReturnValue = true;
                        //sErrMsg = "";
                        ErrorStr = String.Empty;
                    }
                }
                else
                {
                    ReturnValue = false;
                    ErrorStr = "Company not connected.";
                }
            }
            catch (Exception exec)
            {
                ReturnValue = false; ErrorStr = "System Error : " + exec.Message;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(doc);
                doc = null;
            }

            return ReturnValue;
        }
        private void DeleteMatrixAllRowmatDIT(SAPbouiCOM.Matrix oMatrix)
        {
            try
            {
                for (int i=1; i <= oMatrix.VisualRowCount;  i++)
                {  
                   oMatrix.DeleteRow(i);
                }
            }
            catch (Exception ex)
            {
                //SBOMain.SBO_Application.MessageBox(ex.Message.ToString(), 1, "ok", "", "");
            }
        }
        private void DeleteMatrixBlankRowmatDIT(SAPbouiCOM.Matrix oMatrix)
        {
            try
            {
                for (int i = oMatrix.VisualRowCount; i >= 1; i--)
                {
                    {
                        if (string.IsNullOrEmpty(((SAPbouiCOM.EditText)oMatrix.Columns.Item("impinvno").Cells.Item(i).Specific).Value))
                        {
                            oMatrix.DeleteRow(i);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                //SBOMain.SBO_Application.MessageBox(ex.Message.ToString(), 1, "ok", "", "");
            }
        }

        public void Form_Load_Components(SAPbouiCOM.Form oForm, string mode)
        {
            oForm.Freeze(true);
            oForm.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized;
            SAPbouiCOM.EditText oEdit;
            string Table = "@EXRU";
            DateTime now = DateTime.Now;
            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
            oEdit = oForm.Items.Item("tCode").Specific;
            objCU.GetNextDocNum(ref oEdit, ref Table);
            oForm.Items.Item("tDocNum").Specific.value = oForm.BusinessObject.GetNextSerialNumber("DocEntry", "EXRU");
            Events.Series.SeriesCombo("EXRU", "cSer");
            oForm.Items.Item("cSer").DisplayDesc = true;

            oForm.Items.Item("tab1").Visible = true;
            oForm.Items.Item("tab1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
            oForm.PaneLevel = 1;

            oForm.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized;
            oForm.Items.Item("schsd").Specific.value = DateTime.Now.ToString("yyyyMMdd");

            cb = (SAPbouiCOM.ComboBox)oForm.Items.Item("schst").Specific;
            cb.ExpandType = BoExpandType.et_DescriptionOnly;
            cb.Select("O");

            cb1 = (SAPbouiCOM.ComboBox)oForm.Items.Item("schtype").Specific;
            cb1.ExpandType = BoExpandType.et_DescriptionOnly;
            cb1.Select("-");

            SAPbouiCOM.OptionBtn tstatus = (SAPbouiCOM.OptionBtn)oForm.Items.Item("tstatus").Specific;
            SAPbouiCOM.OptionBtn tstatusI = (SAPbouiCOM.OptionBtn)oForm.Items.Item("tstatusI").Specific;
            tstatusI.GroupWith("tstatus"); 
            tstatus.Selected = true;

            oForm.Freeze(false);
        }


    }
}
