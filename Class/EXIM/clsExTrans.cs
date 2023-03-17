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
    class clsExTrans
    {
        //[DllImport("user32.dll")]
        //private static extern IntPtr GetForegroundWindow();

        #region VariableDeclaration

        private SAPbouiCOM.Form oForm;
        public string cFormID = string.Empty;

        SAPbouiCOM.EditText exson, exdn, exinvno, et3yc;

        private SAPbouiCOM.ChooseFromList oCFL = null;
        private SAPbouiCOM.Conditions oConds = null;
        private SAPbouiCOM.Condition oCond = null;

        SAPbouiCOM.ChooseFromList CFL_PR, CFL_PQ, CFL_PO, CFL_PI;
        SAPbouiCOM.ChooseFromList CFL_ORDR, CFL_ODLN, CFL_OINV, CFL_OPOR, CFL_OPDN, CFL_OPCH;

        SAPbouiCOM.LinkedButton oLinkSOPO = null;
        SAPbouiCOM.LinkedButton oLinkDLGR = null;
        SAPbouiCOM.LinkedButton oLinkINAP = null;

        SAPbouiCOM.LinkedButton oLinkPR = null;
        SAPbouiCOM.LinkedButton oLinkPQ = null;
        SAPbouiCOM.LinkedButton oLinkPO = null;
        SAPbouiCOM.LinkedButton oLinkPI = null;

        SAPbouiCOM.ComboBox cb1, cb4, cb2, cb5;

        int LineId;

        string DocType = null;
        string Query = null;
        string Query1 = null;
        string Code = null;
        string portcode = null;
        string expcode = null;
        string itemcode = null;
        string liceneno = null;
        string getDocEntry = null;
        string BrowseFilePath = string.Empty;

        double qty = 0;
        double amtLC = 0;
        double amtFC = 0;
        double flqty = 0;
        double fllC = 0;
        double flfC = 0;
        double freight = 0;
        double insurance = 0;
        double othercharge = 0;
        double exchangeRate = 0;
        double fobROW = 0;
        double freightROW = 0;
        double InsuranceROW = 0;
        double schrateROW = 0;
        double ex10vrpkROW = 0;
        double finalvalROW = 0;
        double ex4lfqty = 0;
        double ex4lffc = 0;
        double ex4lflc = 0;

        CommonUtility objCU = new CommonUtility();

        #endregion VariableDeclaration


        private void SBO_Application_RightClickEvent(ref ContextMenuInfo eventInfo, out bool BubbleEvent)
        {
            SBO_Application_RightClickEvent(ref eventInfo, out BubbleEvent);
            //if (eventInfo.ItemUID == "mtItems" && eventInfo.ActionSuccess == true)
            if (eventInfo.ItemUID == "matEXPAC")
            {
                //DeleterowNumber = eventInfo.Row; 
            }
        }

        public clsExTrans(OutwardToEximTracking outClass)
        {

            try
            {
                if (outClass != null)
                {
                    oForm = SBOMain.SBO_Application.Forms.ActiveForm;

                    if (outClass.FromFrmName == "FindMode")
                    {
                        assignFormValues(oForm, outClass);
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                        oForm.Items.Item("tCode").Specific.value = outClass.DocEntry;
                        oForm.Items.Item("1").Click();
                    }
                    else if (outClass.FromFrmName == "PurchaseOrder")
                    {
                        cb1 = (SAPbouiCOM.ComboBox)oForm.Items.Item("exdt").Specific;
                        cb1.ExpandType = BoExpandType.et_DescriptionOnly;
                        cb1.Select("I");
                        setChooseFromListForImport(oForm);
                        assignFormValues(oForm, outClass);
                        oForm.Items.Item("exsonde").Specific.value = outClass.SODocEnt;
                        oForm.Items.Item("exson").Specific.value = outClass.SODocNo;
                    }
                    else if (outClass.FromFrmName == "PurchaseOrderEXIST")
                    {
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                        oForm.Items.Item("exsonde").Specific.value = outClass.SODocEnt;
                        oForm.Items.Item("1").Click();
                    }
                    else if (outClass.FromFrmName == "SalesOrderEXIST")
                    {
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                        oForm.Items.Item("exsonde").Specific.value = outClass.SODocEnt;
                        oForm.Items.Item("1").Click();
                    }
                    else if (outClass.FromFrmName == "SalesOrder")
                    {
                        assignFormValues(oForm, outClass);
                        oForm.Items.Item("exsonde").Specific.value = outClass.SODocEnt;
                        oForm.Items.Item("exson").Specific.value = outClass.SODocNo;
                    }
                    else if (outClass.FromFrmName == "SOexitDLNot")
                    {
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                        oForm.Items.Item("exsonde").Specific.value = outClass.SODocEnt;
                        oForm.Items.Item("exson").Specific.value = outClass.SODocNo;
                        oForm.Items.Item("1").Click();
                        oForm.Items.Item("exdnde").Specific.value = outClass.DocEntry;
                        oForm.Items.Item("exdn").Specific.value = outClass.DocNum;
                    }
                    else if (outClass.FromFrmName == "DeliveryEXIST")
                    {
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                        oForm.Items.Item("exdnde").Specific.value = outClass.DelDocEnt;
                        oForm.Items.Item("1").Click();
                    }
                    else if (outClass.FromFrmName == "DeliveryEXISTARNot")
                    {
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                        oForm.Items.Item("exdnde").Specific.value = outClass.DelDocEnt;
                        oForm.Items.Item("1").Click();
                        oForm.Items.Item("exinvnode").Specific.value = outClass.DocEntry;
                        oForm.Items.Item("exinvno").Specific.value = outClass.DocNum;

                        string statusval = oForm.Items.Item("exls").Specific.value.ToString();

                        if (statusval == "Licence"){
                            fillRLFMatrix(oForm, outClass.DocEntry, outClass.BPCode);
                        }
                        else if (statusval != "Licence"){
                            fillDBKMatrix(oForm, outClass.DocEntry, outClass.BPCode);
                            fillRoDTEPMatrix(oForm, outClass.DocEntry, outClass.BPCode);
                        }
                    }
                    else if (outClass.FromFrmName == "Delivery")
                    {
                        assignFormValues(oForm, outClass);
                        oForm.Items.Item("exdnde").Specific.value = outClass.DelDocEnt;
                        oForm.Items.Item("exdn").Specific.value = outClass.DelDocNo;
                        oForm.Items.Item("exsonde").Specific.value = outClass.SODocEnt;
                        oForm.Items.Item("exson").Specific.value = outClass.SODocNo;
                    }

                    else if (outClass.FromFrmName == "GRPOEXIST")
                    {
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                        oForm.Items.Item("exdnde").Specific.value = outClass.DelDocEnt;
                        oForm.Items.Item("1").Click();
                        NameValueCollection list3 = new NameValueCollection()
                                                 {  { "1", "Vendor Code" }, { "2","Vendor Name" }, { "3", "Purchase Order No" } ,
                                                    { "4", "GRN No" }, { "5","A/P Invoice No" }, { "6", "A/P Invoice Ex. Rate" },
                                                    { "7","Bill of Entry No" }, { "8", "Bill of Entry Date" }};

                        // It will set Label Caption based on Doc Type
                        setLableText(oForm, list3);
                        setChooseFromListForImport(oForm);
                    }
                    else if (outClass.FromFrmName == "GRPO")
                    {
                        cb1 = (SAPbouiCOM.ComboBox)oForm.Items.Item("exdt").Specific;
                        cb1.ExpandType = BoExpandType.et_DescriptionOnly;
                        cb1.Select("I");
                        NameValueCollection list3 = new NameValueCollection()
                                                 {  { "1", "Vendor Code" }, { "2","Vendor Name" }, { "3", "Purchase Order No" } ,
                                                    { "4", "GRN No" }, { "5","A/P Invoice No" }, { "6", "A/P Invoice Ex. Rate" },
                                                    { "7","Bill of Entry No" }, { "8", "Bill of Entry Date" }};

                        // It will set Label Caption based on Doc Type
                        setLableText(oForm, list3);

                        setChooseFromListForImport(oForm);
                        assignFormValues(oForm, outClass);
                        oForm.Items.Item("exdnde").Specific.value = outClass.DelDocEnt;
                        oForm.Items.Item("exdn").Specific.value = outClass.DelDocNo;
                        oForm.Items.Item("exsonde").Specific.value = outClass.SODocEnt;
                        //oForm.Items.Item("exson").Specific.value = outClass.SODocNo;
                        fillRoDTEPScriptMatrix(oForm, outClass.DelDocEnt, outClass.BPCode);
                    }
                    else if (outClass.FromFrmName == "ARInvoiceEXIST")
                    {
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                        oForm.Items.Item("exinvnode").Specific.value = outClass.DocEntry;
                        oForm.Items.Item("1").Click();
                    }
                    else if (outClass.FromFrmName == "ARInvoice")
                    {
                        assignFormValues(oForm, outClass);
                        oForm.Items.Item("exdd").Specific.value = outClass.DocDate;

                        oForm.Items.Item("exinvnode").Specific.value = outClass.DocEntry;
                        oForm.Items.Item("exinvno").Specific.value = outClass.DocNum;

                        oForm.Items.Item("exdnde").Specific.value = outClass.DelDocEnt;
                        oForm.Items.Item("exdn").Specific.value = outClass.DelDocNo;

                        oForm.Items.Item("exsonde").Specific.value = outClass.SODocEnt;
                        oForm.Items.Item("exson").Specific.value = outClass.SODocNo;

                        freight = Convert.ToDouble(oForm.Items.Item("frtAFC").Specific.value);
                        insurance = Convert.ToDouble(oForm.Items.Item("insAFC").Specific.value);
                        exchangeRate = Convert.ToDouble(oForm.Items.Item("exinvdt").Specific.value);

                        string statusval = oForm.Items.Item("exls").Specific.value.ToString();

                        if (statusval == "Licence")
                        {
                            fillRLFMatrix(oForm, outClass.DocEntry, outClass.BPCode);
                        }
                        else if (statusval != "Licence")
                        {
                            fillDBKMatrix(oForm, outClass.DocEntry, outClass.BPCode);
                            fillRoDTEPMatrix(oForm, outClass.DocEntry, outClass.BPCode);
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
        }
       
        public void assignFormValues(SAPbouiCOM.Form oForm, OutwardToEximTracking outClass)
        {
            oForm.Items.Item("exbc").Specific.value = outClass.BPCode;
            oForm.Items.Item("exbn").Specific.value = outClass.BPName;
            oForm.Items.Item("exinco").Specific.value = outClass.Incoterm;
            oForm.Items.Item("expcgb").Specific.value = outClass.PrecarriageBy;
            oForm.Items.Item("expcrb").Specific.value = outClass.PrecarrierBy;
            oForm.Items.Item("ex1pol").Specific.value = outClass.Portofloading;
            oForm.Items.Item("ex1pod").Specific.value = outClass.Portlfdischarge;
            oForm.Items.Item("ex1por").Specific.value = outClass.Portofreceipt;
            oForm.Items.Item("ex1fd").Specific.value = outClass.Finaldestination;
            oForm.Items.Item("ex1coo").Specific.value = outClass.Countryoforigin;
            oForm.Items.Item("ex1dc").Specific.value = outClass.DestinationCountry;
        }
       
        public void fillRoDTEPScriptMatrix(SAPbouiCOM.Form oForm, string DocEntry, string BPcode)
        {

            oForm.Items.Item("tab9").Visible = true;
            oForm.Items.Item("tab9").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
            oForm.PaneLevel = 9;
            SAPbouiCOM.Matrix matrix = oForm.Items.Item("matRS").Specific;

            int i = 1;
            getDocEntry = "SELECT SUM(T1.LineTotal) As 'BasicValue'  FROM OPDN AS T0 LEFT JOIN PDN1 AS T1 ON T0.DocEntry = T1.DocEntry WHERE T0.DocEntry = '" + DocEntry + "'  ";
            SAPbobsCOM.Recordset rec = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            rec.DoQuery(getDocEntry);
            if (rec.RecordCount > 0)
            {
                while (!rec.EoF)
                {
                    matrix.AddRow();
                    (matrix.Columns.Item("#").Cells.Item(i).Specific).Value = Convert.ToString(matrix.RowCount);
                    (matrix.Columns.Item("ex13st").Cells.Item(i).Specific).Value = "RodTEP";
                    (matrix.Columns.Item("ex13bv").Cells.Item(i).Specific).Value = rec.Fields.Item("BasicValue").Value;
                    i++;
                    rec.MoveNext();
                }
            }
        }

        public void fillRLFMatrix(SAPbouiCOM.Form oForm, string DocEntry, string BPcode)
        {
            try
            {

                string getDocEntry2 = null;


                oForm.Items.Item("tab4").Visible = true;
                oForm.Items.Item("tab4").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                oForm.PaneLevel = 4;

                int i = 1;
                SAPbouiCOM.Matrix matrix = oForm.Items.Item("matRLF1").Specific;
                SAPbouiCOM.Matrix matRLF2 = oForm.Items.Item("matRLF2").Specific;

                SAPbobsCOM.Recordset rec = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                string getDocEntry = " SELECT DISTINCT(T2.U_EXIMLIC) as ItemCode, (SELECT T5.ITEMNAME FROM OITM As T5 WHERE T5.ITEMCODE = T2.U_EXIMLIC) as ItemName FROM INV1 AS T0 LEFT JOIN OINV AS T1 ON T0.DocEntry = T1.DocEntry ";
                getDocEntry = getDocEntry + " LEFT JOIN OITM AS T2 ON T0.ItemCode = T2.ItemCode LEFT JOIN OITM AS T3 ON T2.ItemCode = T3.ItemCode ";
                getDocEntry = getDocEntry + " WHERE T0.DocEntry = '" + DocEntry + "' AND T1.CardCode = '" + BPcode + "'";

                rec.DoQuery(getDocEntry);
                if (rec.RecordCount > 0)
                {
                    while (!rec.EoF)
                    {
                        matrix.AddRow();
                        (matrix.Columns.Item("#").Cells.Item(i).Specific).Value = Convert.ToString(matrix.RowCount);
                        (matrix.Columns.Item("ex4ic").Cells.Item(i).Specific).Value = rec.Fields.Item("ItemCode").Value;
                        (matrix.Columns.Item("ex4in").Cells.Item(i).Specific).Value = rec.Fields.Item("ItemName").Value;

                        SAPbobsCOM.Recordset rec2 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                        getDocEntry2 = " SELECT T2.U_EXIMLIC, SUM(T0.Quantity) AS 'Total_Qty', SUM(T0.LineTotal) AS 'TotalLC', SUM(T0.Totalfrgn) AS 'TotalFC' ";
                        getDocEntry2 = getDocEntry2 + " FROM INV1 AS T0 LEFT JOIN OINV AS T1 ON T0.DocEntry = T1.DocEntry LEFT JOIN OITM AS T2 ON T0.ItemCode = T2.ItemCode ";
                        getDocEntry2 = getDocEntry2 + " LEFT JOIN OITM AS T3 ON T2.ItemCode = T3.ItemCode WHERE T0.DocEntry = '" + DocEntry + "' AND T1.CardCode = '" + BPcode + "' AND ";
                        getDocEntry2 = getDocEntry2 + " T2.U_EXIMLIC = '" + rec.Fields.Item("ItemCode").Value + "' Group By T2.U_EXIMLIC";
                        rec2.DoQuery(getDocEntry2);
                        (matrix.Columns.Item("ex4lfqty").Cells.Item(i).Specific).Value = rec2.Fields.Item("Total_Qty").Value;
                        (matrix.Columns.Item("ex4lffc").Cells.Item(i).Specific).Value = rec2.Fields.Item("TotalFC").Value;
                        (matrix.Columns.Item("ex4lflc").Cells.Item(i).Specific).Value = rec2.Fields.Item("TotalLC").Value;

                        i++;
                        rec.MoveNext();
                    }
                }
                oForm.Freeze(true);
                SAPbouiCOM.Matrix matETSC1 = oForm.Items.Item("matETSC1").Specific;
                SAPbouiCOM.Matrix matETSC2 = oForm.Items.Item("matETSC2").Specific;
                DeleteMatrixAll(matETSC1);
                DeleteMatrixAll(matETSC2);
                DeleteMatrixBlankRowRFL(matrix);
                DeleteMatrixBlankRowRFL2(matRLF2);
                ArrengeMatrixLineNum(matrix);
                ArrengeMatrixLineNum(matRLF2);
                oForm.Freeze(false);
            }
            catch (Exception ex)
            {

            }
        }

        public void fillRLFMatrixBKP(SAPbouiCOM.Form oForm, string DocEntry, string BPcode)
        {
            oForm.Items.Item("tab4").Visible = true;
            oForm.Items.Item("tab4").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
            oForm.PaneLevel = 4;
            //oForm.Freeze(true);
            fobROW = freightROW = InsuranceROW = schrateROW = ex10vrpkROW = finalvalROW = 0;

            SAPbouiCOM.Matrix matrix = oForm.Items.Item("matRLF1").Specific;
            SAPbouiCOM.Matrix matETSC1 = oForm.Items.Item("matETSC1").Specific;
            SAPbouiCOM.Matrix matETSC2 = oForm.Items.Item("matETSC2").Specific;
            SAPbouiCOM.Matrix matRLF2 = oForm.Items.Item("matRLF2").Specific;
            DeleteMatrixAll(matETSC1);
            DeleteMatrixAll(matETSC2);

            int i = 1;

            string getDocEntry = "SELECT  T2.U_EXIMLIC, T2.U_EXIMLIN, T5.U_schno,  T5.U_schname, SUM(T0.Quantity) AS 'Total_Qty', SUM(T0.LineTotal) AS 'TotalLC', ";
            getDocEntry = getDocEntry + " SUM(T0.Totalfrgn) AS 'TotalFC' , T4.U_rmqty, T4.U_rmLC,T4.U_rmFC , T5.U_schLEDE FROM INV1 AS T0 LEFT JOIN OINV AS T1 ON T0.DocEntry = T1.DocEntry ";
            getDocEntry = getDocEntry + " LEFT JOIN OITM AS T2 ON T0.ItemCode = T2.ItemCode LEFT JOIN dbo.[@XSM1] AS T4 ON T4.U_itemcode = T2.U_EXIMLIC ";
            getDocEntry = getDocEntry + " LEFT JOIN dbo.[@EXSM] AS T5 ON T5.Code = T4.Code WHERE T0.DocEntry = '" + DocEntry + "' AND ";
            getDocEntry = getDocEntry + " T1.CardCode = '" + BPcode + "' GROUP BY T2.U_EXIMLIC, T2.U_EXIMLIN , T5.U_schno, T5.U_schname , T4.U_rmqty, T4.U_rmLC, T4.U_rmFC , T5.U_schLEDE";

            SAPbobsCOM.Recordset rec = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            rec.DoQuery(getDocEntry);
            if (rec.RecordCount > 0)
            {
                while (!rec.EoF)
                {
                    matrix.AddRow();
                    (matrix.Columns.Item("#").Cells.Item(i).Specific).Value = Convert.ToString(matrix.RowCount);
                    (matrix.Columns.Item("ex4ic").Cells.Item(i).Specific).Value = rec.Fields.Item("U_EXIMLIC").Value;
                    (matrix.Columns.Item("ex4in").Cells.Item(i).Specific).Value = rec.Fields.Item("U_EXIMLIN").Value;
                    (matrix.Columns.Item("ex4ln").Cells.Item(i).Specific).Value = rec.Fields.Item("U_schno").Value;
                    (matrix.Columns.Item("ex4lv").Cells.Item(i).Specific).Value = rec.Fields.Item("U_schLEDE").Value;

                    (matrix.Columns.Item("ex4lrqty").Cells.Item(i).Specific).Value = rec.Fields.Item("U_rmqty").Value;
                    (matrix.Columns.Item("ex4lrafc").Cells.Item(i).Specific).Value = rec.Fields.Item("U_rmFC").Value;
                    (matrix.Columns.Item("ex4lraklc").Cells.Item(i).Specific).Value = rec.Fields.Item("U_rmLC").Value;

                    (matrix.Columns.Item("ex4lfqty").Cells.Item(i).Specific).Value = rec.Fields.Item("Total_Qty").Value;
                    (matrix.Columns.Item("ex4lffc").Cells.Item(i).Specific).Value = rec.Fields.Item("TotalFC").Value;
                    (matrix.Columns.Item("ex4lflc").Cells.Item(i).Specific).Value = rec.Fields.Item("TotalLC").Value;
                    i++;

                    int j = 1;
                    string Qry2 = "Select T1.U_iritemcd,T1.U_iritemnm, T1.U_irQtyPer, T1.U_irExQtP FROM dbo.[@EXSM] AS T0 LEFT JOIN dbo.[@XSM3] AS T1 ON T0.Code = T1.Code Where T0.U_schno = '" + rec.Fields.Item("U_schno").Value + "'";
                    SAPbobsCOM.Recordset rec2 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                    rec2.DoQuery(Qry2);
                    if (rec2.RecordCount > 0)
                    {
                        while (!rec2.EoF)
                        {
                            matRLF2.AddRow();
                            (matRLF2.Columns.Item("#").Cells.Item(j).Specific).Value = Convert.ToString(matRLF2.RowCount);
                            (matRLF2.Columns.Item("ex5ic").Cells.Item(j).Specific).Value = rec2.Fields.Item("U_iritemcd").Value;
                            (matRLF2.Columns.Item("ex5in").Cells.Item(j).Specific).Value = rec2.Fields.Item("U_iritemnm").Value;
                            (matRLF2.Columns.Item("ex5np").Cells.Item(j).Specific).Value = rec2.Fields.Item("U_irQtyPer").Value;
                            // double netwt = 22758 * (Convert.ToDouble(rec2.Fields.Item("U_irQtyPer").Value)) / 100;
                            double netwt = Convert.ToDouble(rec.Fields.Item("Total_Qty").Value) * (Convert.ToDouble(rec2.Fields.Item("U_irQtyPer").Value)) / 100;
                            (matRLF2.Columns.Item("ex5nw").Cells.Item(j).Specific).Value = netwt;
                            (matRLF2.Columns.Item("ex5exp").Cells.Item(j).Specific).Value = rec2.Fields.Item("U_irExQtP").Value;
                            (matRLF2.Columns.Item("ex5exw").Cells.Item(j).Specific).Value = netwt * Convert.ToDouble(rec2.Fields.Item("U_irExQtP").Value);

                            j++;
                            rec2.MoveNext();
                        }
                    }

                    rec.MoveNext();
                }
            }
            DeleteMatrixBlankRowRFL(matrix);
            DeleteMatrixBlankRowRFL2(matRLF2);
            ArrengeMatrixLineNum(matrix);
            ArrengeMatrixLineNum(matRLF2);

            //oForm.Freeze(false);
            // doAutoColSum(matrix, "ex10qt");
            //doAutoColSum(matrix, "ex10fofc");

        }
       
        public void fillDBKMatrix(SAPbouiCOM.Form oForm, string DocEntry, string BPcode)
        {
            oForm.Items.Item("tab5").Visible = true;
            oForm.Items.Item("tab5").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
            oForm.PaneLevel = 5;

            fobROW = freightROW = InsuranceROW = schrateROW = ex10vrpkROW = finalvalROW = 0;

            
            SAPbouiCOM.Matrix matRLF1 = oForm.Items.Item("matRLF1").Specific;
            DeleteMatrixAll(matRLF1);
            SAPbouiCOM.Matrix matRLF2 = oForm.Items.Item("matRLF2").Specific;
            DeleteMatrixAll(matRLF2);

            SAPbouiCOM.Matrix matrix = oForm.Items.Item("matETSC1").Specific; 
            int i = 1;
            string getDocEntry = "SELECT T0.DocEntry, T0.ItemCode, T0.Dscription, T0.Quantity, T0.LineTotal, T0.Totalfrgn, T1.DocNum, T3.ChapterID, T5.U_schno, T5.U_schname  , T5.U_schrate ,  T5.U_schratepk ";
            getDocEntry = getDocEntry + " ,T0.LineNum, T0.VisOrder FROM INV1 AS T0 LEFT JOIN OINV AS T1 ON T0.DocEntry = T1.DocEntry ";
            getDocEntry = getDocEntry + " LEFT JOIN OITM AS T2 ON T0.ItemCode = T2.ItemCode  LEFT JOIN OCHP AS T3 ON T2.ChapterID = T3.AbsEntry  ";
            getDocEntry = getDocEntry + " LEFT JOIN dbo.[@XSM5] AS T4 ON T4.U_hsncode = T3.ChapterID  LEFT JOIN dbo.[@EXSM] AS T5 ON T5.Code = T4.Code";
            getDocEntry = getDocEntry + " WHERE T0.DocEntry = '" + DocEntry + "' AND T1.CardCode='" + BPcode + "' AND T5.U_schtype = 'DBK' order by T0.VisOrder ASC";

            SAPbobsCOM.Recordset rec = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            rec.DoQuery(getDocEntry);
            if (rec.RecordCount > 0)
            {
                while (!rec.EoF)
                {
                    freightROW = InsuranceROW = ex10vrpkROW = schrateROW = fobROW = finalvalROW = 0;
                    matrix.AddRow();
                    (matrix.Columns.Item("#").Cells.Item(i).Specific).Value = Convert.ToString(i);
                    (matrix.Columns.Item("ex10ic").Cells.Item(i).Specific).Value = rec.Fields.Item("ItemCode").Value;
                    (matrix.Columns.Item("ex10in").Cells.Item(i).Specific).Value = rec.Fields.Item("Dscription").Value;
                    (matrix.Columns.Item("ex10sc").Cells.Item(i).Specific).Value = rec.Fields.Item("U_schno").Value;
                    (matrix.Columns.Item("ex10sn").Cells.Item(i).Specific).Value = rec.Fields.Item("U_schname").Value;
                    (matrix.Columns.Item("ex10qt").Cells.Item(i).Specific).Value = rec.Fields.Item("Quantity").Value;
                    (matrix.Columns.Item("ex10fofc").Cells.Item(i).Specific).Value = rec.Fields.Item("Totalfrgn").Value;
                    (matrix.Columns.Item("ex10folc").Cells.Item(i).Specific).Value = rec.Fields.Item("LineTotal").Value;

                    (matrix.Columns.Item("ex10frfc").Cells.Item(i).Specific).Value = 0;
                    (matrix.Columns.Item("ex10frlc").Cells.Item(i).Specific).Value = freightROW;

                    (matrix.Columns.Item("ex10infc").Cells.Item(i).Specific).Value = 0;
                    (matrix.Columns.Item("ex10inlc").Cells.Item(i).Specific).Value = InsuranceROW;

                    fobROW = rec.Fields.Item("LineTotal").Value - freightROW - InsuranceROW;
                    (matrix.Columns.Item("ex10fbv").Cells.Item(i).Specific).Value = fobROW;
                    (matrix.Columns.Item("ex10scp").Cells.Item(i).Specific).Value = rec.Fields.Item("U_schrate").Value;

                    schrateROW = fobROW * (Convert.ToDouble(rec.Fields.Item("U_schrate").Value) / 100);
                    (matrix.Columns.Item("ex10sv").Cells.Item(i).Specific).Value = schrateROW;

                    (matrix.Columns.Item("ex10srpk").Cells.Item(i).Specific).Value = rec.Fields.Item("U_schratepk").Value; //rec.Fields.Item("U_schname").Value;

                    ex10vrpkROW = Convert.ToDouble(rec.Fields.Item("Quantity").Value) * Convert.ToDouble(rec.Fields.Item("U_schratepk").Value);
                    (matrix.Columns.Item("ex10vrpk").Cells.Item(i).Specific).Value = ex10vrpkROW;

                    if (ex10vrpkROW > 0)
                    {
                        if (schrateROW <= ex10vrpkROW) { finalvalROW = schrateROW; } else { finalvalROW = ex10vrpkROW; }
                    }
                    else { finalvalROW = schrateROW; }

                    (matrix.Columns.Item("ex10fv").Cells.Item(i).Specific).Value = finalvalROW;

                    (matrix.Columns.Item("ex10li").Cells.Item(i).Specific).Value = rec.Fields.Item("LineNum").Value;
                    (matrix.Columns.Item("ex10vo").Cells.Item(i).Specific).Value = rec.Fields.Item("VisOrder").Value;

                    i++;
                    rec.MoveNext();
                }
            }
            DeleteMatrixBlankRowDBK(matrix);

            doAutoColSum(matrix, "ex10qt");
            doAutoColSum(matrix, "ex10fofc");
            doAutoColSum(matrix, "ex10folc");
            doAutoColSum(matrix, "ex10frfc");
            doAutoColSum(matrix, "ex10frlc");
            doAutoColSum(matrix, "ex10infc");
            doAutoColSum(matrix, "ex10inlc");
            doAutoColSum(matrix, "ex10fbv");
            doAutoColSum(matrix, "ex10sv");
            doAutoColSum(matrix, "ex10vrpk");
            doAutoColSum(matrix, "ex10fv");
        }
        
        public void fillRoDTEPMatrix(SAPbouiCOM.Form oForm, string DocEntry, string BPcode)
        {
            oForm.Items.Item("tab6").Visible = true;
            oForm.Items.Item("tab6").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
            oForm.PaneLevel = 6;

            SAPbouiCOM.Matrix matrix1 = oForm.Items.Item("matETSC2").Specific;
            int i = 1;
            string getDocEntry1 = "SELECT T0.DocEntry, T0.ItemCode, T0.Dscription, T0.Quantity, T0.LineTotal, T0.Totalfrgn, T1.DocNum, T3.ChapterID, T5.U_schno, T5.U_schname  , T5.U_schrate ,  T5.U_schratepk ";
            getDocEntry1 = getDocEntry1 + "  ,T0.LineNum, T0.VisOrder  FROM INV1 AS T0 LEFT JOIN OINV AS T1 ON T0.DocEntry = T1.DocEntry ";
            getDocEntry1 = getDocEntry1 + " LEFT JOIN OITM AS T2 ON T0.ItemCode = T2.ItemCode  LEFT JOIN OCHP AS T3 ON T2.ChapterID = T3.AbsEntry  ";
            getDocEntry1 = getDocEntry1 + " LEFT JOIN dbo.[@XSM5] AS T4 ON T4.U_hsncode = T3.ChapterID  LEFT JOIN dbo.[@EXSM] AS T5 ON T5.Code = T4.Code";
            getDocEntry1 = getDocEntry1 + " WHERE T0.DocEntry = '" + DocEntry + "' AND T1.CardCode='" + BPcode + "' AND T5.U_schtype = 'RoDTEP'  order by T0.VisOrder ASC";

            SAPbobsCOM.Recordset rec1 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            rec1.DoQuery(getDocEntry1);
            if (rec1.RecordCount > 0)
            {
                while (!rec1.EoF)
                {
                    freightROW = InsuranceROW = ex10vrpkROW = schrateROW = fobROW = finalvalROW = 0;
                    matrix1.AddRow();
                    (matrix1.Columns.Item("#").Cells.Item(i).Specific).Value = Convert.ToString(i);
                    (matrix1.Columns.Item("ex11ic").Cells.Item(i).Specific).Value = rec1.Fields.Item("ItemCode").Value;
                    (matrix1.Columns.Item("ex11in").Cells.Item(i).Specific).Value = rec1.Fields.Item("Dscription").Value;
                    (matrix1.Columns.Item("ex11sc").Cells.Item(i).Specific).Value = rec1.Fields.Item("U_schno").Value;
                    (matrix1.Columns.Item("ex11sn").Cells.Item(i).Specific).Value = rec1.Fields.Item("U_schname").Value;
                    (matrix1.Columns.Item("ex11qt").Cells.Item(i).Specific).Value = rec1.Fields.Item("Quantity").Value;
                    (matrix1.Columns.Item("ex11fofc").Cells.Item(i).Specific).Value = rec1.Fields.Item("Totalfrgn").Value;
                    (matrix1.Columns.Item("ex11folc").Cells.Item(i).Specific).Value = rec1.Fields.Item("LineTotal").Value;

                    (matrix1.Columns.Item("ex11frfc").Cells.Item(i).Specific).Value = 0;
                    (matrix1.Columns.Item("ex11frlc").Cells.Item(i).Specific).Value = freightROW;

                    (matrix1.Columns.Item("ex11infc").Cells.Item(i).Specific).Value = 0;
                    (matrix1.Columns.Item("ex11inlc").Cells.Item(i).Specific).Value = InsuranceROW;

                    fobROW = rec1.Fields.Item("LineTotal").Value - freightROW - InsuranceROW;
                    (matrix1.Columns.Item("ex11fbv").Cells.Item(i).Specific).Value = fobROW;
                    (matrix1.Columns.Item("ex11scp").Cells.Item(i).Specific).Value = rec1.Fields.Item("U_schrate").Value;

                    schrateROW = fobROW * (Convert.ToDouble(rec1.Fields.Item("U_schrate").Value) / 100);
                    (matrix1.Columns.Item("ex11sv").Cells.Item(i).Specific).Value = schrateROW;

                    (matrix1.Columns.Item("ex11srpk").Cells.Item(i).Specific).Value = rec1.Fields.Item("U_schratepk").Value;

                    ex10vrpkROW = Convert.ToDouble(rec1.Fields.Item("Quantity").Value) * Convert.ToDouble(rec1.Fields.Item("U_schratepk").Value);
                    (matrix1.Columns.Item("ex11vrpk").Cells.Item(i).Specific).Value = ex10vrpkROW;


                    if (ex10vrpkROW > 0)
                    {
                        if (schrateROW <= ex10vrpkROW) { finalvalROW = schrateROW; } else { finalvalROW = ex10vrpkROW; }
                    }
                    else { finalvalROW = schrateROW; }

                    (matrix1.Columns.Item("ex11fv").Cells.Item(i).Specific).Value = finalvalROW;

                    (matrix1.Columns.Item("ex11li").Cells.Item(i).Specific).Value = rec1.Fields.Item("LineNum").Value;
                    (matrix1.Columns.Item("ex11vo").Cells.Item(i).Specific).Value = rec1.Fields.Item("VisOrder").Value;

                    i++;
                    rec1.MoveNext();
                }
            }
            DeleteMatrixBlankRowRoDTEP(matrix1);
            doAutoColSum(matrix1, "ex11qt");
            doAutoColSum(matrix1, "ex11fofc");
            doAutoColSum(matrix1, "ex11folc");

            doAutoColSum(matrix1, "ex11frfc");
            doAutoColSum(matrix1, "ex11frlc");
            doAutoColSum(matrix1, "ex11infc");
            doAutoColSum(matrix1, "ex11inlc");

            doAutoColSum(matrix1, "ex11fbv");
            doAutoColSum(matrix1, "ex11sv");
            doAutoColSum(matrix1, "ex11vrpk");
            doAutoColSum(matrix1, "ex11fv");
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
                if (pVal.BeforeAction == true)
                {
                    if (Type == "REMOVE")
                    { 
                        String Q1 = " SELECT Count(*) as total FROM dbo.[@XSM2] where U_expexmno = '"+ oForm.Items.Item("tCode").Specific.Value + "'";
                        SAPbobsCOM.Recordset rec1 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                        rec1.DoQuery(Q1);
                        int cnt1 = Convert.ToInt32(rec1.Fields.Item("total").Value);

                        String Q2 = " SELECT Count(*) as total From dbo.[@XRU1] where U_schexno = '" + oForm.Items.Item("tCode").Specific.Value + "'";
                        SAPbobsCOM.Recordset rec2 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                        rec2.DoQuery(Q2);
                        int cnt2 = Convert.ToInt32(rec2.Fields.Item("total").Value);

                        String Q3 = "SELECT Count(*) as total From dbo.[@XRU2] where U_impexno = '" + oForm.Items.Item("tCode").Specific.Value + "'";
                        SAPbobsCOM.Recordset rec3 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                        rec3.DoQuery(Q3);
                        int cnt3 = Convert.ToInt32(rec3.Fields.Item("total").Value);

                        if (cnt1 > 0)
                        {
                            SBOMain.SBO_Application.StatusBar.SetText("Remove operation not allowed because this Exim transaction is assigned in Export Filfilment tab in scheme master", BoMessageTime.bmt_Long, BoStatusBarMessageType.smt_Error);
                            BubbleEvent = false;
                        }
                        else if(cnt2 > 0)
                        {
                            SBOMain.SBO_Application.StatusBar.SetText("Remove operation not allowed because this Exim transaction is assigned in Detail of Export Transaction tab in scheme transaction", BoMessageTime.bmt_Long, BoStatusBarMessageType.smt_Error);
                            BubbleEvent = false;
                        }
                        else if (cnt3 > 0)
                        {
                            SBOMain.SBO_Application.StatusBar.SetText("Remove operation not allowed because this Exim transaction is assigned in Detail of Import Transaction tab in scheme transaction", BoMessageTime.bmt_Long, BoStatusBarMessageType.smt_Error);
                            BubbleEvent = false;
                        } 
                    }
                } 
                 
                if (pVal.BeforeAction == false)
                {
                     
                    if ((oForm.Mode == BoFormMode.fm_ADD_MODE || oForm.Mode == BoFormMode.fm_UPDATE_MODE || Type == "ADDNEWFORM") && (Type != "DEL_ROW") && (Type != "ADD_ROW"))
                    {
                        Form_Load_Components(oForm, "ADD");
                    }
                    if (Type == "navigation")
                    {
                        changelabels(oForm);
                        doAutoSummatETSC1(oForm);
                    }

                    else if (Type == "DEL_ROW" || Type == "ADD_ROW")
                    {
                        SAPbouiCOM.Matrix matFOB = (SAPbouiCOM.Matrix)oForm.Items.Item("matFOB").Specific;
                        SAPbouiCOM.Matrix matEXPAC = (SAPbouiCOM.Matrix)oForm.Items.Item("matEXPAC").Specific;
                        SAPbouiCOM.Matrix matRLF1 = (SAPbouiCOM.Matrix)oForm.Items.Item("matRLF1").Specific;
                        SAPbouiCOM.Matrix matRLF2 = (SAPbouiCOM.Matrix)oForm.Items.Item("matRLF2").Specific;
                        SAPbouiCOM.Matrix matETSC1 = (SAPbouiCOM.Matrix)oForm.Items.Item("matETSC1").Specific;
                        SAPbouiCOM.Matrix matETSC2 = (SAPbouiCOM.Matrix)oForm.Items.Item("matETSC2").Specific;
                        SAPbouiCOM.Matrix matEXVGM = (SAPbouiCOM.Matrix)oForm.Items.Item("matEXVGM").Specific;
                        SAPbouiCOM.Matrix matEXCA = (SAPbouiCOM.Matrix)oForm.Items.Item("matEXCA").Specific;
                        SAPbouiCOM.Matrix matRS = (SAPbouiCOM.Matrix)oForm.Items.Item("matRS").Specific;

                        if (Type == "ADD_ROW")
                        {
                            if (SBOMain.RightClickItemID == "matFOB")
                            {
                                ADDROWMain(matFOB);
                            }
                            else if (SBOMain.RightClickItemID == "matEXPAC")
                            {
                                ADDROWMain(matEXPAC);
                            }
                            else if (SBOMain.RightClickItemID == "matRLF1")
                            {
                                ADDROWMain(matRLF1);
                            }
                            else if (SBOMain.RightClickItemID == "matRLF2")
                            {
                                ADDROWMain(matRLF2);
                            }
                            else if (SBOMain.RightClickItemID == "matETSC1")
                            {
                                ADDROWMain(matETSC1);
                            }
                            else if (SBOMain.RightClickItemID == "matETSC2")
                            {
                                ADDROWMain(matETSC2);
                            }
                            else if (SBOMain.RightClickItemID == "matEXVGM")
                            {
                                ADDROWMain(matEXVGM);
                            }
                            else if (SBOMain.RightClickItemID == "matEXCA")
                            {
                                ADDROWMain(matEXCA);
                            }
                            else if (SBOMain.RightClickItemID == "matRS")
                            {
                                ADDROWMain(matRS);
                            }
                        }
                        if (Type == "DEL_ROW")
                        {
                            if (SBOMain.RightClickItemID == "matFOB")
                            {
                                DeleteMatrixBlankRow(matFOB, "pdipn");
                                ArrengeMatrixLineNum(matFOB); 
                            }
                            else if (SBOMain.RightClickItemID == "matEXPAC")
                            {     
                                DeleteMatrixBlankRow(matEXPAC, "et3expt");
                                ArrengeMatrixLineNum(matEXPAC);
                            }
                            else if (SBOMain.RightClickItemID == "matRLF1")
                            {
                                DeleteMatrixBlankRow(matRLF1, "ex4ic");
                                ArrengeMatrixLineNum(matRLF1);
                             }
                            else if (SBOMain.RightClickItemID == "matRLF2")
                            {
                                DeleteMatrixBlankRow(matRLF2, "ex5ic");
                                ArrengeMatrixLineNum(matRLF2);
                             }
                            else if (SBOMain.RightClickItemID == "matETSC1")
                            {
                                DeleteMatrixBlankRow(matETSC1, "ex10ic");
                                ArrengeMatrixLineNum(matETSC1);
                             }
                            else if (SBOMain.RightClickItemID == "matETSC2")
                            {
                                DeleteMatrixBlankRow(matETSC2, "ex11ic");
                                ArrengeMatrixLineNum(matETSC2);
                             }
                            else if (SBOMain.RightClickItemID == "matEXVGM")
                            {
                                DeleteMatrixBlankRow(matEXVGM, "ex7cs");
                                ArrengeMatrixLineNum(matEXVGM);
                             }
                            else if (SBOMain.RightClickItemID == "matEXCA")
                            {
                                DeleteMatrixBlankRow(matEXCA, "ex8ic");
                                ArrengeMatrixLineNum(matEXCA);
                             }
                            else if (SBOMain.RightClickItemID == "matRS")
                            {
                                DeleteMatrixBlankRow(matRS, "ex13st");
                                ArrengeMatrixLineNum(matRS); 
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
        public void calculateFOBValue(string FOB, string CIF, string Freight, string Insurance, string othercharge, string FormId)
        {
            oForm = SBOMain.SBO_Application.Forms.Item(FormId);
            double CIFval = Convert.ToDouble(oForm.Items.Item(CIF).Specific.value);
            double FRval = Convert.ToDouble(oForm.Items.Item(Freight).Specific.value);
            double INval = Convert.ToDouble(oForm.Items.Item(Insurance).Specific.value);
            double OCval = Convert.ToDouble(oForm.Items.Item(othercharge).Specific.value);
           // oForm.Items.Item(FOB).Specific.value = (CIFval - FRval - INval - OCval).ToString();
            oForm.Items.Item(FOB).Specific.value = (CIFval - FRval - INval).ToString();
        }

        public void calculateRoDTEP(SAPbouiCOM.Form oForm)
        {

            oForm.Items.Item("tab6").Visible = true;
            oForm.Items.Item("tab6").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
            oForm.PaneLevel = 6;

            double freight = Convert.ToDouble(oForm.Items.Item("frtAFC").Specific.value);
            double Insurance = Convert.ToDouble(oForm.Items.Item("insAFC").Specific.value);
            double exchangeRate = Convert.ToDouble(oForm.Items.Item("exinvdt").Specific.value);
            double TOTALFOBFC = 0;

            double fobROW = 0;
            double freightROW = 0;
            double InsuranceROW = 0;
            double schrateROW = 0;
            double ex10vrpkROW = 0;
            double finalvalROW = 0;

            SAPbouiCOM.Matrix matrix = oForm.Items.Item("matETSC2").Specific;

            for (int i = 1; i <= matrix.RowCount; i++)
            {
                TOTALFOBFC = TOTALFOBFC + Convert.ToDouble(((SAPbouiCOM.EditText)matrix.Columns.Item("ex11fofc").Cells.Item(i).Specific).Value);
            }

            for (int i = 1; i <= matrix.RowCount; i++)
            {
                fobROW = 0;
                schrateROW = 0;

                freightROW = Convert.ToDouble(((SAPbouiCOM.EditText)matrix.Columns.Item("ex11fofc").Cells.Item(i).Specific).Value) * freight / TOTALFOBFC;
                (matrix.Columns.Item("ex11frfc").Cells.Item(i).Specific).Value = freightROW;

                freightROW = freightROW * exchangeRate;
                (matrix.Columns.Item("ex11frlc").Cells.Item(i).Specific).Value = freightROW;

                InsuranceROW = Convert.ToDouble(((SAPbouiCOM.EditText)matrix.Columns.Item("ex11fofc").Cells.Item(i).Specific).Value) * Insurance / TOTALFOBFC;
                (matrix.Columns.Item("ex11infc").Cells.Item(i).Specific).Value = InsuranceROW;

                InsuranceROW = InsuranceROW * exchangeRate;
                (matrix.Columns.Item("ex11inlc").Cells.Item(i).Specific).Value = InsuranceROW;

                fobROW = Convert.ToDouble(((SAPbouiCOM.EditText)matrix.Columns.Item("ex11folc").Cells.Item(i).Specific).Value) - freightROW - InsuranceROW;
                (matrix.Columns.Item("ex11fbv").Cells.Item(i).Specific).Value = fobROW;

                schrateROW = fobROW * (Convert.ToDouble(((SAPbouiCOM.EditText)matrix.Columns.Item("ex11scp").Cells.Item(i).Specific).Value) / 100);
                (matrix.Columns.Item("ex11sv").Cells.Item(i).Specific).Value = schrateROW;

                ex10vrpkROW = Convert.ToDouble(((SAPbouiCOM.EditText)matrix.Columns.Item("ex11qt").Cells.Item(i).Specific).Value) * Convert.ToDouble(((SAPbouiCOM.EditText)matrix.Columns.Item("ex11srpk").Cells.Item(i).Specific).Value);
                (matrix.Columns.Item("ex11vrpk").Cells.Item(i).Specific).Value = ex10vrpkROW;

                if (ex10vrpkROW != 0)
                {
                    if (schrateROW <= ex10vrpkROW) { finalvalROW = schrateROW; } else { finalvalROW = ex10vrpkROW; }
                }
                else { finalvalROW = schrateROW; }

                    (matrix.Columns.Item("ex11fv").Cells.Item(i).Specific).Value = finalvalROW;

            }
            doAutoColSum(matrix, "ex11qt");
            doAutoColSum(matrix, "ex11fofc");
            doAutoColSum(matrix, "ex11folc");
            doAutoColSum(matrix, "ex11fbv");
            doAutoColSum(matrix, "ex11sv");
            doAutoColSum(matrix, "ex11vrpk");
            doAutoColSum(matrix, "ex11fv");
            doAutoColSum(matrix, "ex11frfc");
            doAutoColSum(matrix, "ex11frlc");
            doAutoColSum(matrix, "ex11infc");
            doAutoColSum(matrix, "ex11inlc");
        }

        public void calculateRLF(SAPbouiCOM.Form oForm)
        {

            oForm.Items.Item("tab4").Visible = true;
            oForm.Items.Item("tab4").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
            oForm.PaneLevel = 4;

            double fobAFC = Convert.ToDouble(oForm.Items.Item("fobAFC").Specific.value);
            double fobALC = Convert.ToDouble(oForm.Items.Item("fobALC").Specific.value);
            double frtAFC = Convert.ToDouble(oForm.Items.Item("frtAFC").Specific.value);
            double frtALC = Convert.ToDouble(oForm.Items.Item("frtALC").Specific.value);
            double insAFC = Convert.ToDouble(oForm.Items.Item("insAFC").Specific.value);
            double insALC = Convert.ToDouble(oForm.Items.Item("insALC").Specific.value);

            double ex4lrqty = 0;

            SAPbouiCOM.Matrix matrix = oForm.Items.Item("matRLF1").Specific;
            
            if(matrix.RowCount == 1 || ( matrix.RowCount == 2 && (string.IsNullOrEmpty((matrix.Columns.Item("ex4ic").Cells.Item(2).Specific).Value))))
            {
                (matrix.Columns.Item("ex4lffc").Cells.Item(1).Specific).Value = fobAFC;
                (matrix.Columns.Item("ex4lflc").Cells.Item(1).Specific).Value = fobALC;
            }
            else
            {
                double qty = 0;
                double per = 0;
                double fc = 0;
                double lc = 0;
                double nfc = 0;
                double nlc = 0;
                double infc = 0;
                double inlc = 0;

                String getDocEntry1 = "SELECT sum(T1.Quantity) as totalQty from OINV AS T0 Left JOIN INV1 AS T1 on T0.DocEntry = T1.DocEntry AND T0.DocEntry = '"+ oForm.Items.Item("exinvnode").Specific.Value + "'";
                SAPbobsCOM.Recordset rec1 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                rec1.DoQuery(getDocEntry1);
                if (rec1.RecordCount > 0)
                {
                    ex4lrqty = rec1.Fields.Item("totalQty").Value;
                }

                for (int i = 1; i <= matrix.RowCount; i++)
                {                     
                    qty = Convert.ToDouble(((SAPbouiCOM.EditText)matrix.Columns.Item("ex4lfqty").Cells.Item(i).Specific).Value);
                    per = qty * 100 / ex4lrqty;

                    fc = Convert.ToDouble(((SAPbouiCOM.EditText)matrix.Columns.Item("ex4lffc").Cells.Item(i).Specific).Value);
                    lc = Convert.ToDouble(((SAPbouiCOM.EditText)matrix.Columns.Item("ex4lflc").Cells.Item(i).Specific).Value);
                    nfc = frtAFC * per / 100;
                    nlc = frtALC * per / 100;
                    infc = insAFC * per / 100;
                    inlc = insALC * per / 100;
                    (matrix.Columns.Item("ex4lffc").Cells.Item(i).Specific).Value = fc - nfc - infc;
                    (matrix.Columns.Item("ex4lflc").Cells.Item(i).Specific).Value = lc - nlc - inlc; 
                }
                    
            }
            doAutoSummatRLF1(oForm);            

        }

        public void calculateDBK(SAPbouiCOM.Form oForm)
        {

            oForm.Items.Item("tab5").Visible = true;
            oForm.Items.Item("tab5").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
            oForm.PaneLevel = 5;

            freight = Convert.ToDouble(oForm.Items.Item("frtAFC").Specific.value);
            insurance = Convert.ToDouble(oForm.Items.Item("insAFC").Specific.value);
            exchangeRate = Convert.ToDouble(oForm.Items.Item("exinvdt").Specific.value);
            double TOTALFOBFC = 0;
            fobROW = freightROW = InsuranceROW = schrateROW = ex10vrpkROW = finalvalROW = 0;


            SAPbouiCOM.Matrix matrix = oForm.Items.Item("matETSC1").Specific;

            for (int i = 1; i <= matrix.RowCount; i++)
            {
                TOTALFOBFC = TOTALFOBFC + Convert.ToDouble(((SAPbouiCOM.EditText)matrix.Columns.Item("ex10fofc").Cells.Item(i).Specific).Value);
            }

            for (int i = 1; i <= matrix.RowCount; i++)
            {
                fobROW = 0;
                schrateROW = 0;

                freightROW = Convert.ToDouble(((SAPbouiCOM.EditText)matrix.Columns.Item("ex10fofc").Cells.Item(i).Specific).Value) * freight / TOTALFOBFC;
                (matrix.Columns.Item("ex10frfc").Cells.Item(i).Specific).Value = freightROW;

                freightROW = freightROW * exchangeRate;
                (matrix.Columns.Item("ex10frlc").Cells.Item(i).Specific).Value = freightROW;

                InsuranceROW = Convert.ToDouble(((SAPbouiCOM.EditText)matrix.Columns.Item("ex10fofc").Cells.Item(i).Specific).Value) * insurance / TOTALFOBFC;
                (matrix.Columns.Item("ex10infc").Cells.Item(i).Specific).Value = InsuranceROW;

                InsuranceROW = InsuranceROW * exchangeRate;
                (matrix.Columns.Item("ex10inlc").Cells.Item(i).Specific).Value = InsuranceROW;

                fobROW = Convert.ToDouble(((SAPbouiCOM.EditText)matrix.Columns.Item("ex10folc").Cells.Item(i).Specific).Value) - freightROW - InsuranceROW;
                (matrix.Columns.Item("ex10fbv").Cells.Item(i).Specific).Value = fobROW;

                schrateROW = fobROW * (Convert.ToDouble(((SAPbouiCOM.EditText)matrix.Columns.Item("ex10scp").Cells.Item(i).Specific).Value) / 100);
                (matrix.Columns.Item("ex10sv").Cells.Item(i).Specific).Value = schrateROW;

                ex10vrpkROW = Convert.ToDouble(((SAPbouiCOM.EditText)matrix.Columns.Item("ex10qt").Cells.Item(i).Specific).Value) * Convert.ToDouble(((SAPbouiCOM.EditText)matrix.Columns.Item("ex10srpk").Cells.Item(i).Specific).Value);
                (matrix.Columns.Item("ex10vrpk").Cells.Item(i).Specific).Value = ex10vrpkROW;

                if (ex10vrpkROW != 0)
                {
                    if (schrateROW <= ex10vrpkROW) { finalvalROW = schrateROW; } else { finalvalROW = ex10vrpkROW; }
                }
                else { finalvalROW = schrateROW; }

                if (schrateROW <= ex10vrpkROW) { finalvalROW = schrateROW; } else { finalvalROW = ex10vrpkROW; }
                    (matrix.Columns.Item("ex10fv").Cells.Item(i).Specific).Value = finalvalROW;

            }

            doAutoColSum(matrix, "ex10qt");
            doAutoColSum(matrix, "ex10fofc");
            doAutoColSum(matrix, "ex10folc");
            doAutoColSum(matrix, "ex10fbv");
            doAutoColSum(matrix, "ex10sv");
            doAutoColSum(matrix, "ex10vrpk");
            doAutoColSum(matrix, "ex10fv");
            doAutoColSum(matrix, "ex10frfc");
            doAutoColSum(matrix, "ex10frlc");
            doAutoColSum(matrix, "ex10infc");
            doAutoColSum(matrix, "ex10inlc");

        }
        public bool ItemEvent(ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent, string FormId)
        {
            BubbleEvent = true;
            oForm = SBOMain.SBO_Application.Forms.Item(FormId);
            try
            {
                switch (pVal.EventType)
                { // KEY Down event for multiplication

                    case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
                        if (pVal.BeforeAction == true)
                        {

                        }
                        if (pVal.BeforeAction == false)
                        {
                            if ((pVal.ItemUID == "matEXPAC" && pVal.ColUID == "et3expt"))
                            {
                                string abc = null;
                                 
                                    SAPbouiCOM.Matrix matrix = oForm.Items.Item("matEXPAC").Specific;
                                    abc = (matrix.Columns.Item("et3expt").Cells.Item(pVal.Row).Specific).Value;
                                
                                if (!string.IsNullOrEmpty(abc))
                                {
                                    bool plFormOpen = false;
                                    for (int i = 1; i < SBOMain.SBO_Application.Forms.Count; i++)
                                    {
                                        if (SBOMain.SBO_Application.Forms.Item(i).UniqueID == "frmExpMaster")
                                        {
                                            SBOMain.SBO_Application.Forms.Item(i).Select();
                                            plFormOpen = true;
                                        }
                                    }
                                    if (!plFormOpen)
                                    {
                                        SBOMain.SBO_Application.Menus.Item("mnsmEXIM010").Activate();
                                    }

                                    OutwardToEXPMaster data = new OutwardToEXPMaster();
                                    data.itemcode= abc;
                                    clsExpMaster oPrice = new clsExpMaster(data);
                                    //oForm.Close();
                                }
                            }

                            if ((pVal.ItemUID == "matRLF1" && pVal.ColUID == "ex4ln") || (pVal.ItemUID == "matRLF2" && pVal.ColUID == "ex5ln"))
                            {
                                string abc = null;
                                if (pVal.ItemUID == "matRLF1" && pVal.ColUID == "ex4ln")
                                {
                                    SAPbouiCOM.Matrix matrix = oForm.Items.Item("matRLF1").Specific;
                                    abc = (matrix.Columns.Item("ex4ln").Cells.Item(pVal.Row).Specific).Value;
                                }
                                else if (pVal.ItemUID == "matRLF2" && pVal.ColUID == "ex5ln")
                                {
                                    SAPbouiCOM.Matrix matrix = oForm.Items.Item("matRLF2").Specific;
                                    abc = (matrix.Columns.Item("ex5ln").Cells.Item(pVal.Row).Specific).Value;
                                }

                                if (!string.IsNullOrEmpty(abc))
                                {
                                    bool plFormOpen = false;
                                    for (int i = 1; i < SBOMain.SBO_Application.Forms.Count; i++)
                                    {
                                        if (SBOMain.SBO_Application.Forms.Item(i).UniqueID == "frmSchmMaster")
                                        {
                                            SBOMain.SBO_Application.Forms.Item(i).Select();
                                            plFormOpen = true;
                                        }
                                    }
                                    if (!plFormOpen)
                                    {
                                        SBOMain.SBO_Application.Menus.Item("mnsmEXIM011").Activate();
                                    }
                                    OutwardToSchemeMaster OutwardToSchemeMaster = new OutwardToSchemeMaster();
                                    OutwardToSchemeMaster.schmeno = abc;
                                    clsSchmMaster oPrice = new clsSchmMaster(OutwardToSchemeMaster);
                                    //oForm.Close();
                                }
                            }

                            if ((pVal.ItemUID == "matEXPAC" && pVal.ColUID == "et3ede"))
                            {
                                string DocEntry = null;
                                string DocNum = null;
                                string CardCode = null;

                                SAPbouiCOM.Matrix matrix = oForm.Items.Item("matEXPAC").Specific;
                                DocEntry = (matrix.Columns.Item("et3ede").Cells.Item(pVal.Row).Specific).Value;
                                DocNum = (matrix.Columns.Item("et3edn").Cells.Item(pVal.Row).Specific).Value;

                                SAPbouiCOM.ComboBox cb4 = matrix.Columns.Item("et3yc").Cells.Item(pVal.Row).Specific; // oForm.Items.Item("cmbCRCY").Specific;
                                string cmbcal = cb4.Selected.Value.ToString();

                                if (cmbcal == "PR")
                                {
                                    bool plFormOpen = false;
                                    for (int i = 1; i < SBOMain.SBO_Application.Forms.Count; i++)
                                    {
                                        if (SBOMain.SBO_Application.Forms.Item(i).UniqueID == "39724")
                                        {
                                            SBOMain.SBO_Application.Forms.Item(i).Select();
                                            plFormOpen = true;
                                        }
                                    }
                                    if (!plFormOpen)
                                    {
                                        SBOMain.SBO_Application.Menus.Item("39724").Activate();
                                        oForm = SBOMain.SBO_Application.Forms.ActiveForm;
                                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                                        oForm.Items.Item("8").Specific.value = DocNum;
                                        oForm.Items.Item("1").Click();
                                    }
                                }
                                else if (cmbcal == "PQ")
                                {
                                    bool plFormOpen = false;
                                    for (int i = 1; i < SBOMain.SBO_Application.Forms.Count; i++)
                                    {
                                        if (SBOMain.SBO_Application.Forms.Item(i).UniqueID == "39698")
                                        {
                                            SBOMain.SBO_Application.Forms.Item(i).Select();
                                            plFormOpen = true;
                                        }
                                    }
                                    if (!plFormOpen)
                                    {
                                        SBOMain.SBO_Application.Menus.Item("39698").Activate();
                                        oForm = SBOMain.SBO_Application.Forms.ActiveForm;
                                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                                        oForm.Items.Item("8").Specific.value = DocNum;
                                        oForm.Items.Item("1").Click();
                                    }
                                }
                                else if (cmbcal == "PO")
                                {
                                    bool plFormOpen = false;
                                    for (int i = 1; i < SBOMain.SBO_Application.Forms.Count; i++)
                                    {
                                        if (SBOMain.SBO_Application.Forms.Item(i).UniqueID == "2305")
                                        {
                                            SBOMain.SBO_Application.Forms.Item(i).Select();
                                            plFormOpen = true;
                                        }
                                    }
                                    if (!plFormOpen)
                                    {
                                        SBOMain.SBO_Application.Menus.Item("2305").Activate();
                                        oForm = SBOMain.SBO_Application.Forms.ActiveForm;
                                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                                        oForm.Items.Item("8").Specific.value = DocNum;
                                        oForm.Items.Item("1").Click();
                                    }
                                }
                                else if (cmbcal == "PI")
                                {
                                    bool plFormOpen = false;
                                    for (int i = 1; i < SBOMain.SBO_Application.Forms.Count; i++)
                                    {
                                        if (SBOMain.SBO_Application.Forms.Item(i).UniqueID == "2053")
                                        {
                                            SBOMain.SBO_Application.Forms.Item(i).Select();
                                            plFormOpen = true;
                                        }
                                    }
                                    if (!plFormOpen)
                                    {
                                        SBOMain.SBO_Application.Menus.Item("2053").Activate();
                                        oForm = SBOMain.SBO_Application.Forms.ActiveForm;
                                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                                        oForm.Items.Item("8").Specific.value = DocNum;
                                        oForm.Items.Item("1").Click();
                                    }
                                }
                            }
                        }

                        break;

                    case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
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
                                double freight = 0;
                                double insurance = 0;
                                double othercharge = 0;
                                double exchangeRate = 0;

                                if (pVal.ItemUID == "frtDFC" || pVal.ItemUID == "insDFC" || pVal.ItemUID == "ocDFC")
                                {
                                    exchangeRate = Convert.ToDouble(oForm.Items.Item("exinvdt").Specific.value);

                                    if (pVal.ItemUID == "frtDFC")
                                    {
                                        freight = Convert.ToDouble(oForm.Items.Item("frtDFC").Specific.value);
                                        oForm.Items.Item("frtDLC").Specific.value = (freight * exchangeRate).ToString();
                                    }
                                    if (pVal.ItemUID == "insDFC")
                                    {
                                        insurance = Convert.ToDouble(oForm.Items.Item("insDFC").Specific.value);
                                        oForm.Items.Item("insDLC").Specific.value = (insurance * exchangeRate).ToString();
                                    }
                                    if (pVal.ItemUID == "ocDFC")
                                    {
                                        othercharge = Convert.ToDouble(oForm.Items.Item("ocDFC").Specific.value);
                                        oForm.Items.Item("ocDLC").Specific.value = (othercharge * exchangeRate).ToString();
                                    }
                                    calculateFOBValue("fobDFC", "cifDFC", "frtDFC", "insDFC", "ocDFC", FormId);
                                    calculateFOBValue("fobDLC", "cifDLC", "frtDLC", "insDLC", "ocDLC", FormId);
                                }
                                if (pVal.ItemUID == "frtAFC" || pVal.ItemUID == "insAFC" || pVal.ItemUID == "ocAFC")
                                {
                                    exchangeRate = Convert.ToDouble(oForm.Items.Item("exinvdt").Specific.value);
                                    string statusval = oForm.Items.Item("exls").Specific.value.ToString();

                                    if (pVal.ItemUID == "frtAFC")
                                    {
                                        freight = Convert.ToDouble(oForm.Items.Item("frtAFC").Specific.value);
                                        oForm.Items.Item("frtALC").Specific.value = (freight * exchangeRate).ToString();
                                    }
                                    if (pVal.ItemUID == "insAFC")
                                    {
                                        insurance = Convert.ToDouble(oForm.Items.Item("insAFC").Specific.value);
                                        oForm.Items.Item("insALC").Specific.value = (insurance * exchangeRate).ToString();
                                        
                                    }
                                    if (pVal.ItemUID == "ocAFC")
                                    {
                                        othercharge = Convert.ToDouble(oForm.Items.Item("ocAFC").Specific.value);
                                        oForm.Items.Item("ocALC").Specific.value = (othercharge * exchangeRate).ToString(); 
                                    } 
                                    calculateFOBValue("fobAFC", "cifAFC", "frtAFC", "insAFC", "ocAFC", FormId);
                                    calculateFOBValue("fobALC", "cifALC", "frtALC", "insALC", "ocALC", FormId);

                                    if (pVal.ItemUID == "frtAFC" || pVal.ItemUID == "insAFC") {
                                        if (statusval == "Licence") {
                                            calculateRLF(oForm);
                                        }
                                        else if (statusval != "Licence"){
                                            calculateDBK(oForm);
                                            calculateRoDTEP(oForm);
                                        }
                                    }  
                                }
                                if (pVal.ItemUID == "matFOB" && pVal.ColUID == "pdipn")
                                {
                                    SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("matFOB").Specific;
                                    AddMatrixRow(oMatrix, "pdipn");
                                }

                                else if (pVal.ItemUID == "matRS" && pVal.ColUID == "ex13sn")
                                {
                                    SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("matRS").Specific;
                                    string scritno = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("ex13sn").Cells.Item(pVal.Row).Specific).Value;

                                    string getDocEntry = " SELECT T0.U_schup  FROM dbo.[@EXRU] AS T0  LEFT JOIN dbo.[@XRU1] as T1 ON T0.DocEntry = T1.DocEntry where U_schtype = 'RoDTEP' and U_schsn  = '" + scritno + "'";
                                    SAPbobsCOM.Recordset rec = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                    rec.DoQuery(getDocEntry);
                                    if (rec.RecordCount > 0)
                                    {
                                        ((SAPbouiCOM.EditText)oMatrix.Columns.Item("ex13up").Cells.Item(pVal.Row).Specific).Value = Convert.ToString(rec.Fields.Item("U_schup").Value);
                                        double utiledamt = (rec.Fields.Item("U_schup").Value * Convert.ToDouble((oMatrix.Columns.Item("ex13bv").Cells.Item(pVal.Row).Specific).Value)) / 100;
                                        ((SAPbouiCOM.EditText)oMatrix.Columns.Item("ex13sua").Cells.Item(pVal.Row).Specific).Value = utiledamt.ToString();
                                    }
                                }

                                else if (pVal.ItemUID == "matEXPAC" && pVal.ColUID == "et3yc")
                                {
                                    SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("matEXPAC").Specific;
                                    AddMatrixRow(oMatrix, "et3yc");
                                }
                                /*else if (pVal.ItemUID == "matEXPAC" && pVal.ColUID == "et3expt")
                                {
                                    SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("matEXPAC").Specific;
                                    expcode = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("et3expt").Cells.Item(pVal.Row).Specific).Value;
                                     
                                    if (string.IsNullOrEmpty(expcode))
                                    {
                                        AddMatrixRow(oMatrix, "et3expt");
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
                                            Program.ExExpData.EXExpMatCol = "et3expt";
                                            SBOMain.LoadFromXML("frmExpList", "EXIM");
                                            SBOMain.SBO_Application.Forms.Item("frmExpList").Select();

                                            var oForm1 = SBOMain.SBO_Application.Forms.ActiveForm;
                                            oForm1.DataSources.DataTables.Add("tab");
                                            SAPbouiCOM.Grid objGrid = oForm1.Items.Item("gridEXP").Specific;

                                            string Qry = "SELECT U_expcode as 'Expense Code',U_expname as 'Expense Name', U_expadname  as 'Expense Additional Name'FROM dbo.[@EXEM] where U_status = 1";
                                            oForm1.DataSources.DataTables.Item("tab").ExecuteQuery(Qry);
                                            objGrid.DataTable = oForm1.DataSources.DataTables.Item("tab");
                                        }
                                    }
                                }*/
                                else if (pVal.ItemUID == "matRLF1" && pVal.ColUID == "ex4ic")
                                {
                                    SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("matRLF1").Specific;
                                    AddMatrixRow(oMatrix, "ex4ic");
                                }
                                else if (pVal.ItemUID == "matRLF2" && pVal.ColUID == "ex5ic")
                                {
                                    SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("matRLF2").Specific;
                                    AddMatrixRow(oMatrix, "ex5ic");
                                }
                                else if (pVal.ItemUID == "matETSC1" && pVal.ColUID == "ex10sc")
                                {
                                    SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("matETSC1").Specific;
                                    // AddMatrixRow(oMatrix, "ex10sc");
                                }
                                else if (pVal.ItemUID == "matETSC2" && pVal.ColUID == "ex11sc")
                                {
                                    SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("matETSC2").Specific;
                                    // AddMatrixRow(oMatrix, "ex11sc");
                                }
                                else if (pVal.ItemUID == "matEXVGM" && pVal.ColUID == "ex7cs")
                                {
                                    SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("matEXVGM").Specific;
                                    AddMatrixRow(oMatrix, "ex7cs");
                                }
                                else if (pVal.ItemUID == "matEXCA" && pVal.ColUID == "ex8cn")
                                {
                                    SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("matEXCA").Specific;
                                    AddMatrixRow(oMatrix, "ex8cn");
                                }
                            }
                        }
                        break;
                    case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
                        if (pVal.BeforeAction == true)
                        {
                            if (pVal.ItemUID == "matRLF1" && pVal.ColUID == "ex4ln")
                            {
                                CFLConditionLicense("CFL_16", pVal.ItemUID);
                            }
                            if (pVal.ItemUID == "matEXPAC" && pVal.ColUID == "et3expt")
                            {
                                CFLConditionEXPType("CFL_18", pVal.ItemUID);
                            }
                            if (pVal.ItemUID == "matRLF1" && pVal.ColUID == "ex4ic")
                            {
                                CFLConditionLICITEM("CFLOITM2", pVal.ItemUID);
                            }
                            if (pVal.ItemUID == "exinco")
                            {
                                CFLConditionIN("CFL_22", pVal.ItemUID);
                            }
                            if (pVal.ItemUID == "exlcno")
                            {
                                CFLConditionLC("CFL_21", pVal.ItemUID);
                            }
                            if (pVal.ItemUID == "ex1pol")
                            { 
                                CFLConditionPort("CFL_17", pVal.ItemUID);
                            }
                            if (pVal.ItemUID == "ex1pod") 
                            { 
                                CFLConditionPort("CFL_19", pVal.ItemUID);
                            }
                            if(pVal.ItemUID == "ex1por"){

                                CFLConditionPort("CFL_20", pVal.ItemUID);
                            }

                            if (pVal.ItemUID == "exbc")
                            {
                                cb4 = oForm.Items.Item("exdt").Specific;
                                DocType = cb4.Selected.Value.ToString();
                                CFLConditionBP("CFL_OCRD", "exbc", DocType);
                            }

                            if (pVal.ItemUID == "fcCHAcd" || pVal.ItemUID == "ex1slc" || pVal.ItemUID == "twbwc")
                            {
                                NameValueCollection list1 = new NameValueCollection() { { "fcCHAcd", "CFL_OCRD1" }, { "ex1slc", "CFL_OCRD2" }, { "twbwc", "CFL_OCRD3" } };
                                CFLCondition(list1[pVal.ItemUID], pVal.ItemUID);
                            }
                            if (pVal.ItemUID == "exson" || pVal.ItemUID == "exdn" || pVal.ItemUID == "exinvno")
                            {
                                string bpcode = oForm.Items.Item("exbc").Specific.value;
                                string order = oForm.Items.Item("exsonde").Specific.value;
                                string delivery = oForm.Items.Item("exdnde").Specific.value;
                                string arinvoice = oForm.Items.Item("exinvnode").Specific.value;
                                //  string arinvoice = oForm.Items.Item("exinvnode").Specific.value;

                                cb4 = oForm.Items.Item("exdt").Specific;
                                DocType = cb4.Selected.Value.ToString();
                                if (DocType == "E")
                                {
                                    if (pVal.ItemUID == "exson")
                                    {
                                        CFLConditionTransORDR("CFL_ORDR", "exson", bpcode, order, delivery, arinvoice);
                                    }
                                    else if (pVal.ItemUID == "exdn")
                                    {
                                        CFLConditionTransODLN("CFL_ODLN", "exdn", bpcode, order, delivery, arinvoice);
                                    }
                                    else if (pVal.ItemUID == "exinvno")
                                    {
                                        CFLConditionTransOINV("CFL_OINV", "exinvno", bpcode, order, delivery, arinvoice);
                                    }

                                    /* NameValueCollection list2 = new NameValueCollection() { { "exson", "CFL_ORDR" }, { "exdn", "CFL_ODLN" }, { "exinvno", "CFL_OINV" } };
                                     CFLConditionTrans(list2[pVal.ItemUID], pVal.ItemUID, bpcode, order, delivery, arinvoice);*/
                                }
                                else
                                {
                                    if (pVal.ItemUID == "exson")
                                    {
                                        CFLConditionTransOPOR("CFL_OPOR", "exson", bpcode, order, delivery, arinvoice);
                                    }
                                    else if (pVal.ItemUID == "exdn")
                                    {
                                        CFLConditionTransOPDN("CFL_OPDN", "exdn", bpcode, order, delivery, arinvoice);
                                    }
                                    else if (pVal.ItemUID == "exinvno")
                                    {
                                        CFLConditionTransOPCH("CFL_OPCH", "exinvno", bpcode, order, delivery, arinvoice);
                                    }

                                    // NameValueCollection list2 = new NameValueCollection() { { "exson", "CFL_OPOR" }, { "exdn", "CFL_OPDN" }, { "exinvno", "CFL_OPCH" } };
                                    //CFLConditionTrans(list2[pVal.ItemUID], pVal.ItemUID, bpcode, order, delivery, arinvoice);
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
                                    if (pVal.ItemUID == "ex1fd")
                                    {
                                        oForm.Items.Item("ex1fd").Specific.value = oDataTable.GetValue("Name", 0).ToString();
                                    }
                                    else if (pVal.ItemUID == "ex1coo")
                                    {
                                        oForm.Items.Item("ex1coo").Specific.value = oDataTable.GetValue("Name", 0).ToString();
                                    }
                                   else if (pVal.ItemUID == "ex1dc")
                                    {
                                        oForm.Items.Item("ex1dc").Specific.value = oDataTable.GetValue("Name", 0).ToString();
                                    } 
                                   else if (pVal.ItemUID == "exinco")
                                    {
                                        oForm.Items.Item("exinco").Specific.value = oDataTable.GetValue("U_inctcode", 0).ToString();
                                    }
                                    else if (pVal.ItemUID == "exlcno")
                                    {
                                        oForm.Items.Item("exlcno").Specific.value = oDataTable.GetValue("U_lcln", 0).ToString(); 
                                        DateTime lcsd = Convert.ToDateTime(oDataTable.GetValue("U_lcod", 0).ToString()); 
                                        oForm.Items.Item("exlcdt").Specific.value = lcsd.ToString("yyyyMMdd");
                                    }
                                    else if (pVal.ItemUID == "ex1pol")
                                    {
                                        oForm.Items.Item("ex1pol").Specific.value = oDataTable.GetValue("U_portcode", 0).ToString();
                                    }
                                    else if (pVal.ItemUID == "ex1pod")
                                    {
                                        oForm.Items.Item("ex1pod").Specific.value = oDataTable.GetValue("U_portcode", 0).ToString();
                                    }
                                    else if (pVal.ItemUID == "ex1por")
                                    {
                                        oForm.Items.Item("ex1por").Specific.value = oDataTable.GetValue("U_portcode", 0).ToString();
                                    } 
                                    else if (pVal.ItemUID == "matEXPAC" && pVal.ColUID == "et3expt")
                                    {
                                        SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("matEXPAC").Specific;
                                        ((SAPbouiCOM.EditText)oMatrix.Columns.Item("et3expt").Cells.Item(pVal.Row).Specific).Value = oDataTable.GetValue("U_expcode", 0).ToString();
                                        ((SAPbouiCOM.EditText)oMatrix.Columns.Item("et3expnm").Cells.Item(pVal.Row).Specific).Value = oDataTable.GetValue("U_expname", 0).ToString();
                                        AddMatrixRow(oMatrix, "et3expt");
                                    }
                                    else  if (pVal.ItemUID == "exson")
                                    {
                                        oForm.Items.Item("exson").Specific.value = oDataTable.GetValue("DocNum", 0).ToString();
                                        oForm.Items.Item("exsonde").Specific.value = oDataTable.GetValue("DocEntry", 0).ToString();
                                    } 
                                    else if (pVal.ItemUID == "matRLF1" && pVal.ColUID == "ex4ln")
                                    {
                                        SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("matRLF1").Specific;

                                        ((SAPbouiCOM.EditText)oMatrix.Columns.Item("ex4ln").Cells.Item(pVal.Row).Specific).Value = oDataTable.GetValue("U_schno", 0).ToString();

                                        string itemCode = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("ex4ic").Cells.Item(pVal.Row).Specific).Value;
                                        string ExpCode = oDataTable.GetValue("U_schno", 0).ToString();

                                        oForm.Freeze(true);
                                        string getDocEntry = "SELECT T0.U_schLEDE , T1.U_rmqty, T1.U_rmLC, T1.U_rmFC FROM dbo.[@EXSM] as T0 ";
                                        getDocEntry = getDocEntry + " LEFT JOIN dbo.[@XSM1] AS T1 ON T0.Code = T1.Code ";
                                        getDocEntry = getDocEntry + " Where T0.U_schno = '" + ExpCode + "' AND T1.U_itemcode = '" + itemCode + "' ";
                                        SAPbobsCOM.Recordset rec = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                        rec.DoQuery(getDocEntry);
                                        if (rec.RecordCount > 0)
                                        {
                                            DateTime lcsd = Convert.ToDateTime(rec.Fields.Item("U_schLEDE").Value);
                                            string abc = lcsd.ToString("yyyyMMdd");

                                            ((SAPbouiCOM.EditText)oMatrix.Columns.Item("ex4lv").Cells.Item(pVal.Row).Specific).Value = abc;
                                            ((SAPbouiCOM.EditText)oMatrix.Columns.Item("ex4lrqty").Cells.Item(pVal.Row).Specific).Value = Convert.ToString(rec.Fields.Item("U_rmqty").Value);
                                            ((SAPbouiCOM.EditText)oMatrix.Columns.Item("ex4lrafc").Cells.Item(pVal.Row).Specific).Value = Convert.ToString(rec.Fields.Item("U_rmFC").Value);
                                            ((SAPbouiCOM.EditText)oMatrix.Columns.Item("ex4lraklc").Cells.Item(pVal.Row).Specific).Value = Convert.ToString(rec.Fields.Item("U_rmLC").Value);

                                            double qty = Convert.ToDouble((oMatrix.Columns.Item("ex4lfqty").Cells.Item(pVal.Row).Specific).Value);
                                             
                                            SAPbouiCOM.Matrix matRLF2 = (SAPbouiCOM.Matrix)oForm.Items.Item("matRLF2").Specific;
                                             
                                            int j = 1;
                                            string Qry2 = " Select T1.U_iritemcd,T1.U_iritemnm, T1.U_irQtyPer, T1.U_irExQtP FROM dbo.[@EXSM] AS T0 LEFT JOIN dbo.[@XSM3] AS ";
                                            Qry2 = Qry2 + " T1 ON T0.Code = T1.Code Where  T0.U_schno = '" + ExpCode + "'";

                                            SAPbobsCOM.Recordset rec2 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                            rec2.DoQuery(Qry2);
                                            if (rec2.RecordCount > 0)
                                            {
                                                while (!rec2.EoF)
                                                {
                                                    matRLF2.AddRow();
                                                    (matRLF2.Columns.Item("#").Cells.Item(matRLF2.RowCount).Specific).Value = Convert.ToString(matRLF2.RowCount + 1);
                                                    (matRLF2.Columns.Item("ex5ic").Cells.Item(matRLF2.RowCount).Specific).Value = rec2.Fields.Item("U_iritemcd").Value;
                                                    (matRLF2.Columns.Item("ex5in").Cells.Item(matRLF2.RowCount).Specific).Value = rec2.Fields.Item("U_iritemnm").Value;
                                                    (matRLF2.Columns.Item("ex5ln").Cells.Item(matRLF2.RowCount).Specific).Value = ExpCode;
                                                    (matRLF2.Columns.Item("ex5np").Cells.Item(matRLF2.RowCount).Specific).Value = rec2.Fields.Item("U_irQtyPer").Value;
                                                    // double netwt = 22758 * (Convert.ToDouble(rec2.Fields.Item("U_irQtyPer").Value)) / 100;
                                                    double netwt = qty * (Convert.ToDouble(rec2.Fields.Item("U_irQtyPer").Value)) / 100;
                                                    (matRLF2.Columns.Item("ex5nw").Cells.Item(matRLF2.RowCount).Specific).Value = netwt;
                                                    (matRLF2.Columns.Item("ex5exp").Cells.Item(matRLF2.RowCount).Specific).Value = rec2.Fields.Item("U_irExQtP").Value;
                                                    (matRLF2.Columns.Item("ex5exw").Cells.Item(matRLF2.RowCount).Specific).Value = netwt * Convert.ToDouble(rec2.Fields.Item("U_irExQtP").Value);
                                                    j++;
                                                    rec2.MoveNext();
                                                }
                                            }
                                             
                                            ArrengeMatrixLineNum(matRLF2);
                                            DeleteMatrixBlankRowRFL2(matRLF2);
                                        }

                                        oForm.Freeze(false);
                                         
                                    }

                                    else if (pVal.ItemUID == "matEXPAC" && pVal.ColUID == "et3ede")
                                    {
                                        SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("matEXPAC").Specific;
                                        try
                                        {
                                            SAPbouiCOM.ComboBox cb4 = oMatrix.Columns.Item("et3yc").Cells.Item(pVal.Row).Specific; // oForm.Items.Item("cmbCRCY").Specific;
                                            string cmbcal = cb4.Selected.Value.ToString();
                                            double lineTotal = 0;
                                            double FClineTotal = 0;
                                            string Currency = null;
                                            double rate = 0;

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

                                            ((SAPbouiCOM.EditText)oMatrix.Columns.Item("et3edn").Cells.Item(pVal.Row).Specific).Value = oDataTable.GetValue("DocNum", 0).ToString();
                                            ((SAPbouiCOM.EditText)oMatrix.Columns.Item("et3ede").Cells.Item(pVal.Row).Specific).Value = oDataTable.GetValue("DocEntry", 0).ToString();
                                            ((SAPbouiCOM.EditText)oMatrix.Columns.Item("et3pl").Cells.Item(pVal.Row).Specific).Value = lineTotal.ToString();
                                            ((SAPbouiCOM.EditText)oMatrix.Columns.Item("et3pf").Cells.Item(pVal.Row).Specific).Value = FClineTotal.ToString();
                                            ((SAPbouiCOM.EditText)oMatrix.Columns.Item("et3cur").Cells.Item(pVal.Row).Specific).Value = Currency.ToString();
                                            ((SAPbouiCOM.EditText)oMatrix.Columns.Item("et3rt").Cells.Item(pVal.Row).Specific).Value = rate.ToString();

                                        }
                                        catch (Exception ex)
                                        {

                                        }
                                    }

                                    else if (pVal.ItemUID == "matRLF1" && pVal.ColUID == "ex4ic")
                                    {
                                        SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("matRLF1").Specific;
                                        ((SAPbouiCOM.EditText)oMatrix.Columns.Item("ex4ic").Cells.Item(pVal.Row).Specific).Value = oDataTable.GetValue("ItemCode", 0).ToString();
                                        ((SAPbouiCOM.EditText)oMatrix.Columns.Item("ex4in").Cells.Item(pVal.Row).Specific).Value = oDataTable.GetValue("ItemName", 0).ToString();
                                        AddMatrixRow(oMatrix, "ex4ic");
                                    }
                                    
                                    else if (pVal.ItemUID == "ex1slc")
                                    {
                                        oForm.Items.Item("ex1slc").Specific.value = oDataTable.GetValue("CardCode", 0).ToString();
                                        oForm.Items.Item("ex1sln").Specific.value = oDataTable.GetValue("CardName", 0).ToString();
                                    }
                                    else if (pVal.ItemUID == "fcCHAcd")
                                    {
                                        oForm.Items.Item("fcCHAcd").Specific.value = oDataTable.GetValue("CardCode", 0).ToString();
                                        oForm.Items.Item("fcCHAnm").Specific.value = oDataTable.GetValue("CardName", 0).ToString();
                                        oForm.Items.Item("fcCHApn").Specific.value = oDataTable.GetValue("Phone1", 0).ToString();
                                        oForm.Items.Item("fcCHAad").Specific.value = objCU.BPTopOneAddress(oDataTable.GetValue("CardCode", 0).ToString());
                                    }
                                    else if (pVal.ItemUID == "twbwc")
                                    {
                                        oForm.Items.Item("twbwc").Specific.value = oDataTable.GetValue("CardCode", 0).ToString();
                                        oForm.Items.Item("twbwn").Specific.value = oDataTable.GetValue("CardName", 0).ToString();
                                        oForm.Items.Item("twbwa").Specific.value = objCU.BPTopOneAddress(oDataTable.GetValue("CardCode", 0).ToString());
                                    }
                                    else if (pVal.ItemUID == "exbc")
                                    {
                                        oForm.Items.Item("exbc").Specific.value = oDataTable.GetValue("CardCode", 0).ToString();
                                        oForm.Items.Item("exbn").Specific.value = oDataTable.GetValue("CardName", 0).ToString();
                                    }
                                    else if (pVal.ItemUID == "exdn")
                                    {
                                        oForm.Items.Item("exdn").Specific.value = oDataTable.GetValue("DocNum", 0).ToString();
                                        oForm.Items.Item("exdnde").Specific.value = oDataTable.GetValue("DocEntry", 0).ToString();
                                    }
                                    else if (pVal.ItemUID == "exinvno")
                                    {
                                        oForm.Items.Item("exinvno").Specific.value = oDataTable.GetValue("DocNum", 0).ToString();
                                        oForm.Items.Item("exinvnode").Specific.value = oDataTable.GetValue("DocEntry", 0).ToString();
                                        oForm.Items.Item("exinvdt").Specific.value = oDataTable.GetValue("DocRate", 0).ToString();
                                        //string dt = oDataTable.GetValue("DocRate", 0).ToString();
                                        // oForm.Items.Item("exdd").Specific.value = DateTime.dt.ToString("yyyyMMdd");

                                        oForm.Items.Item("cifAFC").Specific.value = oDataTable.GetValue("DocTotalFC", 0).ToString();
                                        oForm.Items.Item("cifALC").Specific.value = oDataTable.GetValue("DocTotalFC", 0) * oDataTable.GetValue("DocRate", 0);
                                        oForm.Items.Item("cifDFC").Specific.value = oDataTable.GetValue("DocTotalFC", 0).ToString();
                                        oForm.Items.Item("cifDLC").Specific.value = oDataTable.GetValue("DocTotalFC", 0) * oDataTable.GetValue("DocRate", 0);

                                        oForm.Items.Item("fobAFC").Specific.value = oDataTable.GetValue("DocTotalFC", 0).ToString();
                                        oForm.Items.Item("fobALC").Specific.value = oDataTable.GetValue("DocTotalFC", 0) * oDataTable.GetValue("DocRate", 0);
                                        oForm.Items.Item("fobDFC").Specific.value = oDataTable.GetValue("DocTotalFC", 0).ToString();
                                        oForm.Items.Item("fobDLC").Specific.value = oDataTable.GetValue("DocTotalFC", 0) * oDataTable.GetValue("DocRate", 0);

                                    }
                                }
                                catch (Exception ex)
                                {
                                    //SBOMain.SBO_Application.MessageBox(ex.Message.ToString(), 1, "ok", "", "");
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
                            if (pVal.ItemUID == "1")
                            {
                                // oForm.Items.Item("tattach").Specific.value = SBOMain.Get_Attach_Folder_Path() + Path.GetFileName(BrowseFilePath);
                                if ((oForm.Mode == BoFormMode.fm_ADD_MODE) || (oForm.Mode == BoFormMode.fm_UPDATE_MODE))
                                {
                                    if (!String.IsNullOrEmpty(oForm.Items.Item("exdnde").Specific.Value.ToString()))
                                    {
                                        string statusval = oForm.Items.Item("exls").Specific.value.ToString();
                                        if (statusval == "")
                                        {
                                            SBOMain.SBO_Application.StatusBar.SetText("Please select type Licence/Scheme.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                            BubbleEvent = false;
                                        }
                                    }
                                }
                            }
                        }
                        if (pVal.BeforeAction == false)
                        {
                            if (pVal.ItemUID == "btnREFEXP")
                            {
                                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("matEXPAC").Specific;
                                string ExpForm = null;
                                string FormDocEntry = null;
                                double PLCAmt = 0;
                                double PFCAmt = 0;
                                double LCAmt = 0;
                                double FCAmt = 0;

                                for (int i = 1; i <= oMatrix.RowCount; i++)
                                {
                                    ExpForm = (oMatrix.Columns.Item("et3yc").Cells.Item(i).Specific).Value;
                                    FormDocEntry = (oMatrix.Columns.Item("et3ede").Cells.Item(i).Specific).Value;

                                    PLCAmt = Convert.ToDouble((oMatrix.Columns.Item("et3pf").Cells.Item(i).Specific).Value);
                                    PFCAmt = Convert.ToDouble((oMatrix.Columns.Item("et3pl").Cells.Item(i).Specific).Value);
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
                                        if (rec.RecordCount > 0)
                                        {
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
                                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("et3al").Cells.Item(i).Specific).Value = LCAmt.ToString();
                                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("et3af").Cells.Item(i).Specific).Value = FCAmt.ToString();
                                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("et3dl").Cells.Item(i).Specific).Value = (PLCAmt - LCAmt).ToString();
                                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("et3df").Cells.Item(i).Specific).Value = (PFCAmt - FCAmt).ToString();
                                }
                            }
                            if (pVal.ItemUID == "lbtLC")
                            {
                                string abc = oForm.Items.Item("exlcno").Specific.Value; 
                                objCU.FormLoadAndActivate("frmLCTrans", "mnsmEXIM012"); 
                                OutwardToLCMaster inEximTracking = new OutwardToLCMaster();
                                inEximTracking.lcno = abc;
                                ClsLCTrans oPrice = new ClsLCTrans(inEximTracking);
                                //oForm.Close();
                            }

                            if (pVal.ItemUID == "lbtIN")
                            {
                                string abc = null;
                                abc = oForm.Items.Item("ex1pol").Specific.Value; 
                                objCU.FormLoadAndActivate("frmInctMaster", "mnsmEXIM008"); 
                                OutwardToIncoMaster inIncoMaster = new OutwardToIncoMaster();
                                inIncoMaster.inctno = abc;
                                clsInctMaster oPort = new clsInctMaster(inIncoMaster);
                                //oForm.Close();
                            }

                            if (pVal.ItemUID == "lbdbkno" || pVal.ItemUID == "rodno")
                            {
                                string abc = null;
                                if (pVal.ItemUID == "lbdbkno")
                                {
                                    abc = oForm.Items.Item("dbkno").Specific.Value;
                                }
                                else if (pVal.ItemUID == "lbRoDTPNO")
                                {
                                    abc = oForm.Items.Item("rodno").Specific.Value;
                                }
                                objCU.FormLoadAndActivate("frmSCTrans", "mnsmEXIM003");

                                OutwardFromEximTracking inEximTracking = new OutwardFromEximTracking();
                                List<ETTransList> lstETTransList = new List<ETTransList>();
                                inEximTracking.ScriptNo = abc;

                                clsSCTrans oPrice = new clsSCTrans(lstETTransList, inEximTracking);

                            }
                                if (pVal.ItemUID == "lbPOL" || pVal.ItemUID == "lbtPOD" || pVal.ItemUID == "lbtPOR")
                            {
                                string abc = null;
                                if (pVal.ItemUID == "lbPOL")
                                {
                                   abc = oForm.Items.Item("ex1pol").Specific.Value;
                                }
                                else if (pVal.ItemUID == "lbtPOD")
                                {
                                    abc = oForm.Items.Item("ex1pod").Specific.Value;
                                }
                                else if (pVal.ItemUID == "lbtPOR")
                                {
                                    abc = oForm.Items.Item("ex1por").Specific.Value;
                                } 
                                objCU.FormLoadAndActivate("frmPortMaster", "mnsmEXIM007"); 

                                OutwardToPortMaster inPortMaster = new OutwardToPortMaster();
                                inPortMaster.portcode = abc;
                                clsPortMaster oPort = new clsPortMaster(inPortMaster);
                                //oForm.Close();
                            }
                            if (pVal.ItemUID == "matRS" && pVal.ColUID == "ex13sn")
                            {
                                SAPbouiCOM.Matrix matrix = oForm.Items.Item("matRS").Specific;
                                string abc = (matrix.Columns.Item("ex13sn").Cells.Item(pVal.Row).Specific).Value;
                                objCU.FormLoadAndActivate("frmSCTrans", "mnsmEXIM003"); 
                                
                                OutwardFromEximTracking inEximTracking = new OutwardFromEximTracking();
                                List<ETTransList> lstETTransList = new List<ETTransList>();
                                inEximTracking.ScriptNo = abc;

                                clsSCTrans oPrice = new clsSCTrans(lstETTransList, inEximTracking);
                                //oForm.Close();
                            }
                            if (
                                (pVal.ItemUID == "matETSC1" && pVal.ColUID == "ex10sc") ||
                                (pVal.ItemUID == "matETSC2" && pVal.ColUID == "ex11sc"))
                            {
                                string abc = null;

                                if (pVal.ItemUID == "matETSC1" && pVal.ColUID == "ex10sc")
                                {
                                    SAPbouiCOM.Matrix matrix = oForm.Items.Item("matETSC1").Specific;
                                    abc = (matrix.Columns.Item("ex10sc").Cells.Item(pVal.Row).Specific).Value;
                                }
                                if (pVal.ItemUID == "matETSC2" && pVal.ColUID == "ex11sc")
                                {
                                    SAPbouiCOM.Matrix matrix = oForm.Items.Item("matETSC2").Specific;
                                    abc = (matrix.Columns.Item("ex11sc").Cells.Item(pVal.Row).Specific).Value;
                                } 
                                objCU.FormLoadAndActivate("frmSchmMaster", "mnsmEXIM011"); 
                                OutwardToSchemeMaster OutwardToSchemeMaster = new OutwardToSchemeMaster();
                                OutwardToSchemeMaster.schmeno = abc;
                                clsSchmMaster oPrice = new clsSchmMaster(OutwardToSchemeMaster);
                                //oForm.Close();
                            }

                            if (pVal.ItemUID == "matEXPAC" && pVal.ColUID == "et3ede")
                            {
                                SAPbouiCOM.Matrix oMatrix = oForm.Items.Item("matEXPAC").Specific;
                                string abc = (oMatrix.Columns.Item("et3edn").Cells.Item(pVal.Row).Specific).Value;
                            }
                            if (pVal.ItemUID == "tab2")
                            {
                                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("matFOB").Specific;
                                AddMatrixRow(oMatrix, "pdipn");
                                doAutoSummatFOB(oForm);
                            }
                            if (pVal.ItemUID == "tab3")
                            {
                                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("matEXPAC").Specific;
                                AddMatrixRow(oMatrix, "et3expt");
                                doAutoSummatEXPAC(oForm);
                                for (int i = 1; i <= oMatrix.RowCount; i++)
                                {
                                    //SAPbouiCOM.Matrix matLCEX = (SAPbouiCOM.Matrix)oForm.Items.Item("matEXPAC").Specific;

                                    /*SAPbouiCOM.ComboBox cb4 = oMatrix.Columns.Item("et3yc").Cells.Item(i).Specific; // oForm.Items.Item("cmbCRCY").Specific;
                                    string cmbcal = cb4.Selected.Value.ToString();

                                    if (cmbcal == "PR")
                                    {
                                        setChooseFromListFieldMAT(oForm, CFL_PR, "CFL_PR", et3yc, "et3ede", oMatrix, i);
                                    }
                                    else if (cmbcal == "PQ")
                                    {
                                        setChooseFromListFieldMAT(oForm, CFL_PQ, "CFL_PQ", et3yc, "et3ede", oMatrix, i);
                                    }
                                    else if (cmbcal == "PO")
                                    {
                                        setChooseFromListFieldMAT(oForm, CFL_PO, "CFL_PO", et3yc, "et3ede", oMatrix, i);
                                    }
                                    else if (cmbcal == "PI")
                                    {
                                        setChooseFromListFieldMAT(oForm, CFL_PI, "CFL_PI", et3yc, "et3ede", oMatrix, i);
                                    }*/
                                }
                            }
                            if (pVal.ItemUID == "tab4")
                            {
                                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("matRLF1").Specific;
                                AddMatrixRow(oMatrix, "ex4ic");
                                doAutoSummatRLF1(oForm);

                                SAPbouiCOM.Matrix oMatrix1 = (SAPbouiCOM.Matrix)oForm.Items.Item("matRLF2").Specific;
                                AddMatrixRow(oMatrix1, "ex5ic");
                                doAutoSummatRLF2(oForm);
                            }
                            if (pVal.ItemUID == "tab5")
                            {
                                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("matETSC1").Specific;
                                AddMatrixRow(oMatrix, "ex10sc");
                                doAutoSummatETSC1(oForm);
                            }
                            if (pVal.ItemUID == "tab6")
                            {
                                SAPbouiCOM.Matrix oMatrix1 = (SAPbouiCOM.Matrix)oForm.Items.Item("matETSC2").Specific;
                                AddMatrixRow(oMatrix1, "ex11sc");
                                doAutoSummatETSC2(oForm);
                            }
                            if (pVal.ItemUID == "tab7")
                            {
                                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("matEXVGM").Specific;
                                AddMatrixRow(oMatrix, "ex7cs");
                            }
                            if (pVal.ItemUID == "tab8")
                            {
                                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("matEXCA").Specific;
                                AddMatrixRow(oMatrix, "ex8cn");
                            }
                            if (pVal.ItemUID == "tab9")
                            {
                                doAutoSummatRS(oForm);
                            }
                            if (pVal.ItemUID == "1") // && oForm.Mode == BoFormMode.fm_ADD_MODE)
                            {
                                try
                                {
                                    //SAPbouiCOM.Matrix matETSC2 = oForm.Items.Item("matETSC2").Specific;
                                    //ArrengeMatrixLineNum(matETSC2); 
                                }
                                catch (Exception ex)
                                {
                                    //  SBOMain.SBO_Application.MessageBox(ex.Message.ToString(), 1, "ok", "", "");
                                }
                            }
                        }
                        break;
                    case BoEventTypes.et_FORM_DATA_LOAD:
                        oForm = SBOMain.SBO_Application.Forms.ActiveForm;
                        if (pVal.BeforeAction == false)
                        {
                            changelabels(oForm);
                            setCFLExpenceTab(oForm);
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
                                    if (cb4.Selected == null) {   
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
                                Form_Load_Components(oForm, "OK");
                            }
                        }
                        break;
                    case BoEventTypes.et_COMBO_SELECT:
                        oForm = SBOMain.SBO_Application.Forms.ActiveForm;
                        if (pVal.BeforeAction == false)
                        {
                            if (pVal.ItemUID == "matEXPAC" && pVal.ColUID == "et3yc")
                            {
                                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                                {
                                    SBOMain.SBO_Application.StatusBar.SetText("Please Save Exim Transaction then add Expenses.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                                    BubbleEvent = false;
                                }
                                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                                {
                                    Program.ExTransData.ExNo = oForm.Items.Item("tCode").Specific.Value;
                                    SAPbouiCOM.Matrix matLCEX = (SAPbouiCOM.Matrix)oForm.Items.Item("matEXPAC").Specific;
                                    SAPbouiCOM.ComboBox cb4 = matLCEX.Columns.Item("et3yc").Cells.Item(pVal.Row).Specific; // oForm.Items.Item("cmbCRCY").Specific;
                                    string cmbcal = cb4.Selected.Value.ToString();
                                    //SAPbouiCOM.LinkedButton oLinkExpdoc = matLCEX.Columns.Item("lc5bden").Cells.Item(pVal.Row).Specific;

                                    SBOMain.sForm = "ET";
                                    if (cmbcal == "PR")
                                    {
                                        setChooseFromListFieldMAT(oForm, CFL_PR, "CFL_PR", et3yc, "et3ede", matLCEX, pVal.Row);
                                        SBOMain.LineNo = pVal.Row;
                                        SBOMain.TFromUID = oForm.UniqueID;
                                        SBOMain.FromCnt = oForm.TypeCount;

                                        matLCEX.Columns.Item("et3ede").Cells.Item(pVal.Row).Click();
                                        SBOMain.SBO_Application.SendKeys("{TAB}");

                                        //oLinkExpdoc.LinkedObjectType = "1470000113"; 
                                    }
                                    else if (cmbcal == "PQ")
                                    {
                                        setChooseFromListFieldMAT(oForm, CFL_PQ, "CFL_PQ", et3yc, "et3ede", matLCEX, pVal.Row);
                                        SBOMain.LineNo = pVal.Row;
                                        SBOMain.TFromUID = oForm.UniqueID;
                                        SBOMain.FromCnt = oForm.TypeCount;
                                        matLCEX.Columns.Item("et3ede").Cells.Item(pVal.Row).Click();
                                        SBOMain.SBO_Application.SendKeys("{TAB}");
                                    }
                                    else if (cmbcal == "PO")
                                    {
                                        setChooseFromListFieldMAT(oForm, CFL_PO, "CFL_PO", et3yc, "et3ede", matLCEX, pVal.Row);
                                        SBOMain.LineNo = pVal.Row;
                                        SBOMain.TFromUID = oForm.UniqueID;
                                        SBOMain.FromCnt = oForm.TypeCount;
                                        matLCEX.Columns.Item("et3ede").Cells.Item(pVal.Row).Click();
                                        SBOMain.SBO_Application.SendKeys("{TAB}");
                                    }
                                    else if (cmbcal == "PI")
                                    {
                                        setChooseFromListFieldMAT(oForm, CFL_PI, "CFL_PI", et3yc, "et3ede", matLCEX, pVal.Row);
                                        SBOMain.LineNo = pVal.Row;
                                        SBOMain.TFromUID = oForm.UniqueID;
                                        SBOMain.FromCnt = oForm.TypeCount;
                                        matLCEX.Columns.Item("et3ede").Cells.Item(pVal.Row).Click();
                                        SBOMain.SBO_Application.SendKeys("{TAB}");
                                    }
                                }
                                // string docdate = oForm.Items.Item("lcod").Specific.value;
                            }

                            if (pVal.ItemUID == "exls")
                            {
                                cb5 = oForm.Items.Item("exls").Specific;
                                string val1 = cb5.Selected.Value.ToString();

                                cb4 = oForm.Items.Item("exls").Specific;
                                string val = cb4.Selected.Value.ToString();

                                if (val == "Licence" || val1 == "E")
                                {
                                    fillRLFMatrix(oForm, oForm.Items.Item("exinvnode").Specific.value, oForm.Items.Item("exbc").Specific.value);
                                }
                                if (val == "Scheme" || val1 == "E")
                                {
                                    fillDBKMatrix(oForm, oForm.Items.Item("exinvnode").Specific.value, oForm.Items.Item("exbc").Specific.value);
                                    fillRoDTEPMatrix(oForm, oForm.Items.Item("exinvnode").Specific.value, oForm.Items.Item("exbc").Specific.value);
                                }
                            }
                            if (pVal.ItemUID == "exdt")
                            {
                                // It will empty fields based on DocType Combo
                                removeValues(oForm);
                                changelabels(oForm);

                            }
                            else if (pVal.ItemUID == "cSer" && pVal.FormMode == 3)
                            {
                                oForm.Items.Item("tDocNum").Specific.Value = oForm.BusinessObject.GetNextSerialNumber(oForm.Items.Item("cSer").Specific.Value, "EXET");
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
        private void setCFLExpenceTab(SAPbouiCOM.Form oForm)
        {
            try
            {
                SAPbouiCOM.Matrix matLCEX = (SAPbouiCOM.Matrix)oForm.Items.Item("matEXPAC").Specific;
                for (int i = 1; i <= matLCEX.RowCount; i++)
                {
                    SAPbouiCOM.ComboBox cb4 = matLCEX.Columns.Item("et3yc").Cells.Item(i).Specific; // oForm.Items.Item("cmbCRCY").Specific;
                    string cmbcal = cb4.Selected.Value.ToString();
                    if (cmbcal == "PR")
                    {
                        setChooseFromListFieldMAT(oForm, CFL_PR, "CFL_PR", et3yc, "et3ede", matLCEX, i);
                        //oLinkExpdoc.LinkedObjectType = "1470000113"; 
                    }
                    else if (cmbcal == "PQ")
                    {
                        setChooseFromListFieldMAT(oForm, CFL_PQ, "CFL_PQ", et3yc, "et3ede", matLCEX, i);
                        // oLinkExpdoc.LinkedObjectType = "540000006";
                    }
                    else if (cmbcal == "PO")
                    {
                        setChooseFromListFieldMAT(oForm, CFL_PO, "CFL_PO", et3yc, "et3ede", matLCEX, i);
                        // oLinkExpdoc.LinkedObjectType = "22";
                    }
                    else if (cmbcal == "PI")
                    {
                        setChooseFromListFieldMAT(oForm, CFL_PI, "CFL_PI", et3yc, "et3ede", matLCEX, i);
                        //oLinkExpdoc.LinkedObjectType = "18";
                    }
                }
            }
            catch (Exception ex)
            {

            }
        }
        private void changelabels(SAPbouiCOM.Form oForm)
        {
            cb4 = oForm.Items.Item("exdt").Specific;
            string DocType = cb4.Selected.Value.ToString();
            if (DocType == "E")
            {
                NameValueCollection list3 = new NameValueCollection()
                                                 {  { "1", "BP Code" }, { "2","BP Name" }, { "3", "Sales Order No" } ,
                                                    { "4", "Delivery No" }, { "5","A/R Invoice No" }, { "6", "A/R Invoice Ex. Rate" },
                                                    { "7","Shipping Bill Number" }, { "8", "Shipping Bill Date" }};

                // It will set Label Caption based on Doc Type 
                setLableText(oForm, list3);

                // It will set Choose From list based on Doc Type
                setChooseFromListForExport(oForm);
            }
            else
            {
                NameValueCollection list3 = new NameValueCollection()
                                                 {  { "1", "Vendor Code" }, { "2","Vendor Name" }, { "3", "Purchase Order No" } ,
                                                    { "4", "GRN No" }, { "5","A/P Invoice No" }, { "6", "A/P Invoice Ex. Rate" },
                                                    { "7","Bill of Entry No" }, { "8", "Bill of Entry Date" }};

                // It will set Label Caption based on Doc Type
                setLableText(oForm, list3);

                // It will set Choose From list based on Doc Type
                setChooseFromListForImport(oForm);

            }
        }
        private void setLableText(SAPbouiCOM.Form oForm, NameValueCollection list)
        {
            oForm.Items.Item("lblbpcode").Specific.Caption = list["1"];
            oForm.Items.Item("lblbpn").Specific.Caption = list["2"];
            oForm.Items.Item("Item_9").Specific.Caption = list["3"];
            oForm.Items.Item("Item_10").Specific.Caption = list["4"];
            oForm.Items.Item("Item_6").Specific.Caption = list["5"];
            oForm.Items.Item("Item_5").Specific.Caption = list["6"];
            oForm.Items.Item("Item_18").Specific.Caption = list["7"];
            oForm.Items.Item("Item_31").Specific.Caption = list["8"];
        }
        public void doAutoColSum(SAPbouiCOM.Matrix matrix, string ColumnName)
        {
            SAPbouiCOM.Column mCol = matrix.Columns.Item(ColumnName);
            mCol.RightJustified = true;
            mCol.ColumnSetting.SumType = BoColumnSumType.bst_Auto;
        }
        public void Form_Load_Components(SAPbouiCOM.Form oForm, string mode)
        {
            oForm.Freeze(true);

            oForm.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized;
            SAPbouiCOM.EditText oEdit;
            string Table = "@EXET";
            DateTime now = DateTime.Now;

            if (mode != "OK")
            {
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                oEdit = oForm.Items.Item("tCode").Specific;
                objCU.GetNextDocNum(ref oEdit, ref Table);
                oForm.Items.Item("tDocNum").Specific.value = oForm.BusinessObject.GetNextSerialNumber("DocEntry", "EXET");
                Events.Series.SeriesCombo("EXET", "cSer");
                oForm.Items.Item("cSer").DisplayDesc = true;

                oForm.Items.Item("tab1").Visible = true;
                oForm.Items.Item("tab1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                oForm.PaneLevel = 1;

                oForm.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized;

                oForm.Items.Item("exdd").Specific.value = DateTime.Now.ToString("yyyyMMdd");

                SAPbouiCOM.ComboBox cb = (SAPbouiCOM.ComboBox)oForm.Items.Item("tStatus").Specific;
                cb.ExpandType = BoExpandType.et_DescriptionOnly;
                cb.Select("O");

                setChooseFromListForExport(oForm);

                cb1 = (SAPbouiCOM.ComboBox)oForm.Items.Item("exdt").Specific;
                cb1.ExpandType = BoExpandType.et_DescriptionOnly;
                cb1.Select("E");

                /*cb2 = (SAPbouiCOM.ComboBox)oForm.Items.Item("exls").Specific;
                cb2.ExpandType = BoExpandType.et_DescriptionOnly;
                cb2.Select("Scheme");*/
                oForm.Items.Item("exbc").Click();
            }
            oForm.Freeze(false);
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
        private void ArrengeMatrixLineNumNew(SAPbouiCOM.Matrix matrix)
        {
            try
            {
                for (int i = 1; i < matrix.RowCount; i++)
                {
                    if (!string.IsNullOrEmpty(Convert.ToString(matrix.Columns.Item("#").Cells.Item(i).Specific.value)))
                    {
                        matrix.Columns.Item("#").Cells.Item(i).Specific.value = i;
                    }
                }
            }
            catch (Exception ex)
            {

            }
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
        #endregion

        private void CFLConditionBP(string CFLID, string ItemUID, string DocType)
        {
            oCFL = oForm.ChooseFromLists.Item(CFLID);
            oConds = SBOMain.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions);

            oCond = oConds.Add();
            oCond.Alias = "validFor";
            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
            oCond.CondVal = "Y";
            oCFL.SetConditions(oConds);

            if (DocType == "E")
            {
                oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;
                oCond = oConds.Add();
                oCond.Alias = "CardType";
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCond.CondVal = "C";
                oCFL.SetConditions(oConds);
            }
            else if (DocType == "I")
            {
                oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;
                oCond = oConds.Add();
                oCond.Alias = "CardType";
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCond.CondVal = "S";
                oCFL.SetConditions(oConds);
            }
            oCFL = null;
            oCond = null;
            oConds = null;
        }
        private void CFLConditionLICITEM(string CFLID, string ItemUID)
        {
            oCFL = oForm.ChooseFromLists.Item(CFLID);
            oConds = SBOMain.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions);

            if (CFLID == "CFLOITM2")
            {
                oCond = oConds.Add();
                oCond.Alias = "ItemCode";
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_START;
                oCond.CondVal = "LIC";

            }
            oCFL.SetConditions(oConds);
            oCFL = null;
            oCond = null;
            oConds = null;
        }

        private void CFLConditionIN(string CFLID, string ItemUID)
        {
            oCFL = oForm.ChooseFromLists.Item(CFLID);
            oConds = SBOMain.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions);

            if (CFLID == "CFL_22")
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
        private void CFLConditionLC(string CFLID, string ItemUID)
        {
            oCFL = oForm.ChooseFromLists.Item(CFLID);
            oConds = SBOMain.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions);

            if (CFLID == "CFL_21")
            {
                oCond = oConds.Add();
                oCond.Alias = "Status";
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCond.CondVal = "O";
                oCFL.SetConditions(oConds);
            }
            oCFL = null;
            oCond = null;
            oConds = null;
        }

        private void CFLConditionPort(string CFLID, string ItemUID)
        {
            oCFL = oForm.ChooseFromLists.Item(CFLID);
            oConds = SBOMain.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions);

            if (CFLID == "CFL_17" || CFLID == "CFL_19" || CFLID == "CFL_20")
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
        
        private void CFLConditionLicense(string CFLID, string ItemUID)
        {
            oCFL = oForm.ChooseFromLists.Item(CFLID);
            oConds = SBOMain.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions);

            if (CFLID == "CFL_16")
            {
                oCond = oConds.Add();
                oCond.Alias = "U_schtype";
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCond.CondVal = "Advanced";
                oCFL.SetConditions(oConds);
            }
            oCFL = null;
            oCond = null;
            oConds = null;

        }
        
        private void CFLConditionEXPType(string CFLID, string ItemUID)
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
            oCFL = null;
            oCond = null;
            oConds = null;

        }
          
        private void CFLCondition(string CFLID, string ItemUID)
        {
            oCFL = oForm.ChooseFromLists.Item(CFLID);
            oConds = SBOMain.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions);

            if (CFLID == "CFL_OCRD1" || CFLID == "CFL_OCRD2" || CFLID == "CFL_OCRD3")
            {
                oCond = oConds.Add();
                oCond.Alias = "validFor";
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCond.CondVal = "Y";
                oCFL.SetConditions(oConds);
            }
            if ((CFLID == "CFL_OCRD1" && ItemUID == "fcCHAcd") || (CFLID == "CFL_OCRD2" && ItemUID == "ex1slc") || (CFLID == "CFL_OCRD3" && ItemUID == "twbwc"))
            {
                oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;
                oCond = oConds.Add();
                oCond.Alias = "CardType";
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCond.CondVal = "S";
                oCFL.SetConditions(oConds);
            }

            oCFL = null;
            oCond = null;
            oConds = null;

        }

        private void CFLConditionTransORDR(string CFLID, string ItemUID, string CardCode, string SOPO, string delivery, string ARinvoice)
        {
            oCFL = oForm.ChooseFromLists.Item(CFLID);
            oConds = SBOMain.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions);

            oCond = oConds.Add();
            oCond.Alias = "CardCode";
            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
            oCond.CondVal = CardCode;

            oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;
            oCond = oConds.Add();
            oCond.Alias = "CANCELED";
            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
            oCond.CondVal = "N";

            if (String.IsNullOrEmpty(delivery))
            {
                SAPbobsCOM.Recordset rec1 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                string Qry1 = "SELECT U_exsonde from [dbo].[@EXET] WHERE U_exdt = 'E'";
                rec1.DoQuery(Qry1);
                if (rec1.RecordCount > 0)
                {
                    while (!rec1.EoF)
                    {
                        oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;
                        oCond = oConds.Add();
                        oCond.Alias = "DocEntry";
                        oCond.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_EQUAL;
                        oCond.CondVal = rec1.Fields.Item("U_exsonde").Value.ToString();
                        rec1.MoveNext();
                    }
                }
            }

            if (!String.IsNullOrEmpty(SOPO))
            {
                oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;
                oCond = oConds.Add();
                oCond.Alias = "DocEntry";
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCond.CondVal = SOPO;
            }

            oCFL.SetConditions(oConds);
            oCFL = null;
            oCond = null;
            oConds = null;
        }
        private void CFLConditionTransODLN(string CFLID, string ItemUID, string CardCode, string SOPO, string delivery, string ARinvoice)
        {
            oCFL = oForm.ChooseFromLists.Item(CFLID);
            oConds = SBOMain.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions);

            oCond = oConds.Add();
            oCond.Alias = "CardCode";
            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
            oCond.CondVal = CardCode;

            oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;
            oCond = oConds.Add();
            oCond.Alias = "CANCELED";
            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
            oCond.CondVal = "N";

            SAPbobsCOM.Recordset rec2 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            string Qry2 = "SELECT U_exdnde from [dbo].[@EXET] WHERE U_exdt = 'I'";
            rec2.DoQuery(Qry2);
            if (rec2.RecordCount > 0)
            {
                while (!rec2.EoF)
                {
                    oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;
                    oCond = oConds.Add();
                    oCond.Alias = "DocEntry";
                    oCond.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_EQUAL;
                    oCond.CondVal = rec2.Fields.Item("U_exdnde").Value.ToString();
                    rec2.MoveNext();
                }
            }
            if (!String.IsNullOrEmpty(delivery))
            {
                oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;
                oCond = oConds.Add();
                oCond.Alias = "DocEntry";
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCond.CondVal = delivery;
            }

            bool checkcon3 = false;
            string Qry3 = null;
            SAPbobsCOM.Recordset rec3 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            if (!String.IsNullOrEmpty(SOPO))
            {
                checkcon3 = true;
                Qry3 = "Select DISTINCT(T1.DocEntry) FROM DLN1 AS T1 WHERE T1.BaseType = 17 AND T1.BaseEntry = '" + SOPO + "'";
            }

            if (checkcon3)
            {
                rec3.DoQuery(Qry3);
                if (rec3.RecordCount > 0)
                {
                    while (!rec3.EoF)
                    {
                        oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;
                        oCond = oConds.Add();
                        oCond.Alias = "DocEntry";
                        oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        oCond.CondVal = rec3.Fields.Item("DocEntry").Value.ToString();
                        rec3.MoveNext();
                    }
                }
            }

            oCFL.SetConditions(oConds);

            oCFL = null;
            oCond = null;
            oConds = null;
        }
        private void CFLConditionTransOINV(string CFLID, string ItemUID, string CardCode, string SOPO, string delivery, string ARinvoice)
        {
            oCFL = oForm.ChooseFromLists.Item(CFLID);
            oConds = SBOMain.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions);

            oCond = oConds.Add();
            oCond.Alias = "CardCode";
            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
            oCond.CondVal = CardCode;

            oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;
            oCond = oConds.Add();
            oCond.Alias = "CANCELED";
            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
            oCond.CondVal = "N";

            SAPbobsCOM.Recordset rec4 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            SAPbobsCOM.Recordset rec1 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

            string Qry4 = "SELECT U_exinvnode from [dbo].[@EXET] WHERE U_exdt = 'E'";
            rec4.DoQuery(Qry4);
            if (rec4.RecordCount > 0)
            {
                while (!rec4.EoF)
                {
                    oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;
                    oCond = oConds.Add();
                    oCond.Alias = "DocEntry";
                    oCond.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_EQUAL;
                    oCond.CondVal = rec4.Fields.Item("U_exinvnode").Value.ToString();
                    rec4.MoveNext();
                }
            }

            bool checkcon1 = false;
            string Query1 = null;

            if (CFLID == "CFL_OINV" && (!String.IsNullOrEmpty(delivery)))
            {
                checkcon1 = true;
                Query1 = "Select DISTINCT(T1.DocEntry) FROM INV1 AS T1 WHERE T1.BaseType = 15 AND T1.BaseEntry = '" + delivery + "'";
            }
            else if (CFLID == "CFL_OINV" && (!String.IsNullOrEmpty(SOPO)))
            {
                checkcon1 = true;
                Query1 = "Select DISTINCT(T1.DocEntry) FROM INV1 AS T1 WHERE T1.BaseType = 17 AND T1.BaseEntry = '" + SOPO + "'";
            }
            if (checkcon1)
            {
                rec1.DoQuery(Query1);
                if (rec1.RecordCount > 0)
                {
                    while (!rec1.EoF)
                    {
                        oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;
                        oCond = oConds.Add();
                        oCond.Alias = "DocEntry";
                        oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        oCond.CondVal = rec1.Fields.Item("DocEntry").Value.ToString();
                        rec1.MoveNext();
                    }
                }
            }

            oCFL.SetConditions(oConds);

            oCFL = null;
            oCond = null;
            oConds = null;
        }

        private void CFLConditionTransOPOR(string CFLID, string ItemUID, string CardCode, string SOPO, string delivery, string ARinvoice)
        {
            oCFL = oForm.ChooseFromLists.Item(CFLID);
            oConds = SBOMain.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions);

            oCond = oConds.Add();
            oCond.Alias = "CardCode";
            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
            oCond.CondVal = CardCode;

            oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;
            oCond = oConds.Add();
            oCond.Alias = "CANCELED";
            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
            oCond.CondVal = "N";

            if (String.IsNullOrEmpty(delivery))
            {
                SAPbobsCOM.Recordset rec1 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                string Qry1 = "SELECT U_exsonde from [dbo].[@EXET] WHERE U_exdt = 'I'";
                rec1.DoQuery(Qry1);
                if (rec1.RecordCount > 0)
                {
                    while (!rec1.EoF)
                    {
                        oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;
                        oCond = oConds.Add();
                        oCond.Alias = "DocEntry";
                        oCond.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_EQUAL;
                        oCond.CondVal = rec1.Fields.Item("U_exsonde").Value.ToString();
                        rec1.MoveNext();
                    }
                }
            }

            if (!String.IsNullOrEmpty(SOPO))
            {
                oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;
                oCond = oConds.Add();
                oCond.Alias = "DocEntry";
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCond.CondVal = SOPO;
            }

            oCFL.SetConditions(oConds);
            oCFL = null;
            oCond = null;
            oConds = null;
        }
        private void CFLConditionTransOPDN(string CFLID, string ItemUID, string CardCode, string SOPO, string delivery, string ARinvoice)
        {
            oCFL = oForm.ChooseFromLists.Item(CFLID);
            oConds = SBOMain.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions);

            oCond = oConds.Add();
            oCond.Alias = "CardCode";
            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
            oCond.CondVal = CardCode;

            oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;
            oCond = oConds.Add();
            oCond.Alias = "CANCELED";
            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
            oCond.CondVal = "N";

            SAPbobsCOM.Recordset rec2 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            string Qry2 = "SELECT U_exdnde from [dbo].[@EXET] WHERE U_exdt = 'I'";
            rec2.DoQuery(Qry2);
            if (rec2.RecordCount > 0)
            {
                while (!rec2.EoF)
                {
                    oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;
                    oCond = oConds.Add();
                    oCond.Alias = "DocEntry";
                    oCond.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_EQUAL;
                    oCond.CondVal = rec2.Fields.Item("U_exdnde").Value.ToString();
                    rec2.MoveNext();
                }
            }
            if (!String.IsNullOrEmpty(delivery))
            {
                oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;
                oCond = oConds.Add();
                oCond.Alias = "DocEntry";
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCond.CondVal = delivery;
            }

            bool checkcon3 = false;
            string Qry3 = null;
            SAPbobsCOM.Recordset rec3 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            if (!String.IsNullOrEmpty(SOPO))
            {
                checkcon3 = true;
                Qry3 = "Select DISTINCT(T1.DocEntry) FROM PDN1 AS T1 WHERE T1.BaseType = 22 AND T1.BaseEntry = '" + SOPO + "'";
            }

            if (checkcon3)
            {
                rec3.DoQuery(Qry3);
                if (rec3.RecordCount > 0)
                {
                    while (!rec3.EoF)
                    {
                        oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;
                        oCond = oConds.Add();
                        oCond.Alias = "DocEntry";
                        oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        oCond.CondVal = rec3.Fields.Item("DocEntry").Value.ToString();
                        rec3.MoveNext();
                    }
                }
            }

            oCFL.SetConditions(oConds);

            oCFL = null;
            oCond = null;
            oConds = null;
        }
        private void CFLConditionTransOPCH(string CFLID, string ItemUID, string CardCode, string SOPO, string delivery, string ARinvoice)
        {
            oCFL = oForm.ChooseFromLists.Item(CFLID);
            oConds = SBOMain.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions);

            oCond = oConds.Add();
            oCond.Alias = "CardCode";
            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
            oCond.CondVal = CardCode;

            oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;
            oCond = oConds.Add();
            oCond.Alias = "CANCELED";
            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
            oCond.CondVal = "N";

            SAPbobsCOM.Recordset rec4 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            SAPbobsCOM.Recordset rec1 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

            string Qry4 = "SELECT U_exinvnode from [dbo].[@EXET] WHERE U_exdt = 'I'";
            rec4.DoQuery(Qry4);
            if (rec4.RecordCount > 0)
            {
                while (!rec4.EoF)
                {
                    oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;
                    oCond = oConds.Add();
                    oCond.Alias = "DocEntry";
                    oCond.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_EQUAL;
                    oCond.CondVal = rec4.Fields.Item("U_exinvnode").Value.ToString();
                    rec4.MoveNext();
                }
            }

            oCFL.SetConditions(oConds);

            oCFL = null;
            oCond = null;
            oConds = null;
        }

        private void CFLConditionTrans(string CFLID, string ItemUID, string CardCode, string SOPO, string delivery, string ARinvoice)
        {
            oCFL = oForm.ChooseFromLists.Item(CFLID);
            oConds = SBOMain.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions);

            if (CFLID == "CFL_ORDR" || CFLID == "CFL_OPOR")
            {
                oCond = oConds.Add();
                oCond.Alias = "CardCode";
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCond.CondVal = CardCode;

                oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;
                oCond = oConds.Add();
                oCond.Alias = "CANCELED";
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCond.CondVal = "N";

                SAPbobsCOM.Recordset rec = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                if (CFLID == "CFL_ORDR")
                {
                    Query = "SELECT U_exsonde from [dbo].[@EXET] WHERE U_exdt = 'E'";
                }
                else
                {
                    Query = "SELECT U_exsonde from [dbo].[@EXET] WHERE U_exdt = 'I'";
                }
                rec.DoQuery(Query);
                if (rec.RecordCount > 0)
                {
                    while (!rec.EoF)
                    {
                        oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;
                        oCond = oConds.Add();
                        oCond.Alias = "DocEntry";
                        oCond.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_EQUAL;
                        oCond.CondVal = rec.Fields.Item("U_exsonde").Value.ToString();
                        rec.MoveNext();
                    }
                }
                if (CFLID == "CFL_ORDR" && (!String.IsNullOrEmpty(SOPO)))
                {
                    oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;
                    oCond = oConds.Add();
                    oCond.Alias = "DocEntry";
                    oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    oCond.CondVal = SOPO;
                }
                oCFL.SetConditions(oConds);
            }

            if (CFLID == "CFL_ODLN" || CFLID == "CFL_OPDN")
            {
                oCond = oConds.Add();
                oCond.Alias = "CardCode";
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCond.CondVal = CardCode;

                oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;
                oCond = oConds.Add();
                oCond.Alias = "CANCELED";
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCond.CondVal = "N";

                SAPbobsCOM.Recordset rec = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                if (CFLID == "CFL_ODLN" || CFLID == "CFL_OPDN")
                {
                    Query = "SELECT U_exdnde from [dbo].[@EXET] WHERE U_exdt = 'E'";
                }
                else
                {
                    Query = "SELECT U_exdnde from [dbo].[@EXET] WHERE U_exdt = 'I'";
                }
                rec.DoQuery(Query);
                if (rec.RecordCount > 0)
                {
                    while (!rec.EoF)
                    {
                        oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;
                        oCond = oConds.Add();
                        oCond.Alias = "DocEntry";
                        oCond.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_EQUAL;
                        oCond.CondVal = rec.Fields.Item("U_exdnde").Value.ToString();
                        rec.MoveNext();
                    }
                }
                if ((CFLID == "CFL_ODLN" || CFLID == "CFL_OPDN") && (!String.IsNullOrEmpty(delivery)))
                {
                    oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;
                    oCond = oConds.Add();
                    oCond.Alias = "DocEntry";
                    oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    oCond.CondVal = delivery;
                }

                string Query3 = null;
                bool checkcon3 = false;
                if (CFLID == "CFL_ODLN" && (!String.IsNullOrEmpty(SOPO)))
                {
                    checkcon3 = true;
                    Query3 = "Select DISTINCT(T1.DocEntry) FROM DLN1 AS T1 WHERE T1.BaseType = 17 AND T1.BaseEntry = '" + SOPO + "'";
                }
                if (CFLID == "CFL_OPDN" && (!String.IsNullOrEmpty(SOPO)))
                {
                    checkcon3 = true;
                    Query3 = "Select DISTINCT(T1.DocEntry) FROM PDN1 AS T1 WHERE T1.BaseType = 22 AND T1.BaseEntry = '" + SOPO + "'";
                }

                if (checkcon3)
                {
                    rec.DoQuery(Query3);
                    if (rec.RecordCount > 0)
                    {
                        while (!rec.EoF)
                        {
                            oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;
                            oCond = oConds.Add();
                            oCond.Alias = "DocEntry";
                            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                            oCond.CondVal = rec.Fields.Item("DocEntry").Value.ToString();
                            rec.MoveNext();
                        }
                    }
                    else { checkcon3 = false; }
                }
                else { checkcon3 = false; }

                oCFL.SetConditions(oConds);
            }

            if (CFLID == "CFL_OINV" || CFLID == "CFL_OPCH")
            {
                oCond = oConds.Add();
                oCond.Alias = "CardCode";
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCond.CondVal = CardCode;

                oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;
                oCond = oConds.Add();
                oCond.Alias = "CANCELED";
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCond.CondVal = "N";

                SAPbobsCOM.Recordset rec = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                SAPbobsCOM.Recordset rec1 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

                if (CFLID == "CFL_OINV")
                {
                    Query = "SELECT U_exinvnode from [dbo].[@EXET] WHERE U_exdt = 'E'";
                }
                else
                {
                    Query = "SELECT U_exinvnode from [dbo].[@EXET] WHERE U_exdt = 'I'";
                }
                rec.DoQuery(Query);
                if (rec.RecordCount > 0)
                {
                    while (!rec.EoF)
                    {
                        oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;
                        oCond = oConds.Add();
                        oCond.Alias = "DocEntry";
                        oCond.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_EQUAL;
                        oCond.CondVal = rec.Fields.Item("U_exinvnode").Value.ToString();
                        rec.MoveNext();
                    }
                }

                if (CFLID == "CFL_OINV" && (!String.IsNullOrEmpty(ARinvoice)))
                {
                    oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;
                    oCond = oConds.Add();
                    oCond.Alias = "DocEntry";
                    oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    oCond.CondVal = ARinvoice;
                }

                bool checkcon1 = false;
                string Query1 = null;

                if (CFLID == "CFL_OINV" && (!String.IsNullOrEmpty(delivery)))
                {
                    checkcon1 = true;
                    Query1 = "Select DISTINCT(T1.DocEntry) FROM INV1 AS T1 WHERE T1.BaseType = 15 AND T1.BaseEntry = '" + delivery + "'";
                }
                else if (CFLID == "CFL_OINV" && (!String.IsNullOrEmpty(SOPO)))
                {
                    checkcon1 = true;
                    Query1 = "Select DISTINCT(T1.DocEntry) FROM INV1 AS T1 WHERE T1.BaseType = 17 AND T1.BaseEntry = '" + SOPO + "'";
                }
                if (CFLID == "CFL_OPCH" && (!String.IsNullOrEmpty(delivery)))
                {
                    checkcon1 = true;
                    Query1 = "Select DISTINCT(T1.DocEntry) FROM PCH1 AS T1 WHERE  T1.BaseType = 20  AND  T1.BaseEntry = '" + delivery + "'";
                }
                if (checkcon1)
                {
                    rec1.DoQuery(Query1);
                    if (rec1.RecordCount > 0)
                    {
                        while (!rec1.EoF)
                        {
                            oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;
                            oCond = oConds.Add();
                            oCond.Alias = "DocEntry";
                            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                            oCond.CondVal = rec1.Fields.Item("DocEntry").Value.ToString();
                            rec1.MoveNext();
                        }
                    }
                    else { checkcon1 = false; }
                }
                else { checkcon1 = false; }

                if (checkcon1 == false)
                {
                    bool checkcon = false;
                    string Query2 = null;

                    if (CFLID == "CFL_OINV" && (!String.IsNullOrEmpty(SOPO)))
                    {
                        Query2 = "Select DISTINCT(T1.DocEntry) FROM INV1 AS T1 WHERE T1.BaseType = 17 AND T1.BaseEntry = '" + SOPO + "'";
                        checkcon = true;
                    }
                    if (CFLID == "CFL_OPCH" && (!String.IsNullOrEmpty(SOPO)))
                    {
                        Query2 = "Select DISTINCT(T1.DocEntry) FROM PCH1 AS T1 WHERE T1.BaseType = 22 AND T1.BaseEntry = '" + SOPO + "'";
                        checkcon = true;
                    }
                    if (checkcon)
                    {
                        rec.DoQuery(Query);
                        if (rec.RecordCount > 0)
                        {
                            while (!rec.EoF)
                            {
                                oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;
                                oCond = oConds.Add();
                                oCond.Alias = "DocEntry";
                                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                                oCond.CondVal = rec.Fields.Item("DocEntry").Value.ToString();
                                rec.MoveNext();
                            }
                        }
                    }

                }
                oCFL.SetConditions(oConds);
            }
            oCFL = null;
            oCond = null;
            oConds = null;
        }

        private void removeValues(SAPbouiCOM.Form oForm)
        {
            oForm.Freeze(true);
            oForm.Items.Item("exbc").Specific.value = "";
            oForm.Items.Item("exbn").Specific.value = "";
            oForm.Items.Item("exson").Specific.value = "";
            oForm.Items.Item("exsonde").Specific.value = "";
            oForm.Items.Item("exdn").Specific.value = "";
            oForm.Items.Item("exdnde").Specific.value = "";
            oForm.Items.Item("exinvno").Specific.value = "";
            oForm.Items.Item("exinvnode").Specific.value = "";

            oForm.Items.Item("cifAFC").Specific.value = "";
            oForm.Items.Item("cifALC").Specific.value = "";
            oForm.Items.Item("frtAFC").Specific.value = "";
            oForm.Items.Item("frtALC").Specific.value = "";
            oForm.Items.Item("insAFC").Specific.value = "";
            oForm.Items.Item("insALC").Specific.value = "";
            oForm.Items.Item("ocAFC").Specific.value = "";
            oForm.Items.Item("ocALC").Specific.value = "";
            oForm.Items.Item("fobAFC").Specific.value = "";
            oForm.Items.Item("fobALC").Specific.value = "";

            oForm.Items.Item("cifDFC").Specific.value = "";
            oForm.Items.Item("cifDLC").Specific.value = "";
            oForm.Items.Item("frtDFC").Specific.value = "";
            oForm.Items.Item("frtDLC").Specific.value = "";
            oForm.Items.Item("insDFC").Specific.value = "";
            oForm.Items.Item("insDLC").Specific.value = "";
            oForm.Items.Item("ocDFC").Specific.value = "";
            oForm.Items.Item("ocDLC").Specific.value = "";
            oForm.Items.Item("fobDFC").Specific.value = "";
            oForm.Items.Item("fobDLC").Specific.value = "";
            oForm.Freeze(false);
        }
        private void setChooseFromListFieldMAT(SAPbouiCOM.Form oForm, SAPbouiCOM.ChooseFromList CFL, string clfname, SAPbouiCOM.EditText editext, string fieldname, SAPbouiCOM.Matrix matLCEX, int rowno)
        {
            CFL = oForm.ChooseFromLists.Item(clfname);
            SAPbouiCOM.Column oCol = matLCEX.Columns.Item("et3ede");
            oCol.ChooseFromListUID = clfname;
            oCol.ChooseFromListAlias = "DocEntry";

        }

        private void setChooseFromListField(SAPbouiCOM.Form oForm, SAPbouiCOM.ChooseFromList CFL, string clfname, SAPbouiCOM.EditText editext, string fieldname)
        {
            CFL = oForm.ChooseFromLists.Item(clfname);
            editext = oForm.Items.Item(fieldname).Specific;
            editext.ChooseFromListUID = clfname;
            editext.ChooseFromListAlias = "DocNum";
        }
        private void setChooseFromListForExport(SAPbouiCOM.Form oForm)
        {
            setChooseFromListField(oForm, CFL_ORDR, "CFL_ORDR", exson, "exson");
            setChooseFromListField(oForm, CFL_ODLN, "CFL_ODLN", exdn, "exdn");
            setChooseFromListField(oForm, CFL_OINV, "CFL_OINV", exinvno, "exinvno");

            oLinkSOPO = oForm.Items.Item("lbtSo").Specific;
            oLinkSOPO.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_Order;

            oLinkDLGR = oForm.Items.Item("lbtDN").Specific;
            oLinkDLGR.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_DeliveryNotes;

            oLinkINAP = oForm.Items.Item("lbtARIN").Specific;
            oLinkINAP.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_Invoice;

        }
        private void setChooseFromListForImport(SAPbouiCOM.Form oForm)
        {
            setChooseFromListField(oForm, CFL_OPOR, "CFL_OPOR", exson, "exson");
            setChooseFromListField(oForm, CFL_OPDN, "CFL_OPDN", exdn, "exdn");
            setChooseFromListField(oForm, CFL_OPCH, "CFL_OPCH", exinvno, "exinvno");

            oLinkSOPO = oForm.Items.Item("lbtSo").Specific;
            oLinkSOPO.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_PurchaseOrder;

            oLinkDLGR = oForm.Items.Item("lbtDN").Specific;
            oLinkDLGR.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_GoodsReceiptPO;

            oLinkINAP = oForm.Items.Item("lbtARIN").Specific;
            oLinkINAP.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_PurchaseInvoice;

        }
        private void CFLConditionTransORDR(string CFLID, string ItemUID, string CardCode)
        {
            oCFL = oForm.ChooseFromLists.Item(CFLID);
            oConds = SBOMain.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions);

            if (CFLID == "CFL_ORDR")
            {
                oCond = oConds.Add();
                oCond.Alias = "CardCode";
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCond.CondVal = CardCode;
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
                // SBOMain.SBO_Application.MessageBox(ex.Message.ToString(), 1, "ok", "", "");
            }
        }
        private void DeleteMatrixBlankRowRFL(SAPbouiCOM.Matrix oMatrix)
        {
            try
            {
                for (int i = oMatrix.VisualRowCount; i >= 1; i--)
                {
                    {
                        if (string.IsNullOrEmpty(((SAPbouiCOM.EditText)oMatrix.Columns.Item("ex4ic").Cells.Item(i).Specific).Value))
                        {
                            oMatrix.DeleteRow(i);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // SBOMain.SBO_Application.MessageBox(ex.Message.ToString(), 1, "ok", "", "");
            }
        }
        private void DeleteMatrixBlankRowRFL2(SAPbouiCOM.Matrix oMatrix)
        {
            try
            {
                for (int i = oMatrix.VisualRowCount; i >= 1; i--)
                {
                    {
                        if (string.IsNullOrEmpty(((SAPbouiCOM.EditText)oMatrix.Columns.Item("ex5ic").Cells.Item(i).Specific).Value))
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
        private void DeleteMatrixBlankRowDBK(SAPbouiCOM.Matrix oMatrix)
        {
            try
            {
                for (int i = oMatrix.VisualRowCount; i >= 1; i--)
                {
                    {
                        if (string.IsNullOrEmpty(((SAPbouiCOM.EditText)oMatrix.Columns.Item("ex10ic").Cells.Item(i).Specific).Value))
                        {
                            oMatrix.DeleteRow(i);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // SBOMain.SBO_Application.MessageBox(ex.Message.ToString(), 1, "ok", "", "");
            }
        }
        private void DeleteMatrixBlankRowRoDTEP(SAPbouiCOM.Matrix oMatrix)
        {
            try
            {
                for (int i = oMatrix.VisualRowCount; i >= 1; i--)
                {
                    {
                        if (string.IsNullOrEmpty(((SAPbouiCOM.EditText)oMatrix.Columns.Item("ex11ic").Cells.Item(i).Specific).Value))
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

        public void doAutoSummatFOB(SAPbouiCOM.Form oForm)
        {
            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("matFOB").Specific;
            doAutoColSum(oMatrix, "pdamtfc");
            doAutoColSum(oMatrix, "pdamtlc");
        }
        public void doAutoSummatEXPAC(SAPbouiCOM.Form oForm)
        {   
            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("matEXPAC").Specific;
            doAutoColSum(oMatrix, "et3pf");
            doAutoColSum(oMatrix, "et3pl");
            doAutoColSum(oMatrix, "et3af");
            doAutoColSum(oMatrix, "et3al");
            doAutoColSum(oMatrix, "et3df");
            doAutoColSum(oMatrix, "et3dl");
        }
        public void doAutoSummatRLF1(SAPbouiCOM.Form oForm)
        {
            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("matRLF1").Specific;
            doAutoColSum(oMatrix, "ex4lrqty");
            doAutoColSum(oMatrix, "ex4lrafc");
            doAutoColSum(oMatrix, "ex4lraklc");
            doAutoColSum(oMatrix, "ex4lfqty");
            doAutoColSum(oMatrix, "ex4lffc");
            doAutoColSum(oMatrix, "ex4lflc");
        }
        public void doAutoSummatRLF2(SAPbouiCOM.Form oForm)
        {
            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("matRLF2").Specific;
            doAutoColSum(oMatrix, "ex5nw");
            doAutoColSum(oMatrix, "ex5exw");
        }
        public void doAutoSummatETSC1(SAPbouiCOM.Form oForm)
        {
            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("matETSC1").Specific;
            doAutoColSum(oMatrix, "ex10qt");
            doAutoColSum(oMatrix, "ex10fofc");
            doAutoColSum(oMatrix, "ex10folc");
            doAutoColSum(oMatrix, "ex10frfc");
            doAutoColSum(oMatrix, "ex10frlc");
            doAutoColSum(oMatrix, "ex10infc");
            doAutoColSum(oMatrix, "ex10inlc");
            doAutoColSum(oMatrix, "ex10fbv");
            doAutoColSum(oMatrix, "ex10sv");
            doAutoColSum(oMatrix, "ex10vrpk");
            doAutoColSum(oMatrix, "ex10fv");
        }
        public void doAutoSummatETSC2(SAPbouiCOM.Form oForm)
        {
            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("matETSC2").Specific;
            doAutoColSum(oMatrix, "ex11qt");
            doAutoColSum(oMatrix, "ex11fofc");
            doAutoColSum(oMatrix, "ex11folc");
            doAutoColSum(oMatrix, "ex11frfc");
            doAutoColSum(oMatrix, "ex11frlc");
            doAutoColSum(oMatrix, "ex11infc");
            doAutoColSum(oMatrix, "ex11inlc");
            doAutoColSum(oMatrix, "ex11fbv");
            doAutoColSum(oMatrix, "ex11sv");
            doAutoColSum(oMatrix, "ex11vrpk");
            doAutoColSum(oMatrix, "ex11fv");
        }
        public void doAutoSummatRS(SAPbouiCOM.Form oForm)
        {
            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("matRS").Specific;
            doAutoColSum(oMatrix, "ex13bv");
            doAutoColSum(oMatrix, "ex13ra");
            doAutoColSum(oMatrix, "ex13sua");
        } 
     }
}
