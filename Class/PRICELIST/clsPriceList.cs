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

namespace CoreSuteConnect.Class.PRICELIST
{
    class clsPriceList
    {
        #region VariableDeclaration
        public static SAPbouiCOM.Application SBO_Application;
        private SAPbouiCOM.Form oForm;
        private SAPbouiCOM.Matrix oMatrix;
        private SAPbouiCOM.Columns oColumns;
        public string cFormID = string.Empty;

        int Progress = 0;
        private SAPbouiCOM.ChooseFromList oCFL = null;
        private SAPbouiCOM.Conditions oConds = null;
        private SAPbouiCOM.Condition oCond = null;
        FormHeader oHeader = new FormHeader();

        double stprice = 0;
        double total = 0;
        double freight1 = 0;
        double factExp1 = 0;
        double addexp11 = 0;
        double addexp21 = 0;
        double addexp31 = 0;
        double packing = 0;
        double disper = 0;
        double disamt = 0;
        double profper = 0;
        double profamt = 0;
        double untPrFC1 = 0;
        double stdPrice1COL = 0;
        double freight1COL = 0;
        double packingCOL = 0;
        double factExp1COL = 0;
        double addexp11COL = 0;
        double addexp21COL = 0;
        double addexp31COL = 0;
        double discountCOL = 0;
        double profitCOL = 0;
        double untPrice1COL = 0;
        double untPrFC1COL = 0;
        double freight = 0;
        double packingchg = 0;
        double factExp = 0;
        double addexp1 = 0;
        double addexp2 = 0;
        double addexp3 = 0;
        double exchRate = 0;
        string BPCode = null;
        string remarks = null;
        string FinalPrice = null;
        double disCOL = 0;
        double rate = 0;
        string currency = null;
        int rownum = 0;
        string rowyear = null;
        int rowmonth = 0;

        #endregion VariableDeclaration

        public clsPriceList(OutwardToPriceList outClass)
        {
            if (outClass != null)
            {
                oForm = SBOMain.SBO_Application.Forms.ActiveForm;

                int Progress = 0;
                SAPbouiCOM.ProgressBar oProgressBar;
                oProgressBar = SBOMain.SBO_Application.StatusBar.CreateProgressBar("Please Wait", outClass.lstOut.Count, true);
                oProgressBar.Text = "Please Wait";

                SAPbouiCOM.Matrix matrix = oForm.Items.Item("matLABMIX").Specific;

                /*if(oForm.Items.Item("cardcode").Specific.value)
                oForm.Items.Item("cardcode").Specific.value = outClass.BPCode;
                oForm.Items.Item("cardname").Specific.value = outClass.BPName;
                 */
                int clscnt = outClass.lstOut.Count;

                oForm.Freeze(true);
                for (int i = 0; i < clscnt; i++)
                {
                    Progress += 1;
                    oProgressBar.Value = Progress;
                    matrix.AddRow();
                    ((SAPbouiCOM.EditText)matrix.Columns.Item("#").Cells.Item(matrix.RowCount).Specific).Value = (matrix.RowCount).ToString();
                    ((SAPbouiCOM.EditText)matrix.Columns.Item("itemcode1").Cells.Item(matrix.RowCount).Specific).Value = outClass.lstOut[i].Itemno;
                    ((SAPbouiCOM.EditText)matrix.Columns.Item("itemname1").Cells.Item(matrix.RowCount).Specific).Value = outClass.lstOut[i].desc;
                    ((SAPbouiCOM.EditText)matrix.Columns.Item("ioref1").Cells.Item(matrix.RowCount).Specific).Value = outClass.lstOut[i].outwrdno;
                    ((SAPbouiCOM.EditText)matrix.Columns.Item("ptyref1").Cells.Item(matrix.RowCount).Specific).Value = outClass.lstOut[i].refno;
                }
                oForm.Freeze(false);
                oProgressBar.Stop();
            }
        }

        public bool MenuEvent(ref SAPbouiCOM.MenuEvent pVal, string FormId, string Type = "")
        {
            bool BubbleEvent = true;
            try
            {
                cFormID = FormId;
                oForm = SBOMain.SBO_Application.Forms.Item(FormId);

                oForm.EnableMenu("1292", true);//Add Row 
                oForm.EnableMenu("1293", true);//Delete Row
                oForm.EnableMenu("1287", true);//Duplicate Row

                SAPbouiCOM.Matrix matLABMIX = (SAPbouiCOM.Matrix)oForm.Items.Item("matLABMIX").Specific;
                SAPbouiCOM.Matrix MatFG = (SAPbouiCOM.Matrix)oForm.Items.Item("matFGITMS").Specific;

                if (pVal.BeforeAction == true)//&& pVal.MenuUID == "1287"
                { 
                    if (oForm.Mode == BoFormMode.fm_OK_MODE && Type == "Duplicate")
                    {
                        BPCode = oForm.Items.Item("cardcode").Specific.value;
                        oHeader.BPCode = BPCode;
                        freight = Convert.ToDouble(oForm.Items.Item("freight").Specific.value);
                        oHeader.freight = freight;
                        packingchg = Convert.ToDouble(oForm.Items.Item("packingchg").Specific.value);
                        oHeader.packingchg = packingchg;
                        factExp = Convert.ToDouble(oForm.Items.Item("factExp").Specific.value);
                        oHeader.factExp = factExp;
                        addexp1 = Convert.ToDouble(oForm.Items.Item("addexp1").Specific.value);
                        oHeader.addexp1 = addexp1;
                        addexp2 = Convert.ToDouble(oForm.Items.Item("addexp2").Specific.value);
                        oHeader.addexp2 = addexp2;
                        addexp3 = Convert.ToDouble(oForm.Items.Item("addexp3").Specific.value);
                        oHeader.addexp3 = addexp3;
                        exchRate = Convert.ToDouble(oForm.Items.Item("exchRate").Specific.value);
                        oHeader.exchRate = exchRate;
                        disper = Convert.ToDouble(oForm.Items.Item("disPer").Specific.value);
                        oHeader.disPer = disper;
                        disamt = Convert.ToDouble(oForm.Items.Item("disAmt").Specific.value);
                        oHeader.disAmt = disamt;
                        profper = Convert.ToDouble(oForm.Items.Item("profPer").Specific.value);
                        oHeader.profPer = profper;
                        profamt = Convert.ToDouble(oForm.Items.Item("profAmt").Specific.value);
                        oHeader.profAmt = profamt;
                        remarks = oForm.Items.Item("remarks").Specific.value;
                        oHeader.remarks = remarks;

                        // First Tab
                        SAPbouiCOM.Matrix MatLabMix = oForm.Items.Item("matLABMIX").Specific;
                        if (MatLabMix.RowCount > 0)
                        {
                            List<FormChild_Lab> frmChild_Lab = new List<FormChild_Lab>();
                            for (int i = 1; i <= MatLabMix.RowCount; i++)
                            {
                                FormChild_Lab ochild = new FormChild_Lab();
                                ochild.itemcode1 = ((SAPbouiCOM.EditText)MatLabMix.Columns.Item("itemcode1").Cells.Item(i).Specific).Value;
                                ochild.itemname1 = ((SAPbouiCOM.EditText)MatLabMix.Columns.Item("itemname1").Cells.Item(i).Specific).Value;
                                ochild.fgitemcode = ((SAPbouiCOM.EditText)MatLabMix.Columns.Item("fgitemcode").Cells.Item(i).Specific).Value;
                                ochild.fgitemname = ((SAPbouiCOM.EditText)MatLabMix.Columns.Item("fgitemname").Cells.Item(i).Specific).Value;
                                ochild.ioref1 = ((SAPbouiCOM.EditText)MatLabMix.Columns.Item("ioref1").Cells.Item(i).Specific).Value;
                                ochild.inoutno = ((SAPbouiCOM.EditText)MatLabMix.Columns.Item("inoutno").Cells.Item(i).Specific).Value;
                                ochild.ptyref1 = ((SAPbouiCOM.EditText)MatLabMix.Columns.Item("ptyref1").Cells.Item(i).Specific).Value;
                                ochild.stdPrice1 = ((SAPbouiCOM.EditText)MatLabMix.Columns.Item("stdPrice1").Cells.Item(i).Specific).Value;
                                ochild.freight1 = ((SAPbouiCOM.EditText)MatLabMix.Columns.Item("freight1").Cells.Item(i).Specific).Value;
                                ochild.packing = ((SAPbouiCOM.EditText)MatLabMix.Columns.Item("packing").Cells.Item(i).Specific).Value;
                                ochild.factExp1 = ((SAPbouiCOM.EditText)MatLabMix.Columns.Item("factExp1").Cells.Item(i).Specific).Value;
                                ochild.addexp11 = ((SAPbouiCOM.EditText)MatLabMix.Columns.Item("addexp11").Cells.Item(i).Specific).Value;
                                ochild.addexp21 = ((SAPbouiCOM.EditText)MatLabMix.Columns.Item("addexp21").Cells.Item(i).Specific).Value;
                                ochild.addexp31 = ((SAPbouiCOM.EditText)MatLabMix.Columns.Item("addexp31").Cells.Item(i).Specific).Value;
                                ochild.untPrice1 = ((SAPbouiCOM.EditText)MatLabMix.Columns.Item("untPrice1").Cells.Item(i).Specific).Value;
                                ochild.disPer1 = ((SAPbouiCOM.EditText)MatLabMix.Columns.Item("disPer1").Cells.Item(i).Specific).Value;
                                ochild.disAmt1 = ((SAPbouiCOM.EditText)MatLabMix.Columns.Item("disAmt1").Cells.Item(i).Specific).Value;
                                ochild.profPer1 = ((SAPbouiCOM.EditText)MatLabMix.Columns.Item("profPer1").Cells.Item(i).Specific).Value;
                                ochild.profAmt1 = ((SAPbouiCOM.EditText)MatLabMix.Columns.Item("profAmt1").Cells.Item(i).Specific).Value;
                                ochild.stdPrLC1 = ((SAPbouiCOM.EditText)MatLabMix.Columns.Item("stdPrLC1").Cells.Item(i).Specific).Value;
                                ochild.untPrFC1 = ((SAPbouiCOM.EditText)MatLabMix.Columns.Item("untPrFC1").Cells.Item(i).Specific).Value;

                                frmChild_Lab.Add(ochild);
                            }
                            oHeader.frmChild_Lab = frmChild_Lab;
                        }

                        if (MatFG.RowCount > 0)
                        {
                            List<FormChild_FG> lstFG = new List<FormChild_FG>();
                            for (int i = 1; i <= MatFG.RowCount; i++)
                            {
                                FormChild_FG ochild = new FormChild_FG();
                                ochild.itemcode2 = ((SAPbouiCOM.EditText)MatFG.Columns.Item("itemcode2").Cells.Item(i).Specific).Value;
                                ochild.itemname2 = ((SAPbouiCOM.EditText)MatFG.Columns.Item("itemname2").Cells.Item(i).Specific).Value;
                                ochild.stdprice2 = ((SAPbouiCOM.EditText)MatFG.Columns.Item("stdprice2").Cells.Item(i).Specific).Value;
                                ochild.remarks2 = ((SAPbouiCOM.EditText)MatFG.Columns.Item("remarks2").Cells.Item(i).Specific).Value;
                                lstFG.Add(ochild);
                            }

                            oHeader.frmChild_FG = lstFG;
                        }
                    }
                }

                if (pVal.BeforeAction == false)
                {
                    if ((oForm.Mode == BoFormMode.fm_ADD_MODE || oForm.Mode == BoFormMode.fm_UPDATE_MODE || Type == "ADDNEWFORM") && Type != "DEL_ROW")
                    {
                        Form_Load_Components(oForm,"ADD");
                        DeleteMatrixBlankRow(matLABMIX);
                    }
                    if (Type == "DEL_ROW")
                    {
                        DeleteMatrixBlankRow(matLABMIX);
                        ArrengeMatrixLineNum(matLABMIX);
                    }
                    else if (Type == "ADD_ROW")
                    {
                        if (SBOMain.RightClickItemID == "matLABMIX")
                        {
                            matLABMIX.AddRow(1, SBOMain.RightClickLineNum);
                            matLABMIX.ClearRowData(SBOMain.RightClickLineNum + 1);
                            ArrengeMatrixLineNum(matLABMIX);
                        }
                    }

                    else if (Type == "Duplicate")
                    {
                        if (oHeader != null)
                        {
                            // Header Field
                            oForm.Items.Item("cardcode").Specific.value = oHeader.BPCode;
                            oForm.Items.Item("freight").Specific.value = oHeader.freight;
                            oForm.Items.Item("packingchg").Specific.value = oHeader.packingchg;
                            oForm.Items.Item("factExp").Specific.value = oHeader.factExp;
                            oForm.Items.Item("addexp1").Specific.value = oHeader.addexp1;
                            oForm.Items.Item("addexp2").Specific.value = oHeader.addexp2;
                            oForm.Items.Item("addexp3").Specific.value = oHeader.addexp3;
                            oForm.Items.Item("exchRate").Specific.value = oHeader.exchRate;
                            oForm.Items.Item("disPer").Specific.value = oHeader.disPer;
                            oForm.Items.Item("disAmt").Specific.value = oHeader.disAmt;
                            oForm.Items.Item("profPer").Specific.value = oHeader.profPer;
                            oForm.Items.Item("profAmt").Specific.value = oHeader.profAmt;
                            oForm.Items.Item("remarks").Specific.value = oHeader.remarks;

                            // Child Field  
                            for (int i = 0; i < oHeader.frmChild_Lab.Count; i++)
                            {
                                matLABMIX.AddRow();
                                ((SAPbouiCOM.EditText)matLABMIX.Columns.Item("#").Cells.Item(matLABMIX.RowCount).Specific).Value = (i + 1).ToString();
                                ((SAPbouiCOM.EditText)matLABMIX.Columns.Item("itemcode1").Cells.Item(matLABMIX.RowCount).Specific).Value = oHeader.frmChild_Lab[i].itemcode1;
                                ((SAPbouiCOM.EditText)matLABMIX.Columns.Item("itemname1").Cells.Item(matLABMIX.RowCount).Specific).Value = oHeader.frmChild_Lab[i].itemname1;
                                ((SAPbouiCOM.EditText)matLABMIX.Columns.Item("fgitemcode").Cells.Item(matLABMIX.RowCount).Specific).Value = oHeader.frmChild_Lab[i].fgitemcode;
                                ((SAPbouiCOM.EditText)matLABMIX.Columns.Item("fgitemname").Cells.Item(matLABMIX.RowCount).Specific).Value = oHeader.frmChild_Lab[i].fgitemname;
                                ((SAPbouiCOM.EditText)matLABMIX.Columns.Item("ioref1").Cells.Item(matLABMIX.RowCount).Specific).Value = oHeader.frmChild_Lab[i].ioref1;
                                ((SAPbouiCOM.EditText)matLABMIX.Columns.Item("inoutno").Cells.Item(matLABMIX.RowCount).Specific).Value = oHeader.frmChild_Lab[i].inoutno;
                                ((SAPbouiCOM.EditText)matLABMIX.Columns.Item("ptyref1").Cells.Item(matLABMIX.RowCount).Specific).Value = oHeader.frmChild_Lab[i].ptyref1;
                                ((SAPbouiCOM.EditText)matLABMIX.Columns.Item("stdPrice1").Cells.Item(matLABMIX.RowCount).Specific).Value = oHeader.frmChild_Lab[i].stdPrice1;
                                ((SAPbouiCOM.EditText)matLABMIX.Columns.Item("freight1").Cells.Item(matLABMIX.RowCount).Specific).Value = oHeader.frmChild_Lab[i].freight1;
                                ((SAPbouiCOM.EditText)matLABMIX.Columns.Item("packing").Cells.Item(matLABMIX.RowCount).Specific).Value = oHeader.frmChild_Lab[i].packing;
                                ((SAPbouiCOM.EditText)matLABMIX.Columns.Item("factExp1").Cells.Item(matLABMIX.RowCount).Specific).Value = oHeader.frmChild_Lab[i].factExp1;
                                ((SAPbouiCOM.EditText)matLABMIX.Columns.Item("addexp11").Cells.Item(matLABMIX.RowCount).Specific).Value = oHeader.frmChild_Lab[i].addexp11;
                                ((SAPbouiCOM.EditText)matLABMIX.Columns.Item("addexp21").Cells.Item(matLABMIX.RowCount).Specific).Value = oHeader.frmChild_Lab[i].addexp21;
                                ((SAPbouiCOM.EditText)matLABMIX.Columns.Item("addexp31").Cells.Item(matLABMIX.RowCount).Specific).Value = oHeader.frmChild_Lab[i].addexp31;
                                ((SAPbouiCOM.EditText)matLABMIX.Columns.Item("untPrice1").Cells.Item(matLABMIX.RowCount).Specific).Value = oHeader.frmChild_Lab[i].untPrice1;
                                ((SAPbouiCOM.EditText)matLABMIX.Columns.Item("disPer1").Cells.Item(matLABMIX.RowCount).Specific).Value = oHeader.frmChild_Lab[i].disPer1;
                                ((SAPbouiCOM.EditText)matLABMIX.Columns.Item("disAmt1").Cells.Item(matLABMIX.RowCount).Specific).Value = oHeader.frmChild_Lab[i].disAmt1;
                                ((SAPbouiCOM.EditText)matLABMIX.Columns.Item("profPer1").Cells.Item(matLABMIX.RowCount).Specific).Value = oHeader.frmChild_Lab[i].profPer1;
                                ((SAPbouiCOM.EditText)matLABMIX.Columns.Item("profAmt1").Cells.Item(matLABMIX.RowCount).Specific).Value = oHeader.frmChild_Lab[i].profAmt1;
                                ((SAPbouiCOM.EditText)matLABMIX.Columns.Item("stdPrLC1").Cells.Item(matLABMIX.RowCount).Specific).Value = oHeader.frmChild_Lab[i].stdPrLC1;
                                ((SAPbouiCOM.EditText)matLABMIX.Columns.Item("untPrFC1").Cells.Item(matLABMIX.RowCount).Specific).Value = oHeader.frmChild_Lab[i].untPrFC1;
                            }

                            for (int i = 0; i < oHeader.frmChild_FG.Count; i++)
                            {
                                MatFG.AddRow();
                                ((SAPbouiCOM.EditText)MatFG.Columns.Item("#").Cells.Item(MatFG.RowCount).Specific).Value = (i + 1).ToString();
                                ((SAPbouiCOM.EditText)MatFG.Columns.Item("itemcode2").Cells.Item(MatFG.RowCount).Specific).Value = oHeader.frmChild_FG[i].itemcode2;
                                ((SAPbouiCOM.EditText)MatFG.Columns.Item("itemname2").Cells.Item(MatFG.RowCount).Specific).Value = oHeader.frmChild_FG[i].itemname2;
                                ((SAPbouiCOM.EditText)MatFG.Columns.Item("stdprice2").Cells.Item(MatFG.RowCount).Specific).Value = oHeader.frmChild_FG[i].stdprice2;
                                ((SAPbouiCOM.EditText)MatFG.Columns.Item("remarks2").Cells.Item(MatFG.RowCount).Specific).Value = oHeader.frmChild_FG[i].remarks2;
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

        public bool ItemEvent(ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent, string FormId)
        {
            BubbleEvent = true;
            oForm = SBOMain.SBO_Application.Forms.Item(FormId);
            SAPbouiCOM.Matrix matLABMIX = (SAPbouiCOM.Matrix)oForm.Items.Item("matLABMIX").Specific;

            try
            {
                switch (pVal.EventType)
                { // KEY Down event for multiplication

                    case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:
                        if (pVal.BeforeAction == false)
                        {

                        }
                        if (pVal.BeforeAction == true)
                        {

                        }
                        break;

                    case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:

                        if (pVal.BeforeAction == false)
                        { 
                            if (pVal.ItemUID == "disPer")
                            {
                                if (Convert.ToDouble(oForm.Items.Item("disPer").Specific.value) > 0)
                                {
                                    if (Convert.ToDouble(oForm.Items.Item("disAmt").Specific.value) > 0)
                                    {
                                        oForm.Items.Item("disAmt").Specific.value = 0;
                                    }
                                }
                            }
                            if (pVal.ItemUID == "disAmt")
                            {
                                if (Convert.ToDouble(oForm.Items.Item("disAmt").Specific.value) > 0)
                                {
                                    if (Convert.ToDouble(oForm.Items.Item("disPer").Specific.value) > 0)
                                    {
                                        oForm.Items.Item("disPer").Specific.value = 0;
                                    }
                                }
                            }

                            if (pVal.ItemUID == "profPer")
                            {
                                if (Convert.ToDouble(oForm.Items.Item("profPer").Specific.value) > 0)
                                {
                                    if (Convert.ToDouble(oForm.Items.Item("profAmt").Specific.value) > 0)
                                    {
                                        oForm.Items.Item("profAmt").Specific.value = 0;
                                    }
                                }
                            }
                            if (pVal.ItemUID == "profAmt")
                            {
                                if (Convert.ToDouble(oForm.Items.Item("profAmt").Specific.value) > 0)
                                {
                                    if (Convert.ToDouble(oForm.Items.Item("profPer").Specific.value) > 0)
                                    {
                                        oForm.Items.Item("profPer").Specific.value = 0;
                                    }
                                }
                            }
                        }
                        break;

                    case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
                        if (pVal.BeforeAction == false)
                        {
                            if (pVal.ItemUID == "exchRate")
                            {
                                // int abc = 0;
                            }
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
                                try
                                {
                                    if ((pVal.ColUID == "freight1" || pVal.ColUID == "packing" || pVal.ColUID == "factExp1" || pVal.ColUID == "addexp11" || pVal.ColUID == "addexp21" || pVal.ColUID == "addexp31"))
                                    {
                                        stdPrice1COL = Convert.ToDouble(((SAPbouiCOM.EditText)matLABMIX.Columns.Item("stdPrice1").Cells.Item(pVal.Row).Specific).Value);
                                        freight1COL = Convert.ToDouble(((SAPbouiCOM.EditText)matLABMIX.Columns.Item("freight1").Cells.Item(pVal.Row).Specific).Value);
                                        packingCOL = Convert.ToDouble(((SAPbouiCOM.EditText)matLABMIX.Columns.Item("packing").Cells.Item(pVal.Row).Specific).Value);
                                        factExp1COL = Convert.ToDouble(((SAPbouiCOM.EditText)matLABMIX.Columns.Item("factExp1").Cells.Item(pVal.Row).Specific).Value);
                                        addexp11COL = Convert.ToDouble(((SAPbouiCOM.EditText)matLABMIX.Columns.Item("addexp11").Cells.Item(pVal.Row).Specific).Value);
                                        addexp21COL = Convert.ToDouble(((SAPbouiCOM.EditText)matLABMIX.Columns.Item("addexp21").Cells.Item(pVal.Row).Specific).Value);
                                        addexp31COL = Convert.ToDouble(((SAPbouiCOM.EditText)matLABMIX.Columns.Item("addexp31").Cells.Item(pVal.Row).Specific).Value);
                                        discountCOL = Convert.ToDouble(((SAPbouiCOM.EditText)matLABMIX.Columns.Item("disAmt1").Cells.Item(pVal.Row).Specific).Value);
                                        profitCOL = Convert.ToDouble(((SAPbouiCOM.EditText)matLABMIX.Columns.Item("profAmt1").Cells.Item(pVal.Row).Specific).Value);

                                        exchRate = Convert.ToDouble(oForm.Items.Item("exchRate").Specific.value);
                                        untPrice1COL = stdPrice1COL + freight1COL + packingCOL + factExp1COL + addexp11COL + addexp21COL + addexp31COL;
                                        ((SAPbouiCOM.EditText)matLABMIX.Columns.Item("untPrice1").Cells.Item(pVal.Row).Specific).Value = Convert.ToString(untPrice1COL);
                                        untPrice1COL = untPrice1COL - discountCOL + profitCOL;

                                        exchRate = Convert.ToDouble(oForm.Items.Item("exchRate").Specific.value);
                                        ((SAPbouiCOM.EditText)matLABMIX.Columns.Item("stdPrLC1").Cells.Item(pVal.Row).Specific).Value = untPrice1COL.ToString();
                                        if (exchRate > 0)
                                        {
                                            untPrFC1COL = untPrice1COL / exchRate;
                                            FinalPrice = getFCPrice(untPrFC1COL);
                                            ((SAPbouiCOM.EditText)matLABMIX.Columns.Item("untPrFC1").Cells.Item(pVal.Row).Specific).Value = FinalPrice;
                                        }
                                    }
                                    else if (pVal.ColUID == "disPer1" || pVal.ColUID == "disAmt1" || pVal.ColUID == "profPer1" || pVal.ColUID == "profAmt1")
                                    {
                                        stprice = Convert.ToDouble(((SAPbouiCOM.EditText)matLABMIX.Columns.Item("untPrice1").Cells.Item(pVal.Row).Specific).Value);
                                        exchRate = Convert.ToDouble(oForm.Items.Item("exchRate").Specific.value);

                                        if (pVal.ColUID == "disPer1")
                                        {
                                            disper = Convert.ToDouble(((SAPbouiCOM.EditText)matLABMIX.Columns.Item("disPer1").Cells.Item(pVal.Row).Specific).Value);
                                            profitCOL = Convert.ToDouble(((SAPbouiCOM.EditText)matLABMIX.Columns.Item("profAmt1").Cells.Item(pVal.Row).Specific).Value);
                                            disamt = (stprice * disper) / 100;
                                            ((SAPbouiCOM.EditText)matLABMIX.Columns.Item("disAmt1").Cells.Item(pVal.Row).Specific).Value = disamt.ToString();
                                            untPrice1COL = stprice - disamt + profitCOL;
                                        }
                                        else if (pVal.ColUID == "disAmt1")
                                        {
                                            disamt = Convert.ToDouble(((SAPbouiCOM.EditText)matLABMIX.Columns.Item("disAmt1").Cells.Item(pVal.Row).Specific).Value);
                                            profitCOL = Convert.ToDouble(((SAPbouiCOM.EditText)matLABMIX.Columns.Item("profAmt1").Cells.Item(pVal.Row).Specific).Value);
                                            disper = (100 * disamt) / stprice;
                                            ((SAPbouiCOM.EditText)matLABMIX.Columns.Item("disPer1").Cells.Item(pVal.Row).Specific).Value = disper.ToString();
                                            untPrice1COL = stprice - disamt + profitCOL;
                                        }
                                        else if (pVal.ColUID == "profPer1")
                                        {
                                            profper = Convert.ToDouble(((SAPbouiCOM.EditText)matLABMIX.Columns.Item("profPer1").Cells.Item(pVal.Row).Specific).Value);
                                            disCOL = Convert.ToDouble(((SAPbouiCOM.EditText)matLABMIX.Columns.Item("disAmt1").Cells.Item(pVal.Row).Specific).Value);
                                            profamt = (stprice * profper) / 100;
                                            ((SAPbouiCOM.EditText)matLABMIX.Columns.Item("profAmt1").Cells.Item(pVal.Row).Specific).Value = profamt.ToString();
                                            untPrice1COL = stprice - disCOL + profamt;
                                        }
                                        else if (pVal.ColUID == "profAmt1")
                                        {
                                            profamt = Convert.ToDouble(((SAPbouiCOM.EditText)matLABMIX.Columns.Item("profAmt1").Cells.Item(pVal.Row).Specific).Value);
                                            disCOL = Convert.ToDouble(((SAPbouiCOM.EditText)matLABMIX.Columns.Item("disAmt1").Cells.Item(pVal.Row).Specific).Value);
                                            profper = (100 * profamt) / stprice;
                                            ((SAPbouiCOM.EditText)matLABMIX.Columns.Item("profPer1").Cells.Item(pVal.Row).Specific).Value = profper.ToString();
                                            untPrice1COL = stprice - disCOL + profamt;
                                        }

                                        ((SAPbouiCOM.EditText)matLABMIX.Columns.Item("stdPrLC1").Cells.Item(pVal.Row).Specific).Value = untPrice1COL.ToString();
                                        if (exchRate > 0)
                                        {
                                            untPrFC1COL = untPrice1COL / exchRate;
                                            FinalPrice = getFCPrice(untPrFC1COL);
                                            ((SAPbouiCOM.EditText)matLABMIX.Columns.Item("untPrFC1").Cells.Item(pVal.Row).Specific).Value = FinalPrice;
                                        }
                                    }
                                    else if (pVal.ItemUID == "exchRate" && matLABMIX.RowCount > 0)
                                    {
                                        int Progress = 0;
                                        SAPbouiCOM.ProgressBar oProgressBar;
                                        oProgressBar = SBOMain.SBO_Application.StatusBar.CreateProgressBar("Please Wait", matLABMIX.RowCount, true);
                                        oProgressBar.Text = "Please Wait";

                                        oForm.Freeze(true);
                                        for (int i = 1; i <= matLABMIX.RowCount; i++)
                                        {
                                            Progress += 1;
                                            oProgressBar.Value = Progress;

                                            total = Convert.ToDouble(((SAPbouiCOM.EditText)matLABMIX.Columns.Item("stdPrLC1").Cells.Item(i).Specific).Value);
                                            exchRate = Convert.ToDouble(oForm.Items.Item("exchRate").Specific.value);
                                            if (exchRate > 0)
                                            {
                                                untPrFC1 = total / exchRate;
                                                FinalPrice = getFCPrice(untPrFC1);
                                                ((EditText)matLABMIX.Columns.Item("untPrFC1").Cells.Item(i).Specific).Value = FinalPrice;

                                            }
                                        }
                                        oForm.Freeze(false);
                                        oProgressBar.Stop();

                                    }

                                    else if (pVal.ItemUID == "docDate")
                                    {
                                        SAPbouiCOM.ComboBox cb4 = oForm.Items.Item("cmbCRCY").Specific;
                                        currency = cb4.Selected.Value.ToString();

                                        if (currency != "##")
                                        {
                                            string docdate = oForm.Items.Item("docDate").Specific.value;
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
                                                oForm.Items.Item("exchRate").Specific.value = rate.ToString();
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
                                                oForm.Items.Item("exchRate").Specific.value = rate.ToString();
                                            }


                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    SBOMain.SBO_Application.StatusBar.SetText("Calculation Exception : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                    BubbleEvent = false;
                                }
                            }
                        }
                        break;
                    case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
                        //SAPbouiCOM.Form oForms = SBOMain.SBO_Application.Forms.ActiveForm;
                        if (pVal.BeforeAction == true)
                        {
                            if (pVal.ItemUID == "cardcode")
                            {
                                CFLCondition("CFL_OCRD", pVal.ItemUID);
                            }
                            if (pVal.ItemUID == "matLABMIX" && pVal.ColUID == "fgitemcode")
                            {
                                CFLCondition("CFL_FGItem", pVal.ItemUID);
                            }
                        }
                        if (pVal.BeforeAction == false)
                        {
                            rownum = 0;
                            SAPbouiCOM.IChooseFromListEvent oCFLEvento = null;
                            oCFLEvento = ((SAPbouiCOM.IChooseFromListEvent)(pVal));
                            string sCFL_ID = oCFLEvento.ChooseFromListUID;
                            SAPbouiCOM.ChooseFromList oCFL = null;
                            oCFL = oForm.ChooseFromLists.Item(sCFL_ID);
                            SAPbouiCOM.DataTable oDataTable = null;
                            oDataTable = oCFLEvento.SelectedObjects;
                            if (oDataTable != null)
                            {
                                if (pVal.ItemUID == "cardcode")
                                {
                                    try
                                    {
                                        oForm.Items.Item("cardcode").Specific.value = oDataTable.GetValue("CardCode", 0).ToString();
                                        oForm.Items.Item("cardname").Specific.value = oDataTable.GetValue("CardName", 0).ToString();

                                        string bpcode = oDataTable.GetValue("CardCode", 0).ToString();
                                        rate = 1.0;
                                        currency = "INR";

                                        string getQuery1 = @"SELECT Currency FROM OCRD Where CardCode ='" + bpcode + "'";
                                        SAPbobsCOM.Recordset rec1;
                                        rec1 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                        rec1.DoQuery(getQuery1);
                                        if (rec1.RecordCount > 0)
                                        {
                                            while (!rec1.EoF)
                                            {
                                                currency = Convert.ToString(rec1.Fields.Item("Currency").Value);
                                                rec1.MoveNext();
                                            }
                                        }
                                        if (currency != "##")
                                        {
                                            string docdate = oForm.Items.Item("docDate").Specific.value;
                                            string CurYear = docdate.Substring(0, 4);
                                            string CurMonth = docdate.Substring(4, 2);
                                            string CurDate = docdate.Substring(6, 2);
                                            string FromDateConvert = docdate.Substring(0, 4) + "-" + docdate.Substring(4, 2) + "-" + docdate.Substring(6, 2);
                                            DateTime daten = new DateTime(2020, Convert.ToInt16(CurMonth), 1);
                                            //rowmonth = daten.ToString("MMMM");
                                            rowmonth = Convert.ToInt16(CurMonth);
                                            rownum = Convert.ToInt16(CurDate);
                                            rowyear = CurYear; // Convert.ToInt16(CurYear);

                                            string getQuery = @"SELECT Rate FROM ORTT WHERE Currency = (SELECT Currency FROM OCRD Where CardCode ='" + bpcode + "') and RateDate = '" + FromDateConvert + "'";
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
                                            }
                                            else
                                            {

                                                if (currency != "INR")
                                                {
                                                    BubbleEvent = false;
                                                    openExchangeRateForm(currency, rownum, rowmonth, rowyear);
                                                }
                                            }

                                        }
                                        oForm.Items.Item("exchRate").Specific.value = rate.ToString();
                                        SAPbouiCOM.ComboBox cb4 = oForm.Items.Item("cmbCRCY").Specific;
                                        cb4.Select(currency);

                                    }
                                    catch (Exception ex)
                                    {

                                    }
                                }

                                if (pVal.ItemUID == "matLABMIX" && pVal.ColUID == "fgitemcode")
                                {
                                    SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("matLABMIX").Specific;
                                    try
                                    {
                                        ((SAPbouiCOM.EditText)oMatrix.Columns.Item("fgitemcode").Cells.Item(pVal.Row).Specific).Value = oDataTable.GetValue("ItemCode", 0).ToString();
                                        ((SAPbouiCOM.EditText)oMatrix.Columns.Item("fgitemname").Cells.Item(pVal.Row).Specific).Value = oDataTable.GetValue("ItemName", 0).ToString();
                                    }
                                    catch (Exception ex)
                                    {

                                    }
                                }
                            }
                        }
                        break;
                    case BoEventTypes.et_FORM_CLOSE:
                        if (pVal.BeforeAction == true)
                        {
                            oForm = SBOMain.SBO_Application.Forms.ActiveForm;
                            if (oForm.UniqueID == "123")
                            {

                            }
                        }
                        if (pVal.BeforeAction == false)
                        {
                            oForm = null;
                            GC.Collect();
                            GC.WaitForPendingFinalizers();
                        }
                        break;
                    case BoEventTypes.et_COMBO_SELECT:
                        oForm = SBOMain.SBO_Application.Forms.ActiveForm;

                        if (pVal.BeforeAction == true)
                        {
                            oForm = SBOMain.SBO_Application.Forms.ActiveForm;
                            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("matLABMIX").Specific;
                            if (pVal.ItemUID == "cmbREFI")
                            {
                                // Validation
                                if (oForm.Mode == BoFormMode.fm_ADD_MODE || oForm.Mode == BoFormMode.fm_UPDATE_MODE)
                                {
                                    if (matLABMIX.RowCount < 1)
                                    {
                                        SBOMain.SBO_Application.StatusBar.SetText("LabMix Sample Data Grid should not be empty.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                        BubbleEvent = false;
                                    }

                                }
                            }
                        }
                        if (pVal.BeforeAction == false)
                        {
                            rownum = 0;
                            oForm = SBOMain.SBO_Application.Forms.ActiveForm;
                            if (pVal.ItemUID == "cmbCRCY")
                            {
                                SAPbouiCOM.ComboBox cb4 = oForm.Items.Item("cmbCRCY").Specific;
                                currency = cb4.Selected.Value.ToString();

                                if (currency != "##")
                                {
                                    string docdate = oForm.Items.Item("docDate").Specific.value;
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
                                    }
                                    oForm.Items.Item("exchRate").Specific.value = rate.ToString();

                                }
                            }
                        }

                        break;

                    case BoEventTypes.et_CLICK:

                        oForm = SBOMain.SBO_Application.Forms.ActiveForm;
                        if (pVal.BeforeAction == true)
                        {
                            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("matLABMIX").Specific;
                            if (pVal.ItemUID == "1")
                                DeleteMatrixBlankRow(oMatrix);


                            // Validation : Without Add / Update not allow to perform Copy to.
                            if (pVal.ItemUID == "cmbCPYT")
                            {
                                //if (oForm.Mode == BoFormMode.fm_ADD_MODE || oForm.Mode == BoFormMode.fm_UPDATE_MODE)
                                if (oForm.Mode == BoFormMode.fm_ADD_MODE)
                                {
                                    SBOMain.SBO_Application.StatusBar.SetText("Please first Add the pricelist then perform copy.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                    BubbleEvent = false;
                                }
                            }
                            if (pVal.ItemUID == "btnCP" || pVal.ItemUID == "1")
                            {

                                SAPbouiCOM.ComboBox cb4 = oForm.Items.Item("cmbCRCY").Specific;
                                if (cb4.Selected.Value.ToString() == "##")
                                {
                                    SBOMain.SBO_Application.StatusBar.SetText("Please Select Currency!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                    BubbleEvent = false;
                                }
                                else if (oForm.Mode != BoFormMode.fm_FIND_MODE)
                                {     
                                    if (matLABMIX.RowCount < 1)
                                    {
                                        SBOMain.SBO_Application.StatusBar.SetText("LabMix Sample Data Grid should not be empty.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                        BubbleEvent = false;
                                    }
                               
                                    else if (string.IsNullOrEmpty(oForm.Items.Item("cardcode").Specific.value))
                                 {
                                    SBOMain.SBO_Application.StatusBar.SetText("Please select Cardcode.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                    BubbleEvent = false;
                                    }
                                }
                            }
                            if (pVal.ItemUID == "cmbREFI")
                            {
                                if (oForm.Mode == BoFormMode.fm_ADD_MODE || oForm.Mode == BoFormMode.fm_UPDATE_MODE)
                                {
                                    if (string.IsNullOrEmpty(oForm.Items.Item("cardcode").Specific.value))
                                    {
                                        SBOMain.SBO_Application.StatusBar.SetText("Please select Cardcode.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                        BubbleEvent = false;
                                    }
                                }
                            }
                        }

                        if (pVal.BeforeAction == false)
                        {
                            if (pVal.ItemUID == "btnCP")
                            {
                                int Progress = 0;
                                SAPbouiCOM.ProgressBar oProgressBar;
                                oProgressBar = SBOMain.SBO_Application.StatusBar.CreateProgressBar("Please Wait", matLABMIX.RowCount, true);
                                oProgressBar.Text = "Please Wait";

                                oForm.Freeze(true);

                                double disperMain = Convert.ToDouble(oForm.Items.Item("disPer").Specific.value);
                                double disamtMain = Convert.ToDouble(oForm.Items.Item("disAmt").Specific.value);
                                double profperMain = Convert.ToDouble(oForm.Items.Item("profPer").Specific.value);
                                double profamtMain = Convert.ToDouble(oForm.Items.Item("profAmt").Specific.value);


                                for (int i = 1; i <= matLABMIX.RowCount; i++)
                                {
                                    Progress += 1;
                                    oProgressBar.Value = Progress;

                                    stprice = Convert.ToDouble(((SAPbouiCOM.EditText)matLABMIX.Columns.Item("stdPrice1").Cells.Item(i).Specific).Value);
                                    exchRate = Convert.ToDouble(oForm.Items.Item("exchRate").Specific.value);

                                    ((EditText)matLABMIX.Columns.Item("freight1").Cells.Item(i).Specific).Value = oForm.Items.Item("freight").Specific.value.ToString();
                                    freight1 = Convert.ToDouble(oForm.Items.Item("freight").Specific.value);

                                    ((EditText)matLABMIX.Columns.Item("packing").Cells.Item(i).Specific).Value = oForm.Items.Item("packingchg").Specific.value.ToString();
                                    packing = Convert.ToDouble(oForm.Items.Item("packingchg").Specific.value);

                                    ((EditText)matLABMIX.Columns.Item("factExp1").Cells.Item(i).Specific).Value = oForm.Items.Item("factExp").Specific.value.ToString();
                                    factExp1 = Convert.ToDouble(oForm.Items.Item("factExp").Specific.value);

                                    ((EditText)matLABMIX.Columns.Item("addexp11").Cells.Item(i).Specific).Value = oForm.Items.Item("addexp1").Specific.value.ToString();
                                    addexp11 = Convert.ToDouble(oForm.Items.Item("addexp1").Specific.value);

                                    ((EditText)matLABMIX.Columns.Item("addexp21").Cells.Item(i).Specific).Value = oForm.Items.Item("addexp2").Specific.value.ToString();
                                    addexp21 = Convert.ToDouble(oForm.Items.Item("addexp2").Specific.value);

                                    ((EditText)matLABMIX.Columns.Item("addexp31").Cells.Item(i).Specific).Value = oForm.Items.Item("addexp3").Specific.value.ToString();
                                    addexp31 = Convert.ToDouble(oForm.Items.Item("addexp3").Specific.value);

                                    total = packing + stprice + freight1 + factExp1 + addexp11 + addexp21 + addexp31;

                                    ((EditText)matLABMIX.Columns.Item("untPrice1").Cells.Item(i).Specific).Value = total.ToString();
                                    //stprice = Convert.ToDouble(((SAPbouiCOM.EditText)matLABMIX.Columns.Item("untPrice1").Cells.Item(i).Specific).Value);

                                    if (disamtMain > 0)
                                    {
                                        disper = (disamtMain * 100) / total;
                                        ((EditText)matLABMIX.Columns.Item("disPer1").Cells.Item(i).Specific).Value = disper.ToString();
                                        ((EditText)matLABMIX.Columns.Item("disAmt1").Cells.Item(i).Specific).Value = disamtMain.ToString();
                                    }
                                    else if (disperMain > 0)
                                    {
                                        disamt = (total * disperMain) / 100;
                                        ((EditText)matLABMIX.Columns.Item("disPer1").Cells.Item(i).Specific).Value = disperMain.ToString();
                                        ((EditText)matLABMIX.Columns.Item("disAmt1").Cells.Item(i).Specific).Value = disamt.ToString();
                                    }

                                    if (profamtMain > 0)
                                    {
                                        profper = (profamtMain * 100) / total;
                                        ((EditText)matLABMIX.Columns.Item("profPer1").Cells.Item(i).Specific).Value = profper.ToString();
                                        ((EditText)matLABMIX.Columns.Item("profAmt1").Cells.Item(i).Specific).Value = profamtMain.ToString();
                                    }
                                    else if (profperMain > 0)
                                    {
                                        profamt = (total * profperMain) / 100;
                                        ((EditText)matLABMIX.Columns.Item("profPer1").Cells.Item(i).Specific).Value = profperMain.ToString();
                                        ((EditText)matLABMIX.Columns.Item("profAmt1").Cells.Item(i).Specific).Value = profamt.ToString();
                                    }

                                    if (profperMain == 0 && profamtMain == 0)
                                    {
                                        profamt = 0;
                                        ((EditText)matLABMIX.Columns.Item("profPer1").Cells.Item(i).Specific).Value = profperMain.ToString();
                                        ((EditText)matLABMIX.Columns.Item("profAmt1").Cells.Item(i).Specific).Value = profamtMain.ToString();
                                    }
                                    if (disperMain == 0 && disamtMain == 0)
                                    {
                                        disamt = 0;
                                        ((EditText)matLABMIX.Columns.Item("disPer1").Cells.Item(i).Specific).Value = disperMain.ToString();
                                        ((EditText)matLABMIX.Columns.Item("disAmt1").Cells.Item(i).Specific).Value = disamtMain.ToString();
                                    }

                                    ((EditText)matLABMIX.Columns.Item("stdPrLC1").Cells.Item(i).Specific).Value = (total - disamt + profamt).ToString();

                                    if (exchRate > 0)
                                    {
                                        double untPrFC1 = (total - disamt + profamt) / exchRate;
                                        FinalPrice = getFCPrice(untPrFC1);
                                        ((EditText)matLABMIX.Columns.Item("untPrFC1").Cells.Item(i).Specific).Value = FinalPrice;
                                    }
                                }
                                oForm.Freeze(false);
                                oProgressBar.Stop();
                            }

                            // Refresh Button FG - LabMix
                            if (pVal.ItemUID == "cmbREFI")
                            {
                                SAPbouiCOM.ButtonCombo cbx = (SAPbouiCOM.ButtonCombo)oForm.Items.Item("cmbREFI").Specific;
                                if (cbx.Selected != null)
                                {
                                    string descrition = cbx.Selected.Description;
                                    string value = cbx.Selected.Value;
                                    if (value == "FG")
                                    {
                                        string getQuery = @"SELECT DISTINCT T1.Code,T1.ItemName,T2.U_price FROM OITT T0
                                                    INNER JOIN ITT1 T1 ON T0.Code = T1.Father
                                                    LEFT JOIN[@GPL1] T2 ON T1.Code = T2.U_itemcode
                                                    LEFT JOIN[@FGPL] T3 ON T2.Code = T3.Code                                                   
                                                    WHERE T3.Code = (SELECT max(DocEntry) FROM dbo.[@FGPL]) AND T0.CODE IN ";

                                        matLABMIX = (SAPbouiCOM.Matrix)oForm.Items.Item("matLABMIX").Specific;
                                        SAPbouiCOM.Matrix matFGITMS = (SAPbouiCOM.Matrix)oForm.Items.Item("matFGITMS").Specific;

                                        string QueryItemCode = string.Empty;
                                        for (int i = 1; i <= matLABMIX.RowCount; i++)
                                        {
                                            string ItemCode = ((SAPbouiCOM.EditText)matLABMIX.Columns.Item("itemcode1").Cells.Item(i).Specific).Value;
                                            if (string.IsNullOrEmpty(QueryItemCode))
                                            {
                                                QueryItemCode = "'" + ItemCode + "',";
                                            }
                                            else
                                            {
                                                QueryItemCode += "'" + ItemCode + "',";
                                            }
                                        }
                                        QueryItemCode = QueryItemCode.Remove(QueryItemCode.Length - 1, 1);

                                        getQuery += @"(" + QueryItemCode + @") AND T1.Code like 'FG%'";

                                        SAPbobsCOM.Recordset rec;
                                        rec = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                        rec.DoQuery(getQuery);

                                        oForm.Freeze(true);
                                        if (rec.RecordCount > 0)
                                        {
                                            DeleteMatrixAllRow(matFGITMS);
                                            ArrengeMatrixLineNum(matFGITMS);
                                            while (!rec.EoF)
                                            {
                                                matFGITMS.AddRow();
                                                ((SAPbouiCOM.EditText)matFGITMS.Columns.Item("#").Cells.Item(matFGITMS.RowCount).Specific).Value = matFGITMS.RowCount.ToString();
                                                ((SAPbouiCOM.EditText)matFGITMS.Columns.Item("itemcode2").Cells.Item(matFGITMS.RowCount).Specific).Value = Convert.ToString(rec.Fields.Item("Code").Value);
                                                ((SAPbouiCOM.EditText)matFGITMS.Columns.Item("itemname2").Cells.Item(matFGITMS.RowCount).Specific).Value = Convert.ToString(rec.Fields.Item("ItemName").Value);
                                                ((SAPbouiCOM.EditText)matFGITMS.Columns.Item("stdprice2").Cells.Item(matFGITMS.RowCount).Specific).Value = Convert.ToString(rec.Fields.Item("U_price").Value);
                                                rec.MoveNext();
                                            }
                                        }
                                        oForm.Freeze(false);
                                        SAPbouiCOM.Folder Tab2 = (SAPbouiCOM.Folder)oForm.Items.Item("tab2").Specific;
                                        Tab2.Select();
                                    }
                                    if (value == "LABMIX")
                                    {
                                        matLABMIX = (SAPbouiCOM.Matrix)oForm.Items.Item("matLABMIX").Specific;
                                        SAPbouiCOM.Matrix matFGITMS = (SAPbouiCOM.Matrix)oForm.Items.Item("matFGITMS").Specific;

                                        oForm.Freeze(true);

                                        if (matLABMIX.RowCount > 0)
                                        {
                                            System.Data.DataTable dt = new System.Data.DataTable();
                                            dt.Columns.Add("FGItem", typeof(string));
                                            dt.Columns.Add("Price", typeof(decimal));
                                            for (int i = 1; i <= matFGITMS.RowCount; i++)
                                            {
                                                System.Data.DataRow dr = dt.NewRow();

                                                string FGItem = ((SAPbouiCOM.EditText)matFGITMS.Columns.Item("itemcode2").Cells.Item(i).Specific).Value;
                                                string Price = ((SAPbouiCOM.EditText)matFGITMS.Columns.Item("stdprice2").Cells.Item(i).Specific).Value;
                                                dr["FGItem"] = FGItem;
                                                dr["Price"] = Price;
                                                dt.Rows.Add(dr);
                                            }

                                            double FinalPricePerLine = 0;
                                            for (int i = 1; i <= matLABMIX.RowCount; i++)
                                            {
                                                string LabItem = ((SAPbouiCOM.EditText)matLABMIX.Columns.Item("itemcode1").Cells.Item(i).Specific).Value;
                                                SAPbobsCOM.ProductTrees oBOM = (SAPbobsCOM.ProductTrees)SBOMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductTrees);

                                                string firstItemcode = null;
                                                string firstItemDesc = null;

                                                if (oBOM.GetByKey(LabItem))
                                                {
                                                    bool isFound = false;
                                                    int BOMItems = oBOM.Items.Count;

                                                    double TotalVal = 0;
                                                    //decimal TotalQty = oBOM.Quantity;
                                                    for (int j = 0; j < BOMItems; j++)
                                                    {


                                                        oBOM.Items.SetCurrentLine(j);

                                                        if (j == 0)
                                                        {
                                                            firstItemcode = oBOM.Items.ItemCode;
                                                            firstItemDesc = oBOM.Items.ItemName;
                                                        }
                                                        string BOMItem = oBOM.Items.ItemCode;
                                                        System.Data.DataRow[] dr = dt.Select("FGItem ='" + oBOM.Items.ItemCode + "' ");

                                                        //if (BOMItem.Contains("FG"))
                                                        //{
                                                        if (dr.Length > 0)
                                                        {
                                                            string test = oBOM.Items.Price.ToString();
                                                            string Strength = Convert.ToString(oBOM.Items.UserFields.Fields.Item("U_strength").Value);
                                                            if (!string.IsNullOrEmpty(Strength))
                                                                oBOM.Items.Price = (Convert.ToDouble(dr[0]["Price"]) * Convert.ToDouble(Strength)) / 100;
                                                            else
                                                                oBOM.Items.Price = Convert.ToDouble(dr[0]["Price"]);
                                                            oBOM.Items.Currency = "INR";

                                                            // firstItemcode =  ;
                                                            // firstItemDesc = ;

                                                            isFound = true;
                                                            TotalVal += Convert.ToDouble(oBOM.Items.Quantity) * Convert.ToDouble(((Convert.ToDouble(dr[0]["Price"]) * Convert.ToDouble(Strength)) / 100));
                                                        }
                                                        else
                                                        {
                                                            TotalVal += Convert.ToDouble(oBOM.Items.Quantity) * Convert.ToDouble(oBOM.Items.Price);
                                                        }

                                                        //}
                                                    }
                                                    FinalPricePerLine = TotalVal / oBOM.Quantity;
                                                    if (isFound)
                                                    {
                                                        int result = oBOM.Update();
                                                        string error1 = string.Empty;
                                                        int iErr1;
                                                        SBOMain.oCompany.GetLastError(out iErr1, out error1);
                                                    }
                                                     ((SAPbouiCOM.EditText)matLABMIX.Columns.Item("fgitemcode").Cells.Item(i).Specific).Value = firstItemcode.ToString();
                                                    ((SAPbouiCOM.EditText)matLABMIX.Columns.Item("fgitemname").Cells.Item(i).Specific).Value = firstItemDesc.ToString();
                                                    ((SAPbouiCOM.EditText)matLABMIX.Columns.Item("stdPrice1").Cells.Item(i).Specific).Value = FinalPricePerLine.ToString();
                                                    ((SAPbouiCOM.EditText)matLABMIX.Columns.Item("untPrice1").Cells.Item(i).Specific).Value = FinalPricePerLine.ToString();
                                                    ((SAPbouiCOM.EditText)matLABMIX.Columns.Item("stdPrLC1").Cells.Item(i).Specific).Value = FinalPricePerLine.ToString();

                                                    double exchRate = Convert.ToDouble(oForm.Items.Item("exchRate").Specific.value);
                                                    if (exchRate > 0)
                                                    {
                                                        double untPrFC1 = FinalPricePerLine / exchRate;
                                                        FinalPrice = getFCPrice(untPrFC1);
                                                        ((EditText)matLABMIX.Columns.Item("untPrFC1").Cells.Item(i).Specific).Value = FinalPrice;
                                                    }
                                                }
                                            }
                                        }
                                        oForm.Freeze(false);
                                        SAPbouiCOM.Folder Tab1 = (SAPbouiCOM.Folder)oForm.Items.Item("tab1").Specific;
                                        Tab1.Select();
                                    }
                                }
                            }

                            // COPY FROM BUTTON COMBO CLICK EVENT
                            if (pVal.ItemUID == "cmbCPYF")
                            {
                                SAPbouiCOM.ButtonCombo cbx = (SAPbouiCOM.ButtonCombo)oForm.Items.Item("cmbCPYF").Specific;
                                if (cbx.Selected != null)
                                {
                                    string descrition = cbx.Selected.Description;
                                    string value = cbx.Selected.Value;
                                    if (value == "outward")
                                    {
                                        string cardcode = oForm.Items.Item("cardcode").Specific.value;
                                        string cardname = oForm.Items.Item("cardname").Specific.value;
                                        string fromdate = DateTime.Today.AddYears(-1).ToString("yyyyMMdd");
                                        string todate = DateTime.Today.ToString("yyyyMMdd");

                                        PriceListToOutWord PriceToo = new PriceListToOutWord();
                                        PriceToo.BPCode = cardcode;
                                        PriceToo.BPName = cardname;

                                        SBOMain.SBO_Application.Menus.Item("mnsmPL001").Activate();
                                        clsOutwards oPrice = new clsOutwards(PriceToo);
                                    }
                                }

                            }

                            // COPY TO BUTTON COMBO CLICK EVENT
                            if (pVal.ItemUID == "cmbCPYT")
                            {
                                SAPbouiCOM.ButtonCombo cbx = (SAPbouiCOM.ButtonCombo)oForm.Items.Item("cmbCPYT").Specific;
                                if (cbx.Selected != null)
                                {
                                    string descrition = cbx.Selected.Description;
                                    string value = cbx.Selected.Value;

                                    Header oHeader = new Header();
                                    oHeader.BPCode = oForm.Items.Item("cardcode").Specific.value;
                                    oHeader.BPName = oForm.Items.Item("cardname").Specific.value;
                                    oHeader.PRLNum = oForm.Items.Item("tCode").Specific.value;

                                    SAPbouiCOM.ComboBox cb4 = oForm.Items.Item("cmbCRCY").Specific;
                                    oHeader.currency = cb4.Selected.Value.ToString();
                                    List<Child> lstChild = new List<Child>();

                                    for (int i = 1; i <= matLABMIX.RowCount; i++)
                                    {
                                        Child oChild = new Child();
                                        oChild.ItemCode = Convert.ToString(((SAPbouiCOM.EditText)matLABMIX.Columns.Item("fgitemcode").Cells.Item(i).Specific).Value);
                                        oChild.ItemName = Convert.ToString(((SAPbouiCOM.EditText)matLABMIX.Columns.Item("fgitemname").Cells.Item(i).Specific).Value);
                                        oChild.labmixno = Convert.ToString(((SAPbouiCOM.EditText)matLABMIX.Columns.Item("itemcode1").Cells.Item(i).Specific).Value);
                                        oChild.labmixname = Convert.ToString(((SAPbouiCOM.EditText)matLABMIX.Columns.Item("itemname1").Cells.Item(i).Specific).Value);
                                        oChild.partrefno = Convert.ToString(((SAPbouiCOM.EditText)matLABMIX.Columns.Item("ptyref1").Cells.Item(i).Specific).Value);
                                        oChild.UnitPrice = Convert.ToDouble(((SAPbouiCOM.EditText)matLABMIX.Columns.Item("untPrice1").Cells.Item(i).Specific).Value);
                                        oChild.totalLC = ((SAPbouiCOM.EditText)matLABMIX.Columns.Item("stdPrLC1").Cells.Item(i).Specific).Value;
                                        oChild.totalFC = ((SAPbouiCOM.EditText)matLABMIX.Columns.Item("untPrFC1").Cells.Item(i).Specific).Value;

                                        lstChild.Add(oChild);
                                    }
                                    oHeader.lstChild = lstChild;

                                    if (value == "squote")
                                    {
                                        SBOMain.SBO_Application.Menus.Item("2049").Activate();
                                    }
                                    else if (value == "sorder")
                                    {
                                        SBOMain.SBO_Application.Menus.Item("2050").Activate();
                                    }
                                    oForm = SBOMain.SBO_Application.Forms.ActiveForm;
                                    oForm.Items.Item("4").Specific.value = oHeader.BPCode;
                                    oForm.Items.Item("54").Specific.value = oHeader.BPName;
                                    SAPbouiCOM.ComboBox cb6 = oForm.Items.Item("63").Specific;
                                    cb6.Select(oHeader.currency);

                                    SAPbouiCOM.Form oUDFForm = SBOMain.SBO_Application.Forms.Item(oForm.UDFFormUID);
                                    oUDFForm.Items.Item("U_PRLNum").Specific.value = oHeader.PRLNum;
                                    SAPbouiCOM.Matrix matQt = (SAPbouiCOM.Matrix)oForm.Items.Item("38").Specific;

                                    for (int i = 0; i < oHeader.lstChild.Count; i++)
                                    {
                                        ((SAPbouiCOM.EditText)matQt.Columns.Item("1").Cells.Item(i + 1).Specific).Value = oHeader.lstChild[i].ItemCode;
                                        ((SAPbouiCOM.EditText)matQt.Columns.Item("3").Cells.Item(i + 1).Specific).Value = oHeader.lstChild[i].ItemName;
                                        ((SAPbouiCOM.EditText)matQt.Columns.Item("256").Cells.Item(i + 1).Specific).Value = oHeader.lstChild[i].labmixname;
                                        ((SAPbouiCOM.EditText)matQt.Columns.Item("U_UNE_BTNO").Cells.Item(i + 1).Specific).Value = oHeader.lstChild[i].partrefno;
                                        ((SAPbouiCOM.EditText)matQt.Columns.Item("U_Lab_Mix").Cells.Item(i + 1).Specific).Value = oHeader.lstChild[i].labmixno;
                                        ((SAPbouiCOM.EditText)matQt.Columns.Item("14").Cells.Item(i + 1).Specific).Value = oHeader.lstChild[i].totalFC.ToString();
                                    }
                                }
                            }
                        }
                        break;
                    case BoEventTypes.et_ITEM_PRESSED:
                        oForm = SBOMain.SBO_Application.Forms.ActiveForm;
                        if (pVal.BeforeAction == true)
                        {
                            if (pVal.ItemUID == "1" && (oForm.Mode == BoFormMode.fm_ADD_MODE || oForm.Mode == BoFormMode.fm_UPDATE_MODE))
                            {
                                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("matLABMIX").Specific;
                                DeleteMatrixBlankRow(oMatrix);

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
                            }
                        }
                        if (pVal.BeforeAction == false)
                        {
                            if (pVal.ItemUID == "1" && oForm.Mode == BoFormMode.fm_ADD_MODE)
                            {
                                /*if (pVal.ItemUID == "1" && oForm.Mode == BoFormMode.fm_ADD_MODE)
                                {
                                    Form_Load_Components(oForm, "OK");
                                }*/
                                Form_Load_Components(oForm, "OK");
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
                /* if (oForm != null)
                     oForm.Freeze(false);*/
            }  
            return BubbleEvent;
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
        
        public string getFCPrice(double price)
        {
            string finalresult;
            string lastChar;

            string twoDigitRoundI = price.ToString("0.00");
            double twoDigitRoundD = Convert.ToDouble(twoDigitRoundI);

            //string valAfterDot = twoDigitRoundI.Split(Convert.ToChar("."));

            int diff = (int)(Math.Round(((twoDigitRoundD - (int)twoDigitRoundD) * 100), 2));

            if (diff.ToString().Length > 1) { lastChar = diff.ToString().Substring(1, 1); } else { lastChar = diff.ToString(); }

            if (Convert.ToInt32(lastChar) == 1 || Convert.ToInt32(lastChar) == 6) { twoDigitRoundD = twoDigitRoundD - 0.01; }
            if (Convert.ToInt32(lastChar) == 2 || Convert.ToInt32(lastChar) == 7) { twoDigitRoundD = twoDigitRoundD - 0.02; }
            if (Convert.ToInt32(lastChar) == 3 || Convert.ToInt32(lastChar) == 8) { twoDigitRoundD = twoDigitRoundD + 0.02; }
            if (Convert.ToInt32(lastChar) == 4 || Convert.ToInt32(lastChar) == 9) { twoDigitRoundD = twoDigitRoundD + 0.01; }

            finalresult = twoDigitRoundD.ToString();
            return finalresult;
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

        private void ArrengeMatrixLineNumLM(SAPbouiCOM.Matrix matrix)
        {
            SAPbouiCOM.Matrix matLab = oForm.Items.Item("matLABMIX").Specific;
            for (int i = 1; i <= matLab.VisualRowCount; i++)
            {
                matLab.Columns.Item("#").Cells.Item(i).Specific.value = i;
            }
        }
        
        private void ArrengeMatrixLineNum(SAPbouiCOM.Matrix matrix)
        {
            for (int i = 1; i <= matrix.RowCount; i++)
            {
                matrix.Columns.Item("#").Cells.Item(i).Specific.value = i;
            }
        }

        private void DeleteMatrixAllRow(SAPbouiCOM.Matrix oMatrix)
        {
            try
            {
                for (int i = oMatrix.VisualRowCount; i >= 1; i--)
                {
                    {
                        oMatrix.DeleteRow(i);
                    }
                }
            }
            catch (Exception ex)
            {

            }
        }

        private void DeleteMatrixBlankRow(SAPbouiCOM.Matrix oMatrix)
        {
            try
            {
                for (int i = oMatrix.VisualRowCount; i >= 1; i--)
                {
                    {
                        if (string.IsNullOrEmpty(((SAPbouiCOM.EditText)oMatrix.Columns.Item("itemcode1").Cells.Item(i).Specific).Value))
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
                
        public void Form_Load_Components(SAPbouiCOM.Form oForm, string mode)
        {
            try
            {
                oForm.Items.Item("tab1").Visible = true;
                oForm.Items.Item("tab1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                oForm.PaneLevel = 1;

                if (mode != "OK")
                { 
                    SetCode();

                SAPbouiCOM.ComboBox cb = (SAPbouiCOM.ComboBox)oForm.Items.Item("cmbStatus").Specific;
                cb.ExpandType = BoExpandType.et_DescriptionOnly;
                cb.Select("O");

                oForm.Items.Item("docDate").Specific.Value = DateTime.Today.ToString("yyyyMMdd");

                SAPbouiCOM.ButtonCombo cb1 = (SAPbouiCOM.ButtonCombo)oForm.Items.Item("cmbCPYT").Specific;
                cb1.ValidValues.Add("squote", "Sales Quotation");
                cb1.ValidValues.Add("sorder", "Sales Order");
                cb1.ExpandType = BoExpandType.et_DescriptionOnly;

                SAPbouiCOM.ButtonCombo cb2 = (SAPbouiCOM.ButtonCombo)oForm.Items.Item("cmbCPYF").Specific;
                cb2.ValidValues.Add("Inward", "Inward");
                cb2.ValidValues.Add("outward", "Outward");
                cb2.ExpandType = BoExpandType.et_DescriptionOnly;

                SAPbouiCOM.ButtonCombo cb3 = (SAPbouiCOM.ButtonCombo)oForm.Items.Item("cmbREFI").Specific;
                cb3.ValidValues.Add("FG", "FG");
                cb3.ValidValues.Add("LABMIX", "LABMIX");
                cb3.ExpandType = BoExpandType.et_DescriptionOnly;

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
                    oForm.Items.Item("cardcode").Click();
                }
            }
            catch (Exception ex)
            {
                //SBO_Application.StatusBar.SetText("(1): " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }
        
        private void SetCode()
        {
            oForm.Freeze(true);
            oForm.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized;
            string TableName = "SMPL";
            SBOMain.SetCode(oForm.UniqueID, TableName);
            oForm.Freeze(false);
        }
       
        private void CFLCondition(string CFLID, string ItemUID)
        {
            oCFL = oForm.ChooseFromLists.Item(CFLID);
            oConds = SBOMain.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions);

            if (CFLID == "CFL_OCRD")
            {
                oCond = oConds.Add();
                oCond.Alias = "validFor";
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCond.CondVal = "Y";

                oCFL.SetConditions(oConds);
            }

            if (CFLID == "CFL_FGItem")
            {
                oCond = oConds.Add();
                oCond.Alias = "ItemCode";
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_START;
                oCond.CondVal = "FG";
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

    }

    public class PriceListToOutWord
    {
        public string BPCode { get; set; }
        public string BPName { get; set; }

    }

    public class Header
    {
        public string PRLNum { get; set; }
        public string BPCode { get; set; }
        public string BPName { get; set; }
        public string currency { get; set; }
        public List<Child> lstChild { get; set; }
    }
    public class Child
    {
        public string ItemCode { get; set; }
        public string ItemName { get; set; }
        public double Quantity { get; set; }
        public double UnitPrice { get; set; }
        public string labmixno { get; set; }
        public string labmixname { get; set; }
        public string partrefno { get; set; }
        public string totalLC { get; set; }
        public string totalFC { get; set; }

    }
    public class FormHeader
    {
        public string BPCode { get; set; }
        public double freight { get; set; }
        public double packingchg { get; set; }
        public double factExp { get; set; }
        public double addexp1 { get; set; }
        public double addexp2 { get; set; }
        public double addexp3 { get; set; }
        public double exchRate { get; set; }
        public double disPer { get; set; }
        public double disAmt { get; set; }
        public double profPer { get; set; }
        public double profAmt { get; set; }
        public string remarks { get; set; }
        public string currency { get; set; }
        public List<FormChild_Lab> frmChild_Lab { get; set; }
        public List<FormChild_FG> frmChild_FG { get; set; }

    }
    public class FormChild_Lab
    {
        public string itemcode1 { get; set; }
        public string itemname1 { get; set; }
        public string fgitemcode { get; set; }
        public string fgitemname { get; set; }
        public string ioref1 { get; set; }
        public string inoutno { get; set; }
        public string ptyref1 { get; set; }
        public string stdPrice1 { get; set; }
        public string freight1 { get; set; }
        public string packing { get; set; }
        public string factExp1 { get; set; }
        public string addexp11 { get; set; }
        public string addexp21 { get; set; }
        public string addexp31 { get; set; }
        public string untPrice1 { get; set; }
        public string disPer1 { get; set; }
        public string disAmt1 { get; set; }
        public string profPer1 { get; set; }
        public string profAmt1 { get; set; }
        public string stdPrLC1 { get; set; }
        public string untPrFC1 { get; set; }


    }

    public class FormChild_FG
    {
        public string itemcode2 { get; set; }
        public string itemname2 { get; set; }
        public string stdprice2 { get; set; }
        public string updprice2 { get; set; }
        public string remarks2 { get; set; }
    }
}
