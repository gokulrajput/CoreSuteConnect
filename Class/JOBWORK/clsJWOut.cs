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
using CoreSuteConnect.Class.EXIM;
using System.Text.RegularExpressions;
using System.Data;
using System.Drawing.Drawing2D;
using System.Security.Cryptography;
using CoreSuteConnect.Class.PRICELIST;

namespace CoreSuteConnect.Class.JOBWORK
{
    class clsJWOut
    {
        #region VariableDeclaration

        private SAPbouiCOM.Form oForm;
        private SAPbouiCOM.Matrix oMatrix;
        public string cFormID = string.Empty;

        private SAPbouiCOM.ChooseFromList oCFL = null;
        private SAPbouiCOM.Conditions oConds = null;
        private SAPbouiCOM.Condition oCond = null;

        double Totalqty = 0;
        double issuedqty = 0;
        double balqty = 0;
        SAPbouiCOM.EditText oLinkSOPO = null;
        SAPbouiCOM.EditText DocLink = null;

        CommonUtility objCU = new CommonUtility();
        #endregion VariableDeclaration

        public clsJWOut(OutwardToJWO outClass)
        {
            try
            {
                if (outClass != null)
                {
                    oForm = SBOMain.SBO_Application.Forms.ActiveForm;
                    if (outClass.FromFrmName == "POExist")
                    {
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                        oForm.Items.Item("tCode").Specific.value = outClass.DocEntry;
                        oForm.Items.Item("1").Click();
                    }
                    else if (outClass.FromFrmName == "PurchaseOrder")
                    { 
                        oForm.Items.Item("tCardCode").Specific.value = outClass.BPCode;
                        oForm.Items.Item("tCardName").Specific.value = outClass.BPName;
                        oForm.Items.Item("tPoNum").Specific.value = outClass.DocNum;
                        oForm.Items.Item("tPoDe").Specific.value = outClass.DocEntry;
                        oForm.Items.Item("tNumAtCard").Specific.value = outClass.VendorRef;
                        oForm.Items.Item("tWhsCode").Specific.value = outClass.DefWhs;

                        if (outClass.lstItems.Count > 0)
                        {
                            for (int i = 0; i < outClass.lstItems.Count; i++)
                            {
                                SAPbouiCOM.Matrix lineMatrix = oForm.Items.Item("matFGTR").Specific;
                                lineMatrix.AddRow();
                                ((SAPbouiCOM.EditText)lineMatrix.Columns.Item("#").Cells.Item(lineMatrix.RowCount).Specific).Value = Convert.ToString(lineMatrix.RowCount);
                                ((SAPbouiCOM.EditText)lineMatrix.Columns.Item("tItemCode").Cells.Item(lineMatrix.RowCount).Specific).Value = outClass.lstItems[i].ItemCode;
                                ((SAPbouiCOM.EditText)lineMatrix.Columns.Item("tDsc").Cells.Item(lineMatrix.RowCount).Specific).Value = outClass.lstItems[i].ItemName;
                                ((SAPbouiCOM.EditText)lineMatrix.Columns.Item("tQuantity").Cells.Item(lineMatrix.RowCount).Specific).Value = outClass.lstItems[i].Quantity;
                                ((SAPbouiCOM.EditText)lineMatrix.Columns.Item("tUOM").Cells.Item(lineMatrix.RowCount).Specific).Value = outClass.lstItems[i].UOM;
                                ((SAPbouiCOM.EditText)lineMatrix.Columns.Item("tPrice").Cells.Item(lineMatrix.RowCount).Specific).Value = outClass.lstItems[i].UnitPrice;
                                ((SAPbouiCOM.EditText)lineMatrix.Columns.Item("tBalQty").Cells.Item(lineMatrix.RowCount).Specific).Value = outClass.lstItems[i].Quantity;
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

                    }
                }
                if (pVal.BeforeAction == false)
                {

                    if ((oForm.Mode == BoFormMode.fm_ADD_MODE || oForm.Mode == BoFormMode.fm_UPDATE_MODE || Type == "ADDNEWFORM") && (Type != "DEL_ROW") && (Type != "ADD_ROW"))
                    {
                        Form_Load_Components(oForm, "Add");
                    }
                    else if (Type == "DEL_ROW" || Type == "ADD_ROW")
                    {
                        SAPbouiCOM.Matrix matFGTR = (SAPbouiCOM.Matrix)oForm.Items.Item("matFGTR").Specific;
                        SAPbouiCOM.Matrix matCMTR = (SAPbouiCOM.Matrix)oForm.Items.Item("matCMTR").Specific;
                        SAPbouiCOM.Matrix matLDFG = (SAPbouiCOM.Matrix)oForm.Items.Item("matLDFG").Specific;
                        
                        if (Type == "ADD_ROW")
                        {
                            if (SBOMain.RightClickItemID == "matFGTR")
                            {
                                ADDROWMain(matFGTR);
                            }
                            else if (SBOMain.RightClickItemID == "matCMTR")
                            {
                                ADDROWMain(matCMTR);
                            }
                            else if (SBOMain.RightClickItemID == "matLDFG")
                            {
                                ADDROWMain(matLDFG);
                            }
                        }
                        if (Type == "DEL_ROW")
                        {
                            if (SBOMain.RightClickItemID == "matFGTR")
                            {
                                DeleteMatrixBlankRow(matFGTR, "tPoNum");
                                ArrengeMatrixLineNum(matFGTR);
                            }
                            else if (SBOMain.RightClickItemID == "matCMTR")
                            {
                                DeleteMatrixBlankRow(matCMTR, "tPoNum");
                                ArrengeMatrixLineNum(matCMTR);
                            }
                            else if (SBOMain.RightClickItemID == "matLDFG")
                            {
                                DeleteMatrixBlankRow(matLDFG, "tTransType");
                                ArrengeMatrixLineNum(matLDFG);
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
            try
            {
                switch (pVal.EventType)
                {

                    case BoEventTypes.et_COMBO_SELECT:
                        oForm = SBOMain.SBO_Application.Forms.ActiveForm;
                        if (pVal.BeforeAction == false)
                        {
                            if (pVal.ItemUID == "cSer" && pVal.FormMode == 3)
                            {
                                oForm.Items.Item("tDocNum").Specific.Value = oForm.BusinessObject.GetNextSerialNumber(oForm.Items.Item("cSer").Specific.Value, "JOTR");
                            }
                        }

                        break;
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
                        if (pVal.BeforeAction == true)
                        {

                        }
                        if (pVal.BeforeAction == false)
                        {

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
                                try {

                                    if (pVal.ItemUID == "matCMTR" && pVal.ColUID == "tQuantity")
                                    {
                                        SAPbouiCOM.Matrix Omatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("matCMTR").Specific; 
                                        Totalqty = Convert.ToDouble(((SAPbouiCOM.EditText)Omatrix.Columns.Item("tQuantity").Cells.Item(pVal.Row).Specific).Value);
                                        issuedqty = Convert.ToDouble(((SAPbouiCOM.EditText)Omatrix.Columns.Item("tIsueQty").Cells.Item(pVal.Row).Specific).Value);
                                        balqty = Totalqty - issuedqty;
                                        ((SAPbouiCOM.EditText)Omatrix.Columns.Item("tBalQty").Cells.Item(pVal.Row).Specific).Value = balqty.ToString(); 
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
                            string bpcode = oForm.Items.Item("tCardCode").Specific.value;

                            if (pVal.ItemUID == "tCardCode")
                            {
                                CFLCondition("CFL_OCRD", pVal.ItemUID, bpcode);
                            }
                            else if (pVal.ItemUID == "tnopcode")
                            {
                                CFLCondition("CFL_JOPM", pVal.ItemUID, bpcode);
                            }
                            else if (pVal.ItemUID == "tPoNum")
                            {
                                CFLCondition("CFL_OPOR", pVal.ItemUID, bpcode);
                            }
                            else if (pVal.ItemUID == "matFGTR" && pVal.ColUID == "tItemCode")
                            {
                                CFLCondition("CFL_OITM", pVal.ItemUID, bpcode);
                            }
                            else if (pVal.ItemUID == "matCMTR" && pVal.ColUID == "tItemCode")
                            {
                                CFLCondition("CFL_OITM2", pVal.ItemUID, bpcode);
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
                                    if (pVal.ItemUID == "tnopcode")
                                    {
                                        oForm.Items.Item("tnopcode").Specific.value = oDataTable.GetValue("U_nopcode", 0).ToString();
                                    }
                                    else if (pVal.ItemUID == "tCardCode")
                                    {   
                                        oForm.Items.Item("tCardCode").Specific.value = oDataTable.GetValue("CardCode", 0).ToString();
                                        oForm.Items.Item("tCardName").Specific.value = oDataTable.GetValue("CardName", 0).ToString();
                                        oForm.Items.Item("tWhsCode").Specific.value = oDataTable.GetValue("U_DefWhs", 0).ToString();
                                    }
                                    else if (pVal.ItemUID == "tPoNum")
                                    {   
                                        oForm.Items.Item("tPoNum").Specific.value = oDataTable.GetValue("DocNum", 0).ToString();
                                        oForm.Items.Item("tPoDe").Specific.value = oDataTable.GetValue("DocEntry", 0).ToString();
                                        oForm.Items.Item("tNumAtCard").Specific.value = oDataTable.GetValue("NumAtCard", 0).ToString();
                                            
                                        DateTime tDocDate = Convert.ToDateTime(oDataTable.GetValue("DocDate", 0).ToString());
                                        oForm.Items.Item("tDocDate").Specific.value = tDocDate.ToString("yyyyMMdd");
                                            
                                        DateTime lcsd1 = Convert.ToDateTime(oDataTable.GetValue("DocDueDate", 0).ToString());
                                        oForm.Items.Item("tDocDDt").Specific.value = lcsd1.ToString("yyyyMMdd");
                                    }
                                    else if (pVal.ItemUID == "matFGTR" && pVal.ColUID == "tItemCode")
                                    {
                                        SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("matFGTR").Specific;
                                        ((SAPbouiCOM.EditText)oMatrix.Columns.Item("tItemCode").Cells.Item(pVal.Row).Specific).Value = oDataTable.GetValue("ItemCode", 0).ToString();
                                        ((SAPbouiCOM.EditText)oMatrix.Columns.Item("tDsc").Cells.Item(pVal.Row).Specific).Value = oDataTable.GetValue("ItemName", 0).ToString();
                                        ((SAPbouiCOM.EditText)oMatrix.Columns.Item("tWhsCode").Cells.Item(pVal.Row).Specific).Value = oDataTable.GetValue("DfltWH", 0).ToString();
                                        ((SAPbouiCOM.EditText)oMatrix.Columns.Item("tUOM").Cells.Item(pVal.Row).Specific).Value = oDataTable.GetValue("SalUnitMsr", 0).ToString();
                                          
                                    }
                                    else if (pVal.ItemUID == "matCMTR" && pVal.ColUID == "tItemCode")
                                    {
                                        SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("matCMTR").Specific;
                                        ((SAPbouiCOM.EditText)oMatrix.Columns.Item("tItemCode").Cells.Item(pVal.Row).Specific).Value = oDataTable.GetValue("ItemCode", 0).ToString();
                                        ((SAPbouiCOM.EditText)oMatrix.Columns.Item("tDscr").Cells.Item(pVal.Row).Specific).Value = oDataTable.GetValue("ItemName", 0).ToString();
                                        ((SAPbouiCOM.EditText)oMatrix.Columns.Item("tUOM").Cells.Item(pVal.Row).Specific).Value = oDataTable.GetValue("BuyUnitMsr", 0).ToString();
                                        ((SAPbouiCOM.EditText)oMatrix.Columns.Item("tFrmWhs").Cells.Item(pVal.Row).Specific).Value = oDataTable.GetValue("DfltWH", 0).ToString();
                                        ((SAPbouiCOM.EditText)oMatrix.Columns.Item("tPrice").Cells.Item(pVal.Row).Specific).Value = oDataTable.GetValue("LastPurPrc", 0).ToString();
                                        ((SAPbouiCOM.EditText)oMatrix.Columns.Item("tToWhs").Cells.Item(pVal.Row).Specific).Value = oForm.Items.Item("tWhsCode").Specific.value.ToString();
                                          
                                    } 
                                }
                                catch (Exception ex)
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
                        SAPbouiCOM.Matrix matCMTR = (SAPbouiCOM.Matrix)oForm.Items.Item("matCMTR").Specific;
                        SAPbouiCOM.Matrix matFGTR = (SAPbouiCOM.Matrix)oForm.Items.Item("matFGTR").Specific;
                         
                        if (pVal.BeforeAction == true)
                        {
                            // Validation : Without Add / Update not allow to perform Copy to.
                            if (pVal.ItemUID == "cmbCPYT")
                            {
                                //if (oForm.Mode == BoFormMode.fm_ADD_MODE || oForm.Mode == BoFormMode.fm_UPDATE_MODE)
                                if (oForm.Mode == BoFormMode.fm_ADD_MODE)
                                {
                                    // SBOMain.SBO_Application.StatusBar.SetText("Please first Add the form.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                    //BubbleEvent = false;
                                }
                            }
                        }
                        if (pVal.BeforeAction == false)
                        { 
                            // COPY TO BUTTON COMBO CLICK EVENT
                            if (pVal.ItemUID == "cmbCPYT")
                            {
                                SAPbouiCOM.ButtonCombo cbx = (SAPbouiCOM.ButtonCombo)oForm.Items.Item("cmbCPYT").Specific;
                                if (cbx.Selected != null)
                                {
                                    //GETGL Acnt
                                    string glact = objCU.GetJobWorkOutAccount();

                                    string descrition = cbx.Selected.Description;
                                    string value = cbx.Selected.Value;

                                    HeaderJWO oHeader = new HeaderJWO();
                                    oHeader.BPCode = oForm.Items.Item("tCardCode").Specific.value;
                                    oHeader.BPName = oForm.Items.Item("tCardName").Specific.value;
                                    oHeader.PoNum = oForm.Items.Item("tPoNum").Specific.value;
                                    oHeader.PoDe = oForm.Items.Item("tPoDe").Specific.value;
                                    oHeader.NumAtCard = oForm.Items.Item("tNumAtCard").Specific.value;
                                    oHeader.DocDate = oForm.Items.Item("tDocDate").Specific.value;
                                    oHeader.DocDDt = oForm.Items.Item("tDocDDt").Specific.value;
                                    oHeader.Whs = oForm.Items.Item("tWhsCode").Specific.value;
                                    oHeader.JWODE = oForm.Items.Item("tCode").Specific.value;
                                     
                                    SAPbouiCOM.ComboBox cb4 = oForm.Items.Item("tBPLId").Specific;
                                    oHeader.BPLId = cb4.Selected.Value.ToString();


                                    List<ChildJWO> lstChild = new List<ChildJWO>(); 
                                     
                                    if (value == "GI")
                                    {
                                        for (int i = 1; i <= matCMTR.RowCount; i++)
                                        {
                                            ChildJWO oChild = new ChildJWO();
                                            if (((SAPbouiCOM.CheckBox)matCMTR.Columns.Item("chkTrans").Cells.Item(i).Specific).Checked)
                                            {
                                                oChild.ItemCode = Convert.ToString(((SAPbouiCOM.EditText)matCMTR.Columns.Item("tItemCode").Cells.Item(i).Specific).Value);
                                                oChild.ItemName = Convert.ToString(((SAPbouiCOM.EditText)matCMTR.Columns.Item("tDscr").Cells.Item(i).Specific).Value);
                                                oChild.Quantity = Convert.ToDecimal(((SAPbouiCOM.EditText)matCMTR.Columns.Item("tQuantity").Cells.Item(i).Specific).Value);
                                                oChild.BalQty = Convert.ToDecimal(((SAPbouiCOM.EditText)matCMTR.Columns.Item("tBalQty").Cells.Item(i).Specific).Value);
                                                oChild.WhsFrm = Convert.ToString(((SAPbouiCOM.EditText)matCMTR.Columns.Item("tFrmWhs").Cells.Item(i).Specific).Value);
                                                oChild.whsTo = Convert.ToString(((SAPbouiCOM.EditText)matCMTR.Columns.Item("tToWhs").Cells.Item(i).Specific).Value);
                                                lstChild.Add(oChild);
                                            }
                                        }
                                        oHeader.lstChild = lstChild;

                                        SBOMain.SBO_Application.Menus.Item("3079").Activate();

                                        oForm = SBOMain.SBO_Application.Forms.ActiveForm;
                                        oForm.Items.Item("U_BP_Name").Specific.value = oHeader.BPCode;

                                        SAPbouiCOM.Form oUDFForm = SBOMain.SBO_Application.Forms.Item(oForm.UDFFormUID);
                                        oUDFForm.Items.Item("U_JWODe").Specific.value = oHeader.JWODE; 
                                        
                                        SAPbouiCOM.Matrix matGI = (SAPbouiCOM.Matrix)oForm.Items.Item("13").Specific;
                                        int rowNum = 1;
                                        for (int i = 0; i < oHeader.lstChild.Count; i++)
                                        {
                                            ((SAPbouiCOM.EditText)matGI.Columns.Item("1").Cells.Item(rowNum).Specific).Value = oHeader.lstChild[i].ItemCode;
                                            ((SAPbouiCOM.EditText)matGI.Columns.Item("2").Cells.Item(rowNum).Specific).Value = oHeader.lstChild[i].ItemName;
                                            ((SAPbouiCOM.EditText)matGI.Columns.Item("9").Cells.Item(rowNum).Specific).Value = Convert.ToString(oHeader.lstChild[i].BalQty);
                                            ((SAPbouiCOM.EditText)matGI.Columns.Item("15").Cells.Item(rowNum).Specific).Value = oHeader.lstChild[i].whsTo;
                                            ((SAPbouiCOM.EditText)matGI.Columns.Item("59").Cells.Item(rowNum).Specific).Value = glact;
                                            rowNum++;
                                        }
                                    }
                                    else if (value == "GR")
                                    {
                                        for (int i = 1; i <= matFGTR.RowCount; i++)
                                        {
                                            ChildJWO oChild = new ChildJWO(); 
                                            oChild.ItemCode = Convert.ToString(((SAPbouiCOM.EditText)matFGTR.Columns.Item("tItemCode").Cells.Item(i).Specific).Value);
                                            oChild.ItemName = Convert.ToString(((SAPbouiCOM.EditText)matFGTR.Columns.Item("tDsc").Cells.Item(i).Specific).Value);
                                            oChild.Quantity = Convert.ToDecimal(((SAPbouiCOM.EditText)matFGTR.Columns.Item("tQuantity").Cells.Item(i).Specific).Value);
                                            oChild.WhsFrm = Convert.ToString(((SAPbouiCOM.EditText)matFGTR.Columns.Item("tWhsCode").Cells.Item(i).Specific).Value);
                                            lstChild.Add(oChild); 
                                        }
                                        oHeader.lstChild = lstChild; 
                                         
                                        SBOMain.SBO_Application.Menus.Item("3078").Activate();
                                        oForm = SBOMain.SBO_Application.Forms.ActiveForm;
                                        oForm.Items.Item("U_BP_Name").Specific.value = oHeader.BPCode;

                                        SAPbouiCOM.Form oUDFForm = SBOMain.SBO_Application.Forms.Item(oForm.UDFFormUID);
                                        oUDFForm.Items.Item("U_JWODe").Specific.value = oHeader.JWODE; 
                                        
                                        SAPbouiCOM.Matrix matGR = (SAPbouiCOM.Matrix)oForm.Items.Item("13").Specific;
                                        int rowNum = 1;
                                        for (int i = 0; i < oHeader.lstChild.Count; i++)
                                        {
                                            ((SAPbouiCOM.EditText)matGR.Columns.Item("1").Cells.Item(rowNum).Specific).Value = oHeader.lstChild[i].ItemCode;
                                            ((SAPbouiCOM.EditText)matGR.Columns.Item("2").Cells.Item(rowNum).Specific).Value = oHeader.lstChild[i].ItemName;
                                            ((SAPbouiCOM.EditText)matGR.Columns.Item("9").Cells.Item(rowNum).Specific).Value = Convert.ToString(oHeader.lstChild[i].Quantity);
                                            ((SAPbouiCOM.EditText)matGR.Columns.Item("59").Cells.Item(rowNum).Specific).Value = glact;
                                            ((SAPbouiCOM.EditText)matGR.Columns.Item("15").Cells.Item(rowNum).Specific).Value = oHeader.lstChild[i].WhsFrm;
                                           
                                            //((SAPbouiCOM.EditText)matGR.Columns.Item("10").Cells.Item(rowNum).Specific).Value = Convert.ToString(Unitprice);
                                           // ((SAPbouiCOM.EditText)matGR.Columns.Item("U_Price_CSJW").Cells.Item(rowNum).Specific).Value = Convert.ToString(Unitprice);
                                            //matGI.AddRow();
                                            //matGR.ClearRowData(matGR.RowCount);
                                            rowNum++;
                                        }
                                    }
                                    else if (value == "SR")
                                    {
                                        SBOMain.SBO_Application.Menus.Item("2050").Activate();
                                    }
                                    else if (value == "IT")
                                    {
                                        for (int i = 1; i <= matCMTR.RowCount; i++)
                                        {
                                            ChildJWO oChild = new ChildJWO();
                                            if (((SAPbouiCOM.CheckBox)matCMTR.Columns.Item("chkTrans").Cells.Item(i).Specific).Checked)
                                            {
                                                oChild.ItemCode = Convert.ToString(((SAPbouiCOM.EditText)matCMTR.Columns.Item("tItemCode").Cells.Item(i).Specific).Value);
                                                oChild.ItemName = Convert.ToString(((SAPbouiCOM.EditText)matCMTR.Columns.Item("tDscr").Cells.Item(i).Specific).Value);
                                                oChild.Quantity = Convert.ToDecimal(((SAPbouiCOM.EditText)matCMTR.Columns.Item("tQuantity").Cells.Item(i).Specific).Value);
                                                oChild.BalQty = Convert.ToDecimal(((SAPbouiCOM.EditText)matCMTR.Columns.Item("tBalQty").Cells.Item(i).Specific).Value);
                                                oChild.WhsFrm = Convert.ToString(((SAPbouiCOM.EditText)matCMTR.Columns.Item("tFrmWhs").Cells.Item(i).Specific).Value);
                                                oChild.whsTo = Convert.ToString(((SAPbouiCOM.EditText)matCMTR.Columns.Item("tToWhs").Cells.Item(i).Specific).Value);
                                                lstChild.Add(oChild);
                                            }
                                        }
                                        oHeader.lstChild = lstChild;

                                        SBOMain.SBO_Application.Menus.Item("3080").Activate();
                                        oForm = SBOMain.SBO_Application.Forms.ActiveForm;
                                        oForm.Items.Item("3").Specific.value = oHeader.BPCode;
                                        oForm.Items.Item("7").Specific.value = oHeader.BPName;

                                        SAPbouiCOM.Form oUDFForm = SBOMain.SBO_Application.Forms.Item(oForm.UDFFormUID); 
                                        oUDFForm.Items.Item("U_JWODe").Specific.value = oHeader.JWODE;

                                        SAPbouiCOM.Matrix matQt = (SAPbouiCOM.Matrix)oForm.Items.Item("23").Specific;
                                        int rowNum = 1;
                                        for (int i = 0; i < oHeader.lstChild.Count; i++)
                                        {
                                            ((SAPbouiCOM.EditText)matQt.Columns.Item("1").Cells.Item(rowNum).Specific).Value = oHeader.lstChild[i].ItemCode;
                                            ((SAPbouiCOM.EditText)matQt.Columns.Item("2").Cells.Item(rowNum).Specific).Value = oHeader.lstChild[i].ItemName;
                                            ((SAPbouiCOM.EditText)matQt.Columns.Item("10").Cells.Item(rowNum).Specific).Value = Convert.ToString(oHeader.lstChild[i].Quantity);
                                            ((SAPbouiCOM.EditText)matQt.Columns.Item("1470001039").Cells.Item(rowNum).Specific).Value = oHeader.lstChild[i].WhsFrm;
                                            ((SAPbouiCOM.EditText)matQt.Columns.Item("5").Cells.Item(rowNum).Specific).Value = oHeader.lstChild[i].whsTo;
                                             rowNum++;
                                        }
                                    } 
                                }
                            }
                            if (pVal.ItemUID == "tab2")
                            { 

                                /* 
                                int rowNum1 = 1;
                                SAPbouiCOM.Matrix matLDFG = (SAPbouiCOM.Matrix)oForm.Items.Item("matLDFG").Specific;
                                string qry1 = "select DocEntry, DocNum, DocDate from OIGN where U_JWODe = '" + oForm.Items.Item("tCode").Specific.Value + "'";
                                SAPbobsCOM.Recordset rec1 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                rec1.DoQuery(qry1);
                                if (rec1.RecordCount > 0)
                                {
                                    while (!rec1.EoF)
                                    {
                                        (matLDFG.Columns.Item("#").Cells.Item(rowNum1).Specific).Value = Convert.ToString(matLDFG.RowCount);
                                        ((SAPbouiCOM.EditText)matLDFG.Columns.Item("tTransType").Cells.Item(rowNum1).Specific).Value = "Goods Receipt";
                                        ((SAPbouiCOM.EditText)matLDFG.Columns.Item("tBaseDocNo").Cells.Item(rowNum1).Specific).Value = rec1.Fields.Item("DocNum").Value.ToString();
                                        ((SAPbouiCOM.EditText)matLDFG.Columns.Item("tBaseDE").Cells.Item(rowNum1).Specific).Value = rec1.Fields.Item("DocEntry").Value.ToString();
                                        DateTime tDocDate = Convert.ToDateTime(rec1.Fields.Item("DocDate").Value.ToString());
                                        ((SAPbouiCOM.EditText)matLDFG.Columns.Item("tDocDate").Cells.Item(rowNum1).Specific).Value = tDocDate.ToString("yyyyMMdd");
                                        rowNum1++;
                                        rec1.MoveNext();
                                    }
                                } 
                                int rowNum = 1;
                                SAPbouiCOM.Matrix matLDCO = (SAPbouiCOM.Matrix)oForm.Items.Item("matLDCO").Specific;
                                string getDocEntry = "select DocEntry, DocNum, DocDate from OWTR where U_JWODe = '" + oForm.Items.Item("tCode").Specific.Value + "'";
                                SAPbobsCOM.Recordset rec = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                rec.DoQuery(getDocEntry);
                                if (rec.RecordCount > 0)
                                {
                                    while (!rec.EoF)
                                    { 
                                        (matLDCO.Columns.Item("#").Cells.Item(rowNum).Specific).Value = Convert.ToString(matLDCO.RowCount);
                                        ((SAPbouiCOM.EditText)matLDCO.Columns.Item("tTransType").Cells.Item(rowNum).Specific).Value = "Inventory Transfer";
                                        ((SAPbouiCOM.EditText)matLDCO.Columns.Item("tBaseDocNo").Cells.Item(rowNum).Specific).Value = rec.Fields.Item("DocNum").Value.ToString();
                                        ((SAPbouiCOM.EditText)matLDCO.Columns.Item("tBaseDE").Cells.Item(rowNum).Specific).Value = rec.Fields.Item("DocEntry").Value.ToString();
                                        DateTime tDocDate = Convert.ToDateTime(rec.Fields.Item("DocDate").Value.ToString());
                                        ((SAPbouiCOM.EditText)matLDCO.Columns.Item("tDocDate").Cells.Item(rowNum).Specific).Value = tDocDate.ToString("yyyyMMdd");
                                        rowNum++;
                                        rec.MoveNext();
                                    }
                                } 
                                string qry2 = "select DocEntry, DocNum, DocDate from OIGE where U_JWODe = '" + oForm.Items.Item("tCode").Specific.Value + "'";
                                SAPbobsCOM.Recordset rec2 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                rec2.DoQuery(qry2);
                                if (rec2.RecordCount > 0)
                                {
                                    while (!rec2.EoF)
                                    {
                                        matLDCO.AddRow();
                                        (matLDCO.Columns.Item("#").Cells.Item(rowNum).Specific).Value = Convert.ToString(matLDCO.RowCount);
                                        ((SAPbouiCOM.EditText)matLDCO.Columns.Item("tTransType").Cells.Item(rowNum).Specific).Value = "Goods Issue";
                                        ((SAPbouiCOM.EditText)matLDCO.Columns.Item("tBaseDocNo").Cells.Item(rowNum).Specific).Value = rec2.Fields.Item("DocNum").Value.ToString();
                                        ((SAPbouiCOM.EditText)matLDCO.Columns.Item("tBaseDE").Cells.Item(rowNum).Specific).Value = rec2.Fields.Item("DocEntry").Value.ToString();
                                        DateTime tDocDate = Convert.ToDateTime(rec2.Fields.Item("DocDate").Value.ToString());
                                        ((SAPbouiCOM.EditText)matLDCO.Columns.Item("tDocDate").Cells.Item(rowNum).Specific).Value = tDocDate.ToString("yyyyMMdd");

                                        SAPbouiCOM.Column mCol = matLDCO.Columns.Item("tBaseDE"); 
                                        LinkedButton oLinkLns = ((SAPbouiCOM.LinkedButton)(mCol.ExtendedObject));
                                        oLinkLns.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_GoodsIssue; 

                                        //oLinkSOPO = (SAPbouiCOM.EditText)matLDCO.Columns.Item("tBaseDE").Cells.Item(rowNum).Specific;
                                        //oLinkSOPO.lin = SAPbouiCOM.BoLinkedObject.lf_GoodsIssue;

                                        rowNum++;
                                        rec2.MoveNext();
                                    }
                                } 
                                */
                            }

                            if (pVal.ItemUID == "lnknp")
                            {
                                string abc = oForm.Items.Item("tnopcode").Specific.Value;
                                objCU.FormLoadAndActivate("frmJWNPM", "mnsmJW003");
                                OutwardToJWPM InVar = new OutwardToJWPM();
                                InVar.NopCode = abc;
                                clsJWNPM oPort = new clsJWNPM(InVar);
                                //oForm.Close(); 
                            } 
                        }

                        break;
                    case BoEventTypes.et_MATRIX_LINK_PRESSED: 
                        oForm = SBOMain.SBO_Application.Forms.ActiveForm; 
                        SAPbouiCOM.Matrix matLDFG = (SAPbouiCOM.Matrix)oForm.Items.Item("matLDFG").Specific;

                        if (pVal.BeforeAction == true)
                        {
                            if (pVal.ItemUID == "matLDFG" && pVal.ColUID == "tBaseDE")
                            {
                                string de = Convert.ToString(((SAPbouiCOM.EditText)matLDFG.Columns.Item("tBaseDE").Cells.Item(pVal.Row).Specific).Value);
                                string doctype = ((SAPbouiCOM.EditText)matLDFG.Columns.Item("tTransType").Cells.Item(pVal.Row).Specific).Value;

                                if (doctype == "Inventory Transfer")
                                { 
                                    SBOMain.SBO_Application.Menus.Item("3080").Activate();
                                    OutwardToInvTrans InVar = new OutwardToInvTrans();
                                    InVar.DocEntry = de;
                                    clsInvTrans oPort = new clsInvTrans(InVar); 
                                }
                                else if (doctype == "Goods Issue")
                                {
                                    SBOMain.SBO_Application.Menus.Item("3079").Activate();
                                    OutwardToGI InVar1 = new OutwardToGI();
                                    InVar1.DocEntry = de;
                                    clsGoodsIssue oPort = new clsGoodsIssue(InVar1);
                                }
                                else if (doctype == "Goods Receipt")
                                {
                                    SBOMain.SBO_Application.Menus.Item("3078").Activate();
                                    OutwardToGR InVar = new OutwardToGR();
                                    InVar.DocEntry = de;
                                    clsGoodsReceipt oPort = new clsGoodsReceipt(InVar);
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
                                if (string.IsNullOrEmpty(oForm.Items.Item("tCardCode").Specific.value))
                                {
                                    BubbleEvent = false;
                                    SBOMain.SBO_Application.StatusBar.SetText("Please Add Card Code", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                    oForm.Items.Item("tCardCode").Click();
                                }
                                else if (string.IsNullOrEmpty(oForm.Items.Item("tCardName").Specific.value))
                                {
                                    BubbleEvent = false;
                                    SBOMain.SBO_Application.StatusBar.SetText("Please Add Card Name", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                    oForm.Items.Item("tCardName").Click();
                                }
                                else if (string.IsNullOrEmpty(oForm.Items.Item("tPoNum").Specific.value))
                                {
                                    BubbleEvent = false;
                                    SBOMain.SBO_Application.StatusBar.SetText("Please Add PO Number", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                    oForm.Items.Item("tPoNum").Click();
                                }
                                else if (string.IsNullOrEmpty(oForm.Items.Item("tPoDe").Specific.value))
                                {
                                    BubbleEvent = false;
                                    SBOMain.SBO_Application.StatusBar.SetText("PO DocEntry should not be blank", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                    oForm.Items.Item("tPoDe").Click();
                                }
                                else if (string.IsNullOrEmpty(oForm.Items.Item("tDocDate").Specific.value))
                                {
                                    BubbleEvent = false;
                                    SBOMain.SBO_Application.StatusBar.SetText("Please Add Document Date", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                    oForm.Items.Item("tDocDate").Click();
                                }
                                else if (string.IsNullOrEmpty(oForm.Items.Item("tDocDDt").Specific.value))
                                {
                                    BubbleEvent = false;
                                    SBOMain.SBO_Application.StatusBar.SetText("Please Add Delivery Date", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                    oForm.Items.Item("tDocDDt").Click();
                                }
                                else if (string.IsNullOrEmpty(oForm.Items.Item("tnopcode").Specific.value))
                                {
                                    BubbleEvent = false;
                                    SBOMain.SBO_Application.StatusBar.SetText("Please Select Nature of Processing", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                    oForm.Items.Item("tnopcode").Click();
                                }
                            }
                        }
                        if (pVal.BeforeAction == false)
                        {
                            if (pVal.ItemUID == "1" && (oForm.Mode == BoFormMode.fm_ADD_MODE || oForm.Mode == BoFormMode.fm_UPDATE_MODE))
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


            }
            finally
            {
                /*  if (oForm != null)
                oForm.Freeze(false);*/
            }
            return BubbleEvent;
        }

        private void CFLCondition(string CFLID, string ItemUID, string CardCode)
        {
            oCFL = oForm.ChooseFromLists.Item(CFLID);
            oConds = SBOMain.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions);

            if (CFLID == "CFL_OCRD")
            {
                oCond = oConds.Add();
                oCond.Alias = "validFor";
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCond.CondVal = "Y";

                oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

                oCond = oConds.Add();
                oCond.Alias = "CardType";
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCond.CondVal = "S";
                oCFL.SetConditions(oConds);
            }
            if (CFLID == "CFL_JOPM")
            {
                oCond = oConds.Add();
                oCond.Alias = "U_status";
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCond.CondVal = "1";
                oCFL.SetConditions(oConds);
            }

            if (CFLID == "CFL_OPOR")
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

                oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;
                oCond = oConds.Add();
                oCond.Alias = "DocStatus";
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCond.CondVal = "O";

                oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;
                oCond = oConds.Add();
                oCond.Alias = "DocDate";
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_GRATER_EQUAL;
                oCond.CondVal = "20220401";
                oCFL.SetConditions(oConds);

            }
            if (CFLID == "CFL_OITM")
            {
                oCond = oConds.Add();
                oCond.Alias = "validFor";
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCond.CondVal = "Y";

                oCFL.SetConditions(oConds);
            }
            if (CFLID == "CFL_OITM2")
            {
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
        public void Form_Load_Components(SAPbouiCOM.Form oForm, string mode)
        {
            //oForm.Freeze(true);
            oForm.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized;
            SAPbouiCOM.EditText oEdit;
            string Table = "@JOTR";
            DateTime now = DateTime.Now;
            if (mode != "OK")
            {
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                oEdit = oForm.Items.Item("tCode").Specific;
                objCU.GetNextDocNum(ref oEdit, ref Table);
                oForm.Items.Item("tDocNum").Specific.value = oForm.BusinessObject.GetNextSerialNumber("DocEntry", "JOTR");
                Events.Series.SeriesCombo("JOTR", "cSer");
                oForm.Items.Item("cSer").DisplayDesc = true;

                SAPbouiCOM.ComboBox cb = (SAPbouiCOM.ComboBox)oForm.Items.Item("tStatus").Specific;
                cb.ExpandType = BoExpandType.et_DescriptionOnly;
                cb.Select("O");

                SAPbouiCOM.ComboBox cb6 = (SAPbouiCOM.ComboBox)oForm.Items.Item("tBPLId").Specific;
                string getDocEntry = "SELECT BPLId, BPLName from OBPL  where BPLName != 'Main'";
                SAPbobsCOM.Recordset rec = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                rec.DoQuery(getDocEntry);
                if (rec.RecordCount > 0)
                {
                    if (cb6.ValidValues.Count == 0)
                    {
                        while (!rec.EoF)
                        {
                            cb6.ValidValues.Add(rec.Fields.Item("BPLId").Value.ToString(), rec.Fields.Item("BPLName").Value.ToString());
                            cb6.ExpandType = BoExpandType.et_DescriptionOnly;
                            rec.MoveNext();
                        }
                    }
                }
                cb6.Select("1");

                oForm.Items.Item("tab1").Visible = true;
                oForm.Items.Item("tab1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                oForm.PaneLevel = 1;
                oForm.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized;

                /***********/
                SAPbouiCOM.ButtonCombo cb2 = (SAPbouiCOM.ButtonCombo)oForm.Items.Item("cmbCPYT").Specific;
                cb2.ValidValues.Add("GI", "Goods Issue");
                cb2.ValidValues.Add("GR", "Goods Receipt");
                cb2.ValidValues.Add("SR", "Sub Contracting Return");
                cb2.ValidValues.Add("IT", "Inventory Transfer");
                cb2.ExpandType = BoExpandType.et_DescriptionOnly;
                /***********/

                oMatrix = oForm.Items.Item("matCMTR").Specific;
                AddMatrixRow(oMatrix, "tItemCode");
                 
            }
            

            if (mode == "OK")
            {
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                oEdit = oForm.Items.Item("tCode").Specific;
                objCU.GetNextDocNum(ref oEdit, ref Table);
                oForm.Items.Item("tDocNum").Specific.value = oForm.BusinessObject.GetNextSerialNumber("DocEntry", "JOTR");
                Events.Series.SeriesCombo("JOTR", "cSer");
                oForm.Items.Item("cSer").DisplayDesc = true;

                SAPbouiCOM.ComboBox cb6 = (SAPbouiCOM.ComboBox)oForm.Items.Item("tBPLId").Specific;
                cb6.Select("1");
            }



            //oForm.Freeze(false);
        }
        #region MatrixSetLine
        private void DeleteMatrixBlankRow(SAPbouiCOM.Matrix oMatrix, string colname)
        {
            try
            {
                for (int i = oMatrix.VisualRowCount; i >= 1; i--)
                {
                    {
                        if (string.IsNullOrEmpty(((SAPbouiCOM.EditText)oMatrix.Columns.Item(colname).Cells.Item(i).Specific).Value))
                            oMatrix.DeleteRow(i);
                    }
                }
            }
            catch
            {
            }
        }
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
        private void ArrengeMatrixLineNum(SAPbouiCOM.Matrix matrix)
        {
            for (int i = 1; i <= matrix.RowCount; i++)
            {
                matrix.Columns.Item("#").Cells.Item(i).Specific.value = i;
            }
        }
        #endregion
    }
    public class HeaderJWO
    {
        public string PoNum { get; set; }
        public string SoNum { get; set; }

        public string BPCode { get; set; }
        public string BPName { get; set; }
        public string PoDe { get; set; }
        public string SoDe { get; set; }

        public string NumAtCard { get; set; }
        public string DocDate { get; set; }
        public string DocDDt { get; set; }
        public string BPLId { get; set; }
        public string Whs { get; set; }
        public string JWODE { get; set; }
        public string JWIDE { get; set; }

        public List<ChildJWO> lstChild { get; set; }
    }
    public class ChildJWO
    {
        public string ItemCode { get; set; }
        public string ItemName { get; set; }
        public decimal Quantity { get; set; }
        public decimal UnitPrice { get; set; }
        public string WhsFrm { get; set; }
        public string whsTo { get; set; } 
        public decimal BalQty { get; set; }
        public decimal IsdQty { get; set; } 

    }
}
