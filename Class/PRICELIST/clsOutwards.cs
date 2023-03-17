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
    class clsOutwards
    {
        #region VariableDeclaration
        private SAPbouiCOM.Form oForm;
        public string cFormID = string.Empty;
        public static string _CardCode = string.Empty;

        private SAPbouiCOM.ChooseFromList oCFL = null;
        private SAPbouiCOM.Conditions oConds = null;
        private SAPbouiCOM.Condition oCond = null;

        #endregion VariableDeclaration

        public clsOutwards(PriceListToOutWord outClass)
        {
            if (outClass != null)
            {
                oForm = SBOMain.SBO_Application.Forms.ActiveForm;
                SAPbouiCOM.Button oBtnn = oForm.Items.Item("btGetData").Specific; 
                
                _CardCode = outClass.BPCode;
                oForm.Items.Item("cardcode").Specific.value = outClass.BPCode;
                oForm.Items.Item("cardname").Specific.value = outClass.BPName;
               // SBOMain.SBO_Application.SendKeys("{TAB}"); 
            }
        }

        public bool MenuEvent(ref SAPbouiCOM.MenuEvent pVal, string FormId)
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
                    var fromdate = DateTime.Today.AddYears(-1).ToString("yyyyMMdd"); 
                    oForm.Items.Item("tFrom").Specific.Value = fromdate;
                    oForm.Items.Item("tTo").Specific.Value = DateTime.Today.ToString("yyyyMMdd");
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
                    case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
                        if (pVal.BeforeAction == true)
                        {
                            if (pVal.ItemUID == "cardcode")
                            {
                                CFLCondition("CFL_OCRD");
                            }
                        }
                        if (pVal.BeforeAction == false && pVal.ItemUID == "cardcode")
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
                                if (pVal.ItemUID == "cardcode")
                                {
                                    try
                                    {
                                        oForm.Items.Item("cardcode").Specific.value = oDataTable.GetValue("CardCode", 0).ToString();
                                        oForm.Items.Item("cardname").Specific.value = oDataTable.GetValue("CardName", 0).ToString();
                                    }
                                    catch (Exception ex)
                                    {

                                    }

                                }
                            }
                        }
                        if (pVal.BeforeAction == false && pVal.ItemUID == "ItemName")
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
                                if (pVal.ItemUID == "ItemName")
                                {
                                    try
                                    {
                                        oForm.Items.Item("ItemName").Specific.value = oDataTable.GetValue("ItemName", 0).ToString();
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
                            if(pVal.ItemUID == "btGetData")
                            {

                            }
                            if(pVal.ItemUID == "btCopyTo")
                            {

                            }
                        }
                        if(pVal.BeforeAction == false)
                        {
                            if (pVal.ItemUID == "btGetData")
                            {
                                loadMatrixData();
                            }

                            if (pVal.ItemUID == "btCopyTo")
                            {
                                string CardName = oForm.Items.Item("cardname").Specific.value;
                                string CardCode = oForm.Items.Item("cardcode").Specific.value;
                                SAPbouiCOM.Matrix matrix = oForm.Items.Item("mtData").Specific;

                                if (matrix.RowCount > 0)
                                {
                                    OutwardToPriceList outt = new OutwardToPriceList();
                                    outt.BPName = CardName;
                                    outt.BPCode = CardCode;
                                    //outt.BPCode = _CardCode;

                                    List<OutwardToPriceList_Child> listchild = new List<OutwardToPriceList_Child>();

                                    for (int i = 1; i <= matrix.RowCount; i++)
                                    {
                                        bool isChecked = ((SAPbouiCOM.CheckBox)matrix.Columns.Item("select").Cells.Item(i).Specific).Checked;
                                        if (isChecked)
                                        {
                                            OutwardToPriceList_Child outt_child = new OutwardToPriceList_Child();

                                           // outt_child.numerator = ((SAPbouiCOM.EditText)matrix.Columns.Item("numerator").Cells.Item(i).Specific).Value;
                                            outt_child.outwrdno = ((SAPbouiCOM.EditText)matrix.Columns.Item("outwrdno").Cells.Item(i).Specific).Value;
                                            outt_child.oDate = ((SAPbouiCOM.EditText)matrix.Columns.Item("oDate").Cells.Item(i).Specific).Value;
                                            outt_child.Itemno = ((SAPbouiCOM.EditText)matrix.Columns.Item("Itemno").Cells.Item(i).Specific).Value;
                                            outt_child.desc = ((SAPbouiCOM.EditText)matrix.Columns.Item("desc").Cells.Item(i).Specific).Value;
                                            outt_child.refno = ((SAPbouiCOM.EditText)matrix.Columns.Item("refno").Cells.Item(i).Specific).Value;

                                            listchild.Add(outt_child);
                                        }
                                    }
                                    outt.lstOut = listchild;

                                    bool plFormOpen = false;
                                    for (int i = 1; i < SBOMain.SBO_Application.Forms.Count; i++)
                                    {
                                        if (SBOMain.SBO_Application.Forms.Item(i).UniqueID == "frmPriceList")
                                        {
                                            SBOMain.SBO_Application.Forms.Item(i).Select();
                                            plFormOpen = true;
                                        }
                                    }

                                    if (!plFormOpen)
                                    {
                                        SBOMain.SBO_Application.Menus.Item("mnsmPL002").Activate();
                                    }
                                    clsPriceList oPrice = new clsPriceList(outt);
                                    oForm.Close();
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
                                if (string.IsNullOrEmpty(oForm.Items.Item("tdoccode").Specific.value))
                                {
                                    BubbleEvent = false;
                                    SBOMain.SBO_Application.StatusBar.SetText("Please Add Document Code", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                    oForm.Items.Item("tdoccode").Click();
                                }
                            }
                        }
                        if (pVal.BeforeAction == false)
                        {

                        } 
                        break;

                    case BoEventTypes.et_KEY_DOWN:
                        if (pVal.BeforeAction == true)
                        {
                            
                        }
                        if (pVal.BeforeAction == false && pVal.CharPressed == 9 && pVal.ItemUID == "cardname")
                        { 
                            loadMatrixData(); 
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
                /*if (oForm != null)
                    oForm.Freeze(false);*/
            }
            return BubbleEvent;
        }

        public void loadMatrixData()
        {

            try
            {

            
                oForm = SBOMain.SBO_Application.Forms.ActiveForm;

                string CardName = oForm.Items.Item("cardname").Specific.value;
                string FromDate = oForm.Items.Item("tFrom").Specific.value;
                string ToDate = oForm.Items.Item("tTo").Specific.value;
                string FromDateConvert = FromDate.Substring(0, 4) + "-" + FromDate.Substring(4, 2) + "-" + FromDate.Substring(6, 2);
                string ToDateConvert = ToDate.Substring(0, 4) + "-" + ToDate.Substring(4, 2) + "-" + ToDate.Substring(6, 2);


                SAPbouiCOM.Matrix matrix = oForm.Items.Item("mtData").Specific;

                string getQuery = @"select T0.DocEntry,  T0.DocNum as 'Inward/Outward No.',
                                                 Format(t0.DocDate, 'dd/MM/yyyy ') as 'Date', T1.ItemCode, t1.Dscription, 
                                                 (Select case when T1.itemcode like 'LM%' then ISNULL(T0.U_InOutRef, IBT1.Batchnum) else 
                                                ISNULL(T0.U_InOutRef, T1.U_OutRef) end) as 'Inward/Outward Ref no',                                                   
                                                ISNULL(OBTN.U_Ven_Ref, OBTN.U_Cus_Ref) as 'Party Ref No.',
                                                ISNULL(OBTN.U_Ven_RefDt, OBTN.U_Cus_RefDt) as 'Party Ref Date' 
                                                FROM OIGE t0
                                                Left Join IGE1 T1 on T0.DocEntry = T1.DocEntry
                                                LEFT Outer JOIN IBT1 ON IBT1.BaseEntry = t1.DocEntry AND IBT1.ItemCode = t1.ItemCode
                                                AND IBT1.WhsCode = t1.WhsCode AND IBT1.BaseLinNum = t1.LineNum AND IBT1.BaseType = 60
                                                LEFT Outer JOIN OBTN ON IBT1.BatchNum = OBTN.DistNumber AND IBT1.ItemCode = OBTN.ItemCode
                                                Left Join OCRD T3 on t0.U_BP_Name = T3.CardCode
                                                where
                                                t3.CardName = '" + CardName + @"' 
											    AND (CAST(T0.DocDate AS DATE) >= '" + FromDateConvert + @"') 
											    AND (CAST(T0.DocDate AS DATE) <= '" + ToDateConvert + @"')
                                                AND t0.U_Cate = 'OUT'";

                SAPbobsCOM.Recordset rec;
                rec = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                rec.DoQuery(getQuery);

                matrix.Clear();
                matrix.FlushToDataSource();

               int Progress = 0;
               SAPbouiCOM.ProgressBar oProgressBar;
               oProgressBar = SBOMain.SBO_Application.StatusBar.CreateProgressBar("Please Wait", rec.RecordCount, true);
               oProgressBar.Text = "Please Wait";

                int i =1;
                oForm.Freeze(true);
                if (rec.RecordCount > 0)
                {
                    while (!rec.EoF)
                    {
                        Progress += 1;
                        oProgressBar.Value = Progress;
                        matrix.AddRow();
                        (matrix.Columns.Item("#").Cells.Item(i).Specific).Value = Convert.ToString(matrix.RowCount);
                        (matrix.Columns.Item("oDate").Cells.Item(i).Specific).Value = rec.Fields.Item("Date").Value;
                        (matrix.Columns.Item("outwrdno").Cells.Item(i).Specific).Value = rec.Fields.Item("Inward/Outward No.").Value;
                        (matrix.Columns.Item("Itemno").Cells.Item(i).Specific).Value = rec.Fields.Item("ItemCode").Value;
                        (matrix.Columns.Item("desc").Cells.Item(i).Specific).Value = rec.Fields.Item("Dscription").Value;
                        (matrix.Columns.Item("refno").Cells.Item(i).Specific).Value = rec.Fields.Item("Inward/Outward Ref no").Value;
                        i++;
                        rec.MoveNext();
                    }
                }
                matrix.FlushToDataSource();
                oForm.Freeze(false);
                oProgressBar.Stop();
            }
            catch(Exception ex)
            {
                SBOMain.SBO_Application.StatusBar.SetText("Error" + ex.Message.ToString(), BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                oForm.Freeze(false);
            }
        }
        private void CFLCondition(string CFLID)
        {
            oCFL = oForm.ChooseFromLists.Item(CFLID);
            oConds = SBOMain.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions);
            if (CFLID == "CFL_OCRD")
            {
      
                oCond = oConds.Add();
                oCond.Alias = "CardType";
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_EQUAL;
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
    }

    public class OutwardToPriceList
    {
        public string BPCode { get; set; }
        public string BPName { get; set; }

        public List<OutwardToPriceList_Child> lstOut { get; set; }

    }

    public class OutwardToPriceList_Child
    {
        public string numerator { get; set; }
        public string outwrdno { get; set; }
        public string oDate { get; set; }
        public string Itemno { get; set; }
        public string desc { get; set; }
        public string refno { get; set; }
    }
}
