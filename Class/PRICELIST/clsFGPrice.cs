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
    class clsFGPrice
    {
        #region VariableDeclaration
        
        public static SAPbouiCOM.Application SBO_Application;

        private SAPbouiCOM.Form oForm;
        private SAPbouiCOM.Matrix oMatrix;
        public string cFormID = string.Empty;

        private SAPbouiCOM.ChooseFromList oCFL = null;
        private SAPbouiCOM.Conditions oConds = null;
        private SAPbouiCOM.Condition oCond = null;

        HeaderFG oHeaderFG = new HeaderFG();

        #endregion VariableDeclaration

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

                SAPbouiCOM.Matrix matFGItem = (SAPbouiCOM.Matrix)oForm.Items.Item("matFGItem").Specific;

                if (pVal.BeforeAction == true)
                {
                    if (oForm.Mode == BoFormMode.fm_OK_MODE)
                    {
                        string title = oForm.Items.Item("title").Specific.value;
                        oHeaderFG.title = title;
                         
                        if (matFGItem.RowCount > 0)
                        {
                            List<ChildFG> lstChildFG = new List<ChildFG>();
                            for (int i = 1; i <= matFGItem.RowCount; i++)
                            {
                                ChildFG ochild = new ChildFG();
                                ochild.ItemCode = ((SAPbouiCOM.EditText)matFGItem.Columns.Item("itemcode").Cells.Item(i).Specific).Value;
                                ochild.ItemName = ((SAPbouiCOM.EditText)matFGItem.Columns.Item("itemname").Cells.Item(i).Specific).Value;
                                ochild.Price = Convert.ToDouble(((SAPbouiCOM.EditText)matFGItem.Columns.Item("price").Cells.Item(i).Specific).Value);
                                 lstChildFG.Add(ochild);
                            }
                            oHeaderFG.lstChildFG = lstChildFG; 
                        }  
                    }
                }
                if (pVal.BeforeAction == false)
                {   
                    if ((oForm.Mode == BoFormMode.fm_ADD_MODE || Type == "ADDNEWFORM") && Type != "DEL_ROW")
                    {
                        Form_Load_Components(oForm);
                    }

                    if (Type == "DEL_ROW")
                    {
                        ArrengeMatrixLineNum(matFGItem);
                    }
                    else if (Type == "ADD_ROW")
                    {
                        if (SBOMain.RightClickItemID == "matFGItem")
                        {
                            matFGItem.AddRow(1, SBOMain.RightClickLineNum);
                            matFGItem.ClearRowData(SBOMain.RightClickLineNum + 1);
                            ArrengeMatrixLineNum(matFGItem);
                        }
                    } 
                    else if (Type == "Duplicate")
                    {
                        if (oHeaderFG != null)
                        { 
                            oForm.Items.Item("title").Specific.value = oHeaderFG.title; 
                            for (int i = 1; i <= oHeaderFG.lstChildFG.Count; i++)
                            {
                                ((SAPbouiCOM.EditText)matFGItem.Columns.Item("#").Cells.Item(i).Specific).Value = (i).ToString();
                                ((SAPbouiCOM.EditText)matFGItem.Columns.Item("itemcode").Cells.Item(i).Specific).Value = oHeaderFG.lstChildFG[i].ItemCode;
                                ((SAPbouiCOM.EditText)matFGItem.Columns.Item("itemcode").Cells.Item(i).Specific).Value = oHeaderFG.lstChildFG[i].ItemCode;
                                ((SAPbouiCOM.EditText)matFGItem.Columns.Item("itemname").Cells.Item(i).Specific).Value = oHeaderFG.lstChildFG[i].ItemName;
                                ((SAPbouiCOM.EditText)matFGItem.Columns.Item("price").Cells.Item(i).Specific).Value = oHeaderFG.lstChildFG[i].Price.ToString();
                                matFGItem.AddRow();
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
                { // KEY Down event for multiplication  
                    case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
                        if (pVal.BeforeAction == true)
                        {

                        }
                        if (pVal.BeforeAction == false)
                        {
                            if (pVal.CharPressed == 9)
                            {
                                if (pVal.ItemUID == "matFGItem" && pVal.ColUID == "itemcode")
                                {
                                    SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("matFGItem").Specific;
                                    AddMatrixRow(oMatrix, "itemcode");
                                }
                            }
                            if (pVal.ItemUID == "matFGItem" && pVal.ColUID == "price")
                            {
                                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("matFGItem").Specific;
                                AddMatrixRow(oMatrix, "itemcode");
                            }
                        }
                        break;
                    case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
                        if (pVal.BeforeAction == true)
                        {
                            if (pVal.ItemUID == "itemcode")
                            {
                                CFLCondition("CFL_OITM"); 
                            }
                            if (pVal.ItemUID == "matFGItem")
                            { 
                                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("matFGItem").Specific;
                                AddMatrixRow(oMatrix, "itemcode");
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
                                SAPbouiCOM.Matrix matrix = oForm.Items.Item("matFGItem").Specific;
                                if (pVal.ItemUID== "matFGItem" && pVal.ColUID == "itemcode")
                                {
                                    try
                                    {
                                        ((SAPbouiCOM.EditText)matrix.Columns.Item("itemname").Cells.Item(pVal.Row).Specific).Value = oDataTable.GetValue("ItemName", 0).ToString();
                                        ((SAPbouiCOM.EditText)matrix.Columns.Item("itemcode").Cells.Item(pVal.Row).Specific).Value = oDataTable.GetValue("ItemCode", 0).ToString();
                                    }
                                    catch
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
                            
                        }
                        if (pVal.BeforeAction == false)
                        {


                        } 
                        break;
                    case BoEventTypes.et_ITEM_PRESSED:
                        oForm = SBOMain.SBO_Application.Forms.ActiveForm;
                        if (pVal.BeforeAction == true)
                        {
                            if (pVal.ItemUID == "1" && oForm.Mode == BoFormMode.fm_UPDATE_MODE)
                            {
                              string statusval =  oForm.Items.Item("cmbStatus").Specific.value.ToString();
                              if(statusval == "C")
                                {
                                    SBOMain.SBO_Application.StatusBar.SetText("We Can not udpate document because it is Closed", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                     BubbleEvent = false; 
                                }
                            } 
                            else  if (pVal.ItemUID == "1" && oForm.Mode == BoFormMode.fm_ADD_MODE)
                            {
                                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("matFGItem").Specific;
                                DeleteMatrixBlankRow(oMatrix);

                                string getQuery = @"UPDATE dbo.[@FGPL] SET U_status = 'C'"; 
                                SAPbobsCOM.Recordset rec;
                                rec = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                rec.DoQuery(getQuery);

                            }
                        }
                        if (pVal.BeforeAction == false)
                        {
                            if (pVal.ItemUID == "1" && oForm.Mode == BoFormMode.fm_ADD_MODE)
                            { 
                                Form_Load_Components(oForm);
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
                
            }


            return BubbleEvent;
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
        private void DeleteMatrixBlankRow(SAPbouiCOM.Matrix oMatrix)
        {
            try
            {
                for (int i = oMatrix.VisualRowCount; i >= 1; i--)
                {
                    {
                        if (string.IsNullOrEmpty(((SAPbouiCOM.EditText)oMatrix.Columns.Item("itemcode").Cells.Item(i).Specific).Value))
                            oMatrix.DeleteRow(i);
                    }
                }
            }
            catch
            {
            }
        }
        private void CFLCondition(string CFLID)
        {
            oCFL = oForm.ChooseFromLists.Item(CFLID);
            oConds = SBOMain.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions);
            if (CFLID == "CFL_OITM")
            {
                oCond = oConds.Add();
                oCond.Alias = "Series";
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCond.CondVal = "101";

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

        public void Form_Load_Components(SAPbouiCOM.Form oForm)
        {
            SetCode();
            oMatrix = oForm.Items.Item("matFGItem").Specific;
            AddMatrixRow(oMatrix, "itemcode");

            SAPbouiCOM.ComboBox cb = (SAPbouiCOM.ComboBox)oForm.Items.Item("cmbStatus").Specific;
            cb.ExpandType = BoExpandType.et_DescriptionOnly;
            cb.Select("O");
            oForm.Items.Item("docDate").Specific.Value = DateTime.Today.ToString("yyyyMMdd");

        }
        private void SetCode()
        {
            oForm.Freeze(true);
            oForm.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized;
            string TableName = "FGPL";
            SBOMain.SetCode(oForm.UniqueID, TableName);
            oForm.Freeze(false);
        }
    }

    public class HeaderFG
    {
        public string title { get; set; }
      
        public List<ChildFG> lstChildFG { get; set; }
    }
    public class ChildFG
    {
        public string ItemCode { get; set; }
        public string ItemName { get; set; }
        public double Price { get; set; }
         
    }

}
