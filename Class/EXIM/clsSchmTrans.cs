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
     class clsSchmTrans
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
        string BrowseFilePath = string.Empty;

        SAPbouiCOM.ComboBox cb1;
        CommonUtility objCU = new CommonUtility();

        #endregion VariableDeclaration


        public bool MenuEvent(ref SAPbouiCOM.MenuEvent pVal, string FormId)
        {
            bool BubbleEvent = true;
            try
            {
                cFormID = FormId;
                oForm = SBOMain.SBO_Application.Forms.Item(FormId);
                if (oForm.Mode == BoFormMode.fm_ADD_MODE)
                {
                    Form_Load_Components(oForm);
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
                                oForm.Items.Item("tDocNum").Specific.Value = oForm.BusinessObject.GetNextSerialNumber(oForm.Items.Item("cSer").Specific.Value, "EXRU");
                            }
                        }
                        break;
                    case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:

                        if (pVal.BeforeAction == true)
                        {
                            
                        }
                        if (pVal.BeforeAction == false)
                        {
                           
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
                        if (pVal.ItemUID == "1" && pVal.BeforeAction == true)
                        {
                             
                        }
                        if (pVal.ItemUID == "1" && pVal.BeforeAction == false)
                        {
                           
                        }
                        break;
                    case BoEventTypes.et_ITEM_PRESSED:
                        oForm = SBOMain.SBO_Application.Forms.ActiveForm;

                        if (pVal.BeforeAction == false && pVal.ItemUID == "1")
                        {
                            Form_Load_Components(oForm);
                        }
                        else
                        {
                          /*  if (pVal.ItemUID == "btnATC" && pVal.BeforeAction == false)
                            {
                                OpenFile();
                            }*/
                        }
                        break;

                    case BoEventTypes.et_KEY_DOWN:

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
        private void SetCode()
        {
            oForm.Freeze(true);
            oForm.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized;
            string TableName = "EXRU";
            SBOMain.SetCode(oForm.UniqueID, TableName);
            oForm.Freeze(false);
            //throw new NotImplementedException();
        }


        public void Form_Load_Components(SAPbouiCOM.Form oForm)
        {

            //oForm.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized;
           SetCode();
           // oForm.Freeze(true);

           
            /*SAPbouiCOM.EditText oEdit;
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

            SAPbouiCOM.ComboBox cb = (SAPbouiCOM.ComboBox)oForm.Items.Item("schst").Specific;
            cb.ExpandType = BoExpandType.et_DescriptionOnly;
            cb.Select("O");


            cb1 = (SAPbouiCOM.ComboBox)oForm.Items.Item("clctype").Specific;
            cb1.ExpandType = BoExpandType.et_DescriptionOnly;
            cb1.Select("E");

            oMatrix = oForm.Items.Item("matSTDET").Specific;
            AddMatrixRow(oMatrix, "schpcpn");
            */
          //  oForm.Freeze(false);
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
        #endregion
        private void DeleteMatrixBlankRow(SAPbouiCOM.Matrix oMatrix)
        {
            try
            {
                for (int i = oMatrix.VisualRowCount; i >= 1; i--)
                {
                    {
                        if (string.IsNullOrEmpty(((SAPbouiCOM.EditText)oMatrix.Columns.Item("clVenCo").Cells.Item(i).Specific).Value))
                            oMatrix.DeleteRow(i);
                    }
                }

            }
            catch (Exception ex)
            {
            }
        }
    }
}
