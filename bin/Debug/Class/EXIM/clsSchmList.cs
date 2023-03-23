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
    class clsSchmList
    {
        #region VariableDeclaration

        private SAPbouiCOM.Form oForm;

        #endregion VariableDeclaration

        private void ArrengeMatrixLineNum(SAPbouiCOM.Matrix matrix)
        {
            for (int i = 1; i <= matrix.RowCount; i++)
            {
                matrix.Columns.Item("#").Cells.Item(i).Specific.value = i;
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
                    case BoEventTypes.et_FORM_LOAD:
                        // oForm = SBOMain.SBO_Application.Forms.ActiveForm;
                        if (pVal.BeforeAction == true)
                        {

                        }
                        if (pVal.BeforeAction == false)
                        {

                        }
                        break;
                    case BoEventTypes.et_CLICK:
                        oForm = SBOMain.SBO_Application.Forms.ActiveForm;
                        if (pVal.BeforeAction == true)
                        {

                        }
                        if (pVal.BeforeAction == false)
                        {
                            if (pVal.ItemUID == "btnCHPL")
                            {
                                SAPbouiCOM.Grid objGrid = oForm.Items.Item("gridSchm").Specific;

                                string ExpCode = string.Empty;
                                for (int i = 0; i < objGrid.Rows.SelectedRows.Count; i++)
                                {
                                    int index = objGrid.Rows.SelectedRows.Item(i, SAPbouiCOM.BoOrderType.ot_RowOrder);
                                    ExpCode = objGrid.DataTable.GetValue(0, index).ToString();
                                } 

                                oForm.Close();
                                oForm = SBOMain.SBO_Application.Forms.ActiveForm;
                                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item(Program.ETRFLLN_Data.ETMat).Specific;
                                ((SAPbouiCOM.EditText)oMatrix.Columns.Item(Program.ETRFLLN_Data.ETMatCol).Cells.Item(Program.ETRFLLN_Data.ETMatRow).Specific).Value = ExpCode;

                                oForm.Freeze(true);
                                string getDocEntry = "SELECT T0.U_schLEDE , T1.U_rmqty, T1.U_rmLC, T1.U_rmFC FROM dbo.[@EXSM] as T0 ";
                                getDocEntry = getDocEntry + " LEFT JOIN dbo.[@XSM1] AS T1 ON T0.Code = T1.Code ";
                                getDocEntry = getDocEntry + " Where T0.U_schno = '"+ ExpCode + "' AND T1.U_itemcode = '"+ Program.ETRFLLN_Data.ETMatItemcode + "' ";
                                SAPbobsCOM.Recordset rec = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                rec.DoQuery(getDocEntry);
                                if (rec.RecordCount > 0)
                                {
                                    DateTime lcsd = Convert.ToDateTime(rec.Fields.Item("U_schLEDE").Value);
                                    string abc = lcsd.ToString("yyyyMMdd");

                                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("ex4lv").Cells.Item(Program.ETRFLLN_Data.ETMatRow).Specific).Value = abc;
                                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("ex4lrqty").Cells.Item(Program.ETRFLLN_Data.ETMatRow).Specific).Value = Convert.ToString(rec.Fields.Item("U_rmqty").Value);
                                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("ex4lrafc").Cells.Item(Program.ETRFLLN_Data.ETMatRow).Specific).Value = Convert.ToString(rec.Fields.Item("U_rmFC").Value);  
                                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("ex4lraklc").Cells.Item(Program.ETRFLLN_Data.ETMatRow).Specific).Value = Convert.ToString(rec.Fields.Item("U_rmLC").Value);

                                    double  qty = Convert.ToDouble((oMatrix.Columns.Item("ex4lfqty").Cells.Item(Program.ETRFLLN_Data.ETMatRow).Specific).Value);
                                     
                                    SAPbouiCOM.Matrix matRLF2 = (SAPbouiCOM.Matrix)oForm.Items.Item("matRLF2").Specific;
                                    int j =1;
                                    string Qry2 = " Select T1.U_iritemcd,T1.U_iritemnm, T1.U_irQtyPer, T1.U_irExQtP FROM dbo.[@EXSM] AS T0 LEFT JOIN dbo.[@XSM3] AS ";
                                    Qry2 = Qry2 + " T1 ON T0.Code = T1.Code Where T1.U_iritemcd = '" + Program.ETRFLLN_Data.ETMatItemcode + "' AND T0.U_schno = '" + ExpCode + "'";
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
                                }
                                
                                oForm.Freeze(false);
                            }

                            if (pVal.ItemUID == "btnCHS")
                            {
                                try
                                {
                                    SAPbouiCOM.Grid objGrid = oForm.Items.Item("grid").Specific;
                                    oForm.Close();
                                }
                                catch (Exception ex)
                                {

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

            }
            return BubbleEvent;
        }
    }
}
