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
    class clsPortList
    {
        #region VariableDeclaration

        private SAPbouiCOM.Form oForm;

        #endregion VariableDeclaration

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
                                SAPbouiCOM.Grid objGrid = oForm.Items.Item("gridPORT").Specific;

                                string PorCode = string.Empty;
                                for (int i = 0; i < objGrid.Rows.SelectedRows.Count; i++)
                                {
                                    int index = objGrid.Rows.SelectedRows.Item(i, SAPbouiCOM.BoOrderType.ot_RowOrder);

                                    PorCode = objGrid.DataTable.GetValue(0, index).ToString();
                                }

                                oForm.Close();
                                oForm = SBOMain.SBO_Application.Forms.ActiveForm;

                                if (Program.lstStyleCode.PortSelect == "ex1pol")
                                    ((SAPbouiCOM.EditText)oForm.Items.Item("ex1pol").Specific).Value = PorCode;
                                else if (Program.lstStyleCode.PortSelect == "ex1pod")
                                    ((SAPbouiCOM.EditText)oForm.Items.Item("ex1pod").Specific).Value = PorCode;
                                else if (Program.lstStyleCode.PortSelect == "ex1por")
                                    ((SAPbouiCOM.EditText)oForm.Items.Item("ex1por").Specific).Value = PorCode;
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
