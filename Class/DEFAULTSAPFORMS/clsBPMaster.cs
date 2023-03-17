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
using CoreSuteConnect.Class.EXIM;
using CoreSuteConnect.Class.JOBWORK;
using CoreSuteConnect.Events;
using System.Data;
using System.Security.Cryptography;
using System.Drawing.Drawing2D;


namespace CoreSuteConnect.Class.DEFAULTSAPFORMS
{
    class clsBPMaster
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
                { // KEY Down event for multiplication
                    case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
                        //SAPbouiCOM.Form oForms = SBOMain.SBO_Application.Forms.ActiveForm;
                        if (pVal.BeforeAction == true)
                        {
                        }
                        if (pVal.BeforeAction == false)
                        {
                            /*SAPbouiCOM.IChooseFromListEvent oCFLEvento = null;
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
                                    if (pVal.ItemUID == "U_DefWhs")
                                    {
                                        string DefWhs = oForm.Items.Item("U_DefWhs").Specific.value;

                                        if (!string.IsNullOrEmpty(DefWhs))
                                        {
                                            //SAPbouiCOM.Form oUDFForm = SBOMain.SBO_Application.Forms.Item(oForm.UDFFormUID);

                                            string getDefWhs = "select * from OITM where U_DefWhs ='" + DefWhs + "'";
                                            SAPbobsCOM.Recordset rec1 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                            rec1.DoQuery(getDefWhs);
                                            if (rec1.RecordCount > 0)
                                            {
                                                BubbleEvent = false;
                                                SBOMain.SBO_Application.StatusBar.SetText("Jobwork Warehouse already assigned to other Business Partner.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                                oForm.Items.Item("tCardCode").Click();
                                            }
                                        }
                                        oForm.Items.Item("U_DefWhs").Specific.value = oDataTable.GetValue("U_nopcode", 0).ToString();
                                    } 
                                }
                                catch (Exception ex)
                                {

                                }
                            }*/
                        }
                        break;
                    case BoEventTypes.et_ITEM_PRESSED:
                        oForm = SBOMain.SBO_Application.Forms.ActiveForm;
                        if (pVal.BeforeAction == true)
                        { 
                            if (pVal.ItemUID == "1" && (oForm.Mode == BoFormMode.fm_ADD_MODE || oForm.Mode == BoFormMode.fm_UPDATE_MODE))
                            {
                                /*SAPbouiCOM.Form oUDFForm = SBOMain.SBO_Application.Forms.Item(oForm.UDFFormUID);
                                string DefWhs = oUDFForm.Items.Item("U_DefWhs").Specific.value;

                                if (!string.IsNullOrEmpty(DefWhs))
                                {
                                    string getDefWhs = "select * from OITM where U_DefWhs ='" + DefWhs + "'";
                                    SAPbobsCOM.Recordset rec1 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                    rec1.DoQuery(getDefWhs);
                                    if (rec1.RecordCount > 0)
                                    {
                                        BubbleEvent = false;
                                        SBOMain.SBO_Application.StatusBar.SetText("Jobwork Warehouse already assigned to other Business Partner.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                        oForm.Items.Item("tCardCode").Click();
                                    }
                                }*/

                            }
                        }
                        if (pVal.BeforeAction == false)
                        {
                            
                        }
                        break;


                }
            }
            catch (Exception ex)
            {

            }

            return BubbleEvent;
        }
    }
}
