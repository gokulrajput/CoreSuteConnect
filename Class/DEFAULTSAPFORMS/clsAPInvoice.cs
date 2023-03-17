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


namespace CoreSuteConnect.Class.DEFAULTSAPFORMS
{
    class clsAPInvoice
    {
        #region VariableDeclaration
        private SAPbouiCOM.Form oForm;
        public string cFormID = string.Empty;

        public static string getSalesForm = string.Empty;

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

                    case SAPbouiCOM.BoEventTypes.et_FORM_LOAD:
                        // oForm = SBOMain.SBO_Application.Forms.ActiveForm;
                        oForm = SBOMain.SBO_Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);

                        if (pVal.BeforeAction == true)
                        {

                        }
                        if (pVal.BeforeAction == false)
                        {
                            SAPbouiCOM.Item oItem = (SAPbouiCOM.Item)oForm.Items.Item("2");  /// Existing Item on the form of Cancel Button

                            #region EXIM TRACKING Button
                            SAPbouiCOM.Button btn = (SAPbouiCOM.Button)oForm.Items.Add("btnExTrk", SAPbouiCOM.BoFormItemTypes.it_BUTTON).Specific;
                            btn.Item.Top = oItem.Top;
                            btn.Item.Left = oItem.Left + oItem.Width + 7;
                            btn.Item.Width = oItem.Width + 10;
                            btn.Item.Height = oItem.Height;
                            btn.Item.Enabled = true;
                            btn.Caption = "EXIM Tracking";
                            #endregion EXIM TRACKING Button
                        }
                        break;

                    case BoEventTypes.et_CLICK:

                        oForm = SBOMain.SBO_Application.Forms.ActiveForm;
                        if (pVal.BeforeAction == true)
                        {

                        }
                        if (pVal.BeforeAction == false)
                        {
                            if (pVal.ItemUID == "btnExTrk" && (oForm.Mode == BoFormMode.fm_OK_MODE))
                            { 
                                /* OutwardToEximTracking outEximTracking = new OutwardToEximTracking();
                                 outt.BPName = CardName;
                                 outt.BPCode = CardCode;*/ 
                                 objCU.FormLoadAndActivate("frmETTrans", "mnsmEXIM013");  
                                //clsPriceList oPrice = new clsPriceList(outt);
                                //oForm.Close();  
                            }
                        }

                        break;



                }
            }
            catch (Exception ex)
            {

            }

            return BubbleEvent;
        }

        public void SBO_Application_FormDataEvent(ref BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {

                if (BusinessObjectInfo.ActionSuccess == true)
                {
                    if (!string.IsNullOrEmpty(SBOMain.TFromUID) && SBOMain.FromCnt != null)
                    {
                        string DocEntry = "";
                        SBOMain.GetDocEntryFromXml(BusinessObjectInfo.ObjectKey, ref DocEntry);
                         
                        if (SBOMain.sForm == "LC")
                        {
                            SAPbouiCOM.Form LCTran = SBOMain.SBO_Application.Forms.GetForm(SBOMain.TFromUID, Convert.ToInt32(SBOMain.FromCnt));
                            SAPbouiCOM.Matrix oMatrix = LCTran.Items.Item("matLCEX").Specific;

                            string docnum = objCU.getDocNumFromDocKey("OPCH", DocEntry);
                            string currency = objCU.getCurrFromDocKey("OPCH", DocEntry);
                            double lineTotal = objCU.getLineTotalFromDocKey("PCH1", DocEntry);
                            double rate = objCU.getRateFromDocKey("PCH1", DocEntry);
                             

                            ((SAPbouiCOM.EditText)oMatrix.Columns.Item("lc5bdn").Cells.Item(SBOMain.LineNo).Specific).Value = docnum;
                            ((SAPbouiCOM.EditText)oMatrix.Columns.Item("lc5bden").Cells.Item(SBOMain.LineNo).Specific).Value = DocEntry.ToString(); //oDataTable.GetValue("DocDate", 0).ToString();
                            ((SAPbouiCOM.EditText)oMatrix.Columns.Item("lc5pl").Cells.Item(SBOMain.LineNo).Specific).Value = lineTotal.ToString();
                            ((SAPbouiCOM.EditText)oMatrix.Columns.Item("lc5cur").Cells.Item(SBOMain.LineNo).Specific).Value = currency;
                            ((SAPbouiCOM.EditText)oMatrix.Columns.Item("lc5rt").Cells.Item(SBOMain.LineNo).Specific).Value = rate.ToString();
                        }
                        else if (SBOMain.sForm == "ET")
                        {
                            SAPbouiCOM.Form ExTrans = SBOMain.SBO_Application.Forms.GetForm(SBOMain.TFromUID, Convert.ToInt32(SBOMain.FromCnt));
                            SAPbouiCOM.Matrix oMatrix = ExTrans.Items.Item("matEXPAC").Specific;

                             string docnum = objCU.getDocNumFromDocKey("OPCH", DocEntry);
                             string currency = objCU.getCurrFromDocKey("OPCH", DocEntry);
                             double lineTotal = objCU.getLineTotalFromDocKey("PCH1", DocEntry);
                             double rate = objCU.getRateFromDocKey("PCH1", DocEntry);  

                            ((SAPbouiCOM.EditText)oMatrix.Columns.Item("et3edn").Cells.Item(SBOMain.LineNo).Specific).Value = docnum;
                            ((SAPbouiCOM.EditText)oMatrix.Columns.Item("et3ede").Cells.Item(SBOMain.LineNo).Specific).Value = DocEntry.ToString(); //oDataTable.GetValue("DocDate", 0).ToString();
                            ((SAPbouiCOM.EditText)oMatrix.Columns.Item("et3pl").Cells.Item(SBOMain.LineNo).Specific).Value = lineTotal.ToString();
                            ((SAPbouiCOM.EditText)oMatrix.Columns.Item("et3cur").Cells.Item(SBOMain.LineNo).Specific).Value = currency;
                            ((SAPbouiCOM.EditText)oMatrix.Columns.Item("et3rt").Cells.Item(SBOMain.LineNo).Specific).Value = rate.ToString();
                        }

                        SBOMain.TFromUID = "";
                        SBOMain.FromCnt = null;
                        SBOMain.LineNo = null;
                        SBOMain.sForm = "";
                    }
                }
            }
            catch (Exception ex)
            {

            }
        }
    }
}
