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
using CoreSuteConnect.Class.PRICELIST;

namespace CoreSuteConnect.Class.AUTOEMAIL
{
    class clsReportWiseSetup
    {
        #region VariableDeclaration

        public static SAPbouiCOM.Application SBO_Application;

        private SAPbouiCOM.Form oForm;
        private SAPbouiCOM.Matrix oMatrix;
        public string cFormID = string.Empty;

        private SAPbouiCOM.ChooseFromList oCFL = null;
        private SAPbouiCOM.Conditions oConds = null;
        private SAPbouiCOM.Condition oCond = null;

        #endregion VariableDeclaration

        public bool MenuEvent(ref SAPbouiCOM.MenuEvent pVal, string FormId, string Type = "")
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
                    if ((oForm.Mode == BoFormMode.fm_ADD_MODE || Type == "ADDNEWFORM") && Type != "DEL_ROW")
                    {
                        //Form_Load_Components(oForm);
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

        public void Form_Load_Components(SAPbouiCOM.Form oForm)
        {
            SetCode();   
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
}
