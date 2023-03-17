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

namespace CoreSuteConnect.Class.AUTOEMAIL
{
    class clsDocWiseSetup
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
    }
}
