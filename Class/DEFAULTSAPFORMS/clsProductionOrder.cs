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
using CoreSuteConnect.Class.QC;
using System.Drawing.Drawing2D;

namespace CoreSuteConnect.Class.DEFAULTSAPFORMS
{
    class clsProductionOrder
    {
        #region VariableDeclaration
        private SAPbouiCOM.Form oForm;
        public string cFormID = string.Empty;
        public string DocKey = null;
        public static string getSalesForm = string.Empty;

        CommonUtility objCU = new CommonUtility();

        #endregion VariableDeclaration

        public clsProductionOrder(OutwardToProdOrder outClass)
        {
            try
            {
                if (outClass != null)
                {
                    oForm = SBOMain.SBO_Application.Forms.ActiveForm;
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    string Q1 = @"select DocNum,PostDate from OWOR where DocEntry = '" + outClass.DocEntry + "'";
                    SAPbobsCOM.Recordset r1;
                    r1 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    r1.DoQuery(Q1);
                    if (r1.RecordCount > 0)
                    {
                        oForm.Items.Item("18").Specific.value = r1.Fields.Item("DocNum").Value;
                        DateTime tDocDate = Convert.ToDateTime(r1.Fields.Item("PostDate").Value);
                        oForm.Items.Item("24").Specific.value = tDocDate.ToString("yyyyMMdd");
                        oForm.Items.Item("1").Click();
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
    }

    public class OutwardToProdOrder
    {
        public string DocEntry { get; set; }
    }
}
