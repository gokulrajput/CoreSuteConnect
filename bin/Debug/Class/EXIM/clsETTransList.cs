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
    class clsETTransList
    {
        #region VariableDeclaration

        private SAPbouiCOM.Form oForm;
        SAPbouiCOM.Matrix matrix;

        string getDocEntry1 = null; 

        CommonUtility objCU = new CommonUtility();

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
                            if (pVal.ItemUID == "btnCHS")
                            {
                                try
                                {

                                    SAPbouiCOM.Grid objGrid = oForm.Items.Item("grid").Specific;
                                    OutwardFromEximTracking inEximTracking = new OutwardFromEximTracking();
                                    List<ETTransList> lstETTransList = new List<ETTransList>();
                                    for (int i = 0; i < objGrid.Rows.Count; i++)
                                    {
                                        //((SAPbouiCOM.CheckBoxColumn)objGrid.Columns.Item(0)).Check(i, true);
                                        //bool isChecked = ((SAPbouiCOM.CheckBox)objGrid.Columns.Item().).Checked;
                                        ETTransList oETTransList = new ETTransList(); 

                                        if (objGrid.DataTable.Columns.Item("CHK").Cells.Item(i).Value == "Y")
                                        {
                                            oETTransList.eximtrackingno = Convert.ToString(objGrid.DataTable.Columns.Item("Exim Tracking No").Cells.Item(i).Value);
                                            oETTransList.portcode = objGrid.DataTable.Columns.Item("PortCode").Cells.Item(i).Value;
                                            oETTransList.portname = objGrid.DataTable.Columns.Item("PortName").Cells.Item(i).Value;
                                            oETTransList.invoiceno = objGrid.DataTable.Columns.Item("Invoice no").Cells.Item(i).Value;
                                            oETTransList.invoicedocentry = objGrid.DataTable.Columns.Item("Invoice DocEntry").Cells.Item(i).Value;
                                            oETTransList.invoicedate = objGrid.DataTable.Columns.Item("Invoice Date").Cells.Item(i).Value;
                                            oETTransList.shippingbillno = objGrid.DataTable.Columns.Item("Shipping bill no").Cells.Item(i).Value;
                                            oETTransList.shippingbilldate = objGrid.DataTable.Columns.Item("Shipping bill Date").Cells.Item(i).Value;
                                            oETTransList.totalFOB = objGrid.DataTable.Columns.Item("Total FOB").Cells.Item(i).Value;
                                            oETTransList.appliedamount = objGrid.DataTable.Columns.Item("Applied Amount").Cells.Item(i).Value;
                                            //oETTransList.receiveamount = objGrid.DataTable.Columns.Item("Received Amount").Cells.Item(i).Value;

                                            lstETTransList.Add(oETTransList);
                                           
                                        }
                                    }
                                    oForm.Close();
                                    inEximTracking.ScriptNo = "nofind";
                                    clsSCTrans oPrice = new clsSCTrans(lstETTransList, inEximTracking);
                                    
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

    public class OutwardToSchemeTrans
    {
        public List<ETTransList> ETTransList { get; set; }

    }
    public class ETTransList
    {
        public string eximtrackingno { get; set; }
        public string portcode { get; set; }
        public string portname { get; set; }
        public string invoiceno { get; set; }
        public string invoicedocentry { get; set; }
        public DateTime invoicedate { get; set; }
        public string shippingbillno { get; set; }
        public DateTime shippingbilldate { get; set; }
        public double totalFOB { get; set; }
        public double appliedamount { get; set; }
        public double receiveamount { get; set; }

        /* public string invoiceno { get; set; }
         public string shippingbillno { get; set; }
         public string shippingbilldate { get; set; }
 */
    }
}
