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
using System.Collections.Specialized;
 
namespace CoreSuteConnect.Class.DEFAULTSAPFORMS
{
    class clsARInvoice
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

        public int getEximTrackingNoBasedonDocEntry(string fieldname, string docentry, string cardcode)
        {
            string getDocEntry1 = "SELECT DocNum FROM dbo.[@EXET] WHERE "+fieldname+"  = '" + docentry + "' AND U_exbc='" + cardcode + "'";
            SAPbobsCOM.Recordset rec1 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            rec1.DoQuery(getDocEntry1);
            return rec1.RecordCount;
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
                            // Validation : Without Add / Update not allow to perform Copy to.
                            if (pVal.ItemUID == "btnExTrk" && (oForm.Mode == BoFormMode.fm_ADD_MODE))
                            {
                                if (String.IsNullOrEmpty( oForm.Items.Item("4").Specific.Value.ToString()))
                                {
                                    SBOMain.SBO_Application.StatusBar.SetText("Please select Customer.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                    BubbleEvent = false;
                                }
                                SAPbouiCOM.Matrix matrix = oForm.Items.Item("38").Specific;
                                if (matrix.RowCount == 1)
                                {
                                    SBOMain.SBO_Application.StatusBar.SetText("Please select Items.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                    BubbleEvent = false; 
                                } 
                            }
                        }
                        if (pVal.BeforeAction == false)
                        {
                            if (pVal.ItemUID == "btnExTrk" && (oForm.Mode == BoFormMode.fm_OK_MODE))
                            {
                                string status = oForm.Items.Item("81").Specific.Value.ToString();

                                /*if (status != "1")
                                {
                                    SBOMain.SBO_Application.StatusBar.SetText("A/R Invoice is not opened.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                } 
                                else
                                {*/
                                    OutwardToEximTracking outEximTracking = new OutwardToEximTracking();
                                    //string DocNum = string.Empty; string Series = string.Empty;
                                    SAPbouiCOM.EditText CardCode = oForm.Items.Item("4").Specific;
                                    SAPbouiCOM.EditText CardName = oForm.Items.Item("54").Specific;
                                    SAPbouiCOM.EditText DocNum = oForm.Items.Item("8").Specific; 

                                    SAPbouiCOM.ComboBox Series = oForm.Items.Item("88").Specific;
                                      
                                    string getDocEntry = "SELECT DocEntry FROM OINV WHERE DocNum='" + DocNum.Value + "' AND Series='" + Series.Selected.Value + "' AND CardCode='" + CardCode.Value + "'";
                                    SAPbobsCOM.Recordset rec = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                    rec.DoQuery(getDocEntry);
                                    if (rec.RecordCount > 0)
                                    {
                                        string DocEntry = Convert.ToString(rec.Fields.Item("DocEntry").Value);

                                        int ETN = getEximTrackingNoBasedonDocEntry("U_exinvnode", DocEntry, CardCode.Value);

                                        if (ETN > 0)
                                        {
                                            outEximTracking.DocEntry = DocEntry;
                                            outEximTracking.FromFrmName = "ARInvoiceEXIST"; 
                                            objCU.FormLoadAndActivate("frmETTrans", "mnsmEXIM013");
                                            clsExTrans oPrice = new clsExTrans(outEximTracking);
                                            oForm.Close();
                                        }  
                                        else
                                        {
                                            string getBaseEntry = "SELECT DISTINCT(BaseEntry) FROM INV1 WHERE DocEntry = '" + DocEntry + "' AND BaseType = '15'";
                                            SAPbobsCOM.Recordset rec2 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                            rec2.DoQuery(getBaseEntry);
                                            if (rec2.RecordCount > 0)
                                            {
                                                outEximTracking.DelDocEnt = Convert.ToString(rec2.Fields.Item("BaseEntry").Value);

                                                ETN = getEximTrackingNoBasedonDocEntry("U_exdnde", Convert.ToString(rec2.Fields.Item("BaseEntry").Value), CardCode.Value);
                                                if(ETN > 0)
                                                {
                                                    outEximTracking.BPName = CardName.Value;
                                                    outEximTracking.BPCode = CardCode.Value;
                                                    outEximTracking.DocEntry = DocEntry;
                                                    outEximTracking.DocNum = DocNum.Value;
                                                    outEximTracking.FromFrmName = "DeliveryEXISTARNot";
                                                    objCU.FormLoadAndActivate("frmETTrans", "mnsmEXIM013");
                                                    clsExTrans oPrice1 = new clsExTrans(outEximTracking);
                                                    oForm.Close();
                                                }
                                                else
                                                {
                                                    string getBaseEntry2 = "SELECT DocNum FROM ODLN WHERE  DocEntry = '" + rec2.Fields.Item("BaseEntry").Value + "'";
                                                    SAPbobsCOM.Recordset rec3 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                                    rec3.DoQuery(getBaseEntry2);
                                                    if (rec3.RecordCount > 0)
                                                    {
                                                        outEximTracking.DelDocNo = Convert.ToString(rec3.Fields.Item("DocNum").Value);
                                                    }

                                                    string getBaseEntry16 = "SELECT DISTINCT(BaseEntry) FROM DLN1 WHERE DocEntry = '" + rec2.Fields.Item("BaseEntry").Value + "' AND BaseType = '17'";
                                                    SAPbobsCOM.Recordset rec16 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                                    rec16.DoQuery(getBaseEntry16);
                                                    if (rec16.RecordCount > 0)
                                                    {
                                                        outEximTracking.SODocEnt = Convert.ToString(rec16.Fields.Item("BaseEntry").Value);

                                                        string getBaseEntry20 = "SELECT DocNum FROM ORDR WHERE  DocEntry = '" + rec16.Fields.Item("BaseEntry").Value + "'";
                                                        SAPbobsCOM.Recordset rec20 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                                        rec20.DoQuery(getBaseEntry20);
                                                        if (rec20.RecordCount > 0)
                                                        {
                                                            outEximTracking.SODocNo = Convert.ToString(rec20.Fields.Item("DocNum").Value);
                                                        }
                                                    }
                                                }

                                                 

                                            }
                                            else
                                            {
                                                string getBaseEntry1 = "SELECT DISTINCT(BaseEntry) FROM INV1 WHERE DocEntry = '" + DocEntry + "' AND BaseType = '17'";
                                                SAPbobsCOM.Recordset rec5 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                                rec5.DoQuery(getBaseEntry1);
                                                if (rec5.RecordCount > 0)
                                                {
                                                    outEximTracking.SODocEnt = Convert.ToString(rec5.Fields.Item("BaseEntry").Value);

                                                    string getBaseEntry12 = "SELECT DocNum FROM ORDR WHERE DocEntry = '" + rec5.Fields.Item("BaseEntry").Value + "'";
                                                    SAPbobsCOM.Recordset rec13 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                                    rec13.DoQuery(getBaseEntry12);
                                                    if (rec13.RecordCount > 0)
                                                    {
                                                        outEximTracking.SODocNo = Convert.ToString(rec13.Fields.Item("DocNum").Value);
                                                    }
                                                }
                                            }
                                            outEximTracking.BPName = CardName.Value;
                                            outEximTracking.BPCode = CardCode.Value;
                                            outEximTracking.DocNum = DocNum.Value;
                                            outEximTracking.DocDate = oForm.Items.Item("10").Specific.Value;
                                            outEximTracking.DocEntry = DocEntry;
                                            outEximTracking.FromFrmName = "ARInvoice";

                                            SAPbouiCOM.Form oUDFForm = SBOMain.SBO_Application.Forms.Item(oForm.UDFFormUID);
                                            outEximTracking.Incoterm = oUDFForm.Items.Item("U_EXIMINCO").Specific.value;
                                            outEximTracking.PrecarriageBy = oUDFForm.Items.Item("U_EXIMPCGB").Specific.value;
                                            outEximTracking.PrecarrierBy = oUDFForm.Items.Item("U_EXIMPCRB").Specific.value;
                                            outEximTracking.Portofloading = oUDFForm.Items.Item("U_EXIMPOL").Specific.value;
                                            outEximTracking.Portlfdischarge = oUDFForm.Items.Item("U_EXIMPOD").Specific.Value;
                                            outEximTracking.Portofreceipt = oUDFForm.Items.Item("U_EXIMPOR").Specific.Value;
                                            outEximTracking.Finaldestination = oUDFForm.Items.Item("U_EXIMFD").Specific.Value;
                                            outEximTracking.Countryoforigin = oUDFForm.Items.Item("U_EXIMOC").Specific.Value;
                                            outEximTracking.DestinationCountry = oUDFForm.Items.Item("U_EXIMFDC").Specific.Value;


                                            objCU.FormLoadAndActivate("frmETTrans", "mnsmEXIM013"); 
                                            clsExTrans oPrice = new clsExTrans(outEximTracking);
                                            oForm.Close();
                                        }
                                    }
                               // } 
                                  
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

         
    }
}
