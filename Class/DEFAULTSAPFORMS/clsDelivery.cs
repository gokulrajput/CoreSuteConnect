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

namespace CoreSuteConnect.Class.DEFAULTSAPFORMS
{
    class clsDelivery
    {

        #region VariableDeclaration
        private SAPbouiCOM.Form oForm;
        
        CommonUtility objCU = new CommonUtility();

        #endregion VariableDeclaration

        public bool MenuEvent(ref SAPbouiCOM.MenuEvent pVal, string FormId)
        {
            bool BubbleEvent = true;
            try
            {
                 
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
            string getDocEntry1 = "SELECT DocNum FROM dbo.[@EXET] WHERE " + fieldname + "  = '" + docentry + "' AND U_exbc='" + cardcode + "'";
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
                        //oForm = SBOMain.SBO_Application.Forms.ActiveForm;
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

                                string status = oForm.Items.Item("81").Specific.Value.ToString();

                               /* if (status == "1")
                                {
                                    SBOMain.SBO_Application.StatusBar.SetText("Delivery is not opened.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                }
                                else
                                {*/

                                    OutwardToEximTracking outEximTracking = new OutwardToEximTracking();
                                    //string DocNum = string.Empty; string Series = string.Empty;
                                    SAPbouiCOM.EditText CardCode = oForm.Items.Item("4").Specific;
                                    SAPbouiCOM.EditText CardName = oForm.Items.Item("54").Specific;
                                    SAPbouiCOM.EditText DocNum = oForm.Items.Item("8").Specific;
                                    SAPbouiCOM.ComboBox Series = oForm.Items.Item("88").Specific;

                                    string getDocEntry = "SELECT DocEntry FROM ODLN WHERE DocNum='" + DocNum.Value + "' AND Series='" + Series.Selected.Value + "' AND CardCode='" + CardCode.Value + "'";
                                    SAPbobsCOM.Recordset rec = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                    rec.DoQuery(getDocEntry);
                                    if (rec.RecordCount > 0)
                                    {
                                        string DocEntry = Convert.ToString(rec.Fields.Item("DocEntry").Value);

                                        string getDocEntry1 = "SELECT DocNum FROM dbo.[@EXET] WHERE U_exdnde = '" + DocEntry + "' AND U_exbc='" + CardCode.Value + "'";
                                        SAPbobsCOM.Recordset rec1 = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                        rec1.DoQuery(getDocEntry1);
                                        if (rec1.RecordCount > 0)
                                        {
                                                outEximTracking.DelDocEnt = DocEntry;
                                                outEximTracking.FromFrmName = "DeliveryEXIST"; 
                                                objCU.FormLoadAndActivate("frmETTrans", "mnsmEXIM013");
                                                clsExTrans oPrice = new clsExTrans(outEximTracking);
                                                oForm.Close();

                                          // SBOMain.SBO_Application.StatusBar.SetText("Delivery is linked with Exim Transaction no: '" + Convert.ToString(rec1.Fields.Item("DocNum").Value) + "'.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                        }
                                        else
                                        {
                                            outEximTracking.BPName = CardName.Value;
                                            outEximTracking.BPCode = CardCode.Value;
                                            outEximTracking.DelDocNo = DocNum.Value;
                                            outEximTracking.DelDocEnt = DocEntry;
                                            outEximTracking.FromFrmName = "Delivery";

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

                                            string getBaseEntry16 = "SELECT DISTINCT(BaseEntry) FROM DLN1 WHERE DocEntry = '" + DocEntry + "' AND BaseType = '17'";
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

                                                /*int ETN = getEximTrackingNoBasedonDocEntry("U_exsonde", Convert.ToString(rec16.Fields.Item("BaseEntry").Value), CardCode.Value);
                                                if (ETN > 0)
                                                { 
                                                    outEximTracking.BPName = CardName.Value;
                                                    outEximTracking.BPCode = CardCode.Value;
                                                    outEximTracking.DocEntry = DocEntry;
                                                    outEximTracking.DocNum = DocNum.Value;
                                                    outEximTracking.FromFrmName = "SOexitDLNot"; 
                                                }*/
                                            }

                                                objCU.FormLoadAndActivate("frmETTrans", "mnsmEXIM013");
                                                clsExTrans oPrice = new clsExTrans(outEximTracking);
                                                oForm.Close();
                                          }
                                    } 
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
