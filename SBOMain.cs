using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CoreSuteConnect.Events;
using SAPbouiCOM;
using SAPbobsCOM;
using CoreSuteConnect.Class.Common;
using CoreSuteConnect.Class.EXIM;
using CoreSuteConnect.Class.PRICELIST;
using CoreSuteConnect.Class.DEFAULTSAPFORMS;
using CoreSuteConnect.Class.QC;
using CoreSuteConnect.Class.AUTOEMAIL;
using CoreSuteConnect.Class.JOBWORK;
using System.Xml;
using System.Net.Http.Headers;
using static CoreSuteConnect.Class.DEFAULTSAPFORMS.clsProductionOrder;

namespace CoreSuteConnect
{
    public class SBOMain
    {
        #region Variable Declaration
        public static SAPbouiCOM.Application SBO_Application;
        public static SAPbobsCOM.Company oCompany = null;
        public static string sPath = null;
        public static string sForm = null;
        public static int? LineNo = null;
        public static string TFromUID = "";
        public static int? FromCnt = null; 

        public SAPbouiCOM.EventFilters oFilters;
        public SAPbouiCOM.EventFilter oFilter;
        public SAPbouiCOM.Form oForm;

        public static int RightClickLineNum = 0;
        public static string RightClickItemID = "";

        SAPbouiCOM.Menus oMenus;
        SAPbouiCOM.MenuItem oMenuItem;
        SAPbouiCOM.MenuCreationParams oCreationPackage = null;

        CreateDB cd = new CreateDB();
        clsLicenceManager clsLicenceManager = new clsLicenceManager();
        clsGeneralSettings clsGeneralSettings = new clsGeneralSettings();
         
        clsPortMaster clsPortMaster = new clsPortMaster(null);
        clsDocMaster clsDocMaster = new clsDocMaster();
        clsExpMaster clsExpMaster = new clsExpMaster(null);
        clsInctMaster clsInctMaster = new clsInctMaster(null);
        clsSchmMaster clsSchmMaster = new clsSchmMaster(null);
        clsSchmTrans clsSchmTrans = new clsSchmTrans();
        clsExTrans clsExTrans = new clsExTrans(null);
        ClsLCTrans ClsLCTrans = new ClsLCTrans(null);
        clsSCTrans clsSCTrans = new clsSCTrans(null);

        clsETTransList clsETTransList = new clsETTransList();
        clsPortList clsPortList = new clsPortList();
        clsEXPList clsEXPList = new clsEXPList();
        clsSchmList clsSchmList = new clsSchmList();

        clsSalesOrder clsSalesOrder = new clsSalesOrder();
        clsDelivery clsDelivery = new clsDelivery();
        clsARInvoice clsARInvoice = new clsARInvoice();
        clsPurchaseRequest clsPurchaseRequest = new clsPurchaseRequest();
        clsPurchaseQuotation clsPurchaseQuotation = new clsPurchaseQuotation();
        clsPurchaseOrder clsPurchaseOrder = new clsPurchaseOrder();
         
        clsReceiptFromProduction clsReceiptFromProduction = new clsReceiptFromProduction();
        clsGoodsReceipt clsGoodsReceipt = new clsGoodsReceipt(null);
        clsGoodsIssue clsGoodsIssue = new clsGoodsIssue(null);
        clsGRPO clsGRPO = new clsGRPO();
        clsAPInvoice clsAPInvoice = new clsAPInvoice();
        clsInvTrans clsInvTrans = new clsInvTrans(null);

        clsProductionOrder clsProductionOrder = new clsProductionOrder(null);  

        clsBPMaster clsBPMaster = new clsBPMaster();

        clsOutwards clsOutwards = new clsOutwards(null);
        clsPriceList clsPriceList = new clsPriceList(null);
        clsFGPrice clsFGPrice = new clsFGPrice();


        clsSampleRequest clsSampleRequest = new clsSampleRequest(null);
        clsEmailAutomation clsEmailAutomation = new clsEmailAutomation();

        SBOMainQC SBOMainQC = new SBOMainQC();
        SBOMainAUTMAIL SBOMainAUTMAIL = new SBOMainAUTMAIL();

        clsJWIn clsJWIn = new clsJWIn(null);
        clsJWOut clsJWOut = new clsJWOut(null);
        clsJWNPM clsJWNPM = new clsJWNPM(null); 
         
        #endregion

        #region Constuctor
        public SBOMain()
        {
            try
            {
                SetApplication();
                SetConnection();
                cd.createDB();

                sPath = System.Windows.Forms.Application.StartupPath.ToString();
                SBO_Application.MenuEvent += new SAPbouiCOM._IApplicationEvents_MenuEventEventHandler(SBO_Application_MenuEvent);
                SBO_Application.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(SBO_Application_ItemEvent);
                SBO_Application.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(SBO_Application_AppEvent);
                SBO_Application.RightClickEvent += new SAPbouiCOM._IApplicationEvents_RightClickEventEventHandler(SBO_Application_RightClickEvent);
                SBO_Application.FormDataEvent += new SAPbouiCOM._IApplicationEvents_FormDataEventEventHandler(SBO_Application_FormDataEvent);
                
                AddMenus(); // To ADD Menu Items. 
                SBO_Application.StatusBar.SetText("Add-on Core Sute Connected !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
            }
            catch (Exception ex)
            {
                // SBO_Application.MessageBox(ex.Message.ToString(), 1, "ok", "", "");
            }
        }

        #region MenuEvent
        private void SBO_Application_MenuEvent(ref MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (pVal.BeforeAction == true)
                {
                    // 519 For preview Layout
                    // 520 for Print
                    // 6657 For Email

                    // 7170 Word
                    // 7176 PDF

                    // 4865 Query Manager
                    // 5895 Layout Designer
                    // 1304 Refresh
                    // 5890 For Form Settings
                    // 523 Launch Application
                    // 524 Lock Screen


                    // For REMOVE
                    if (pVal.MenuUID == "1283")
                    {
                        oForm = SBOMain.SBO_Application.Forms.ActiveForm;

                        switch (oForm.TypeEx)
                        {
                            case "frmETTrans":
                                BubbleEvent = clsExTrans.MenuEvent(ref pVal, oForm.UniqueID, "REMOVE");
                                break;
                            case "frmLCTrans":
                                BubbleEvent = ClsLCTrans.MenuEvent(ref pVal, oForm.UniqueID, "REMOVE");
                                break;
                            case "frmSchmMaster":
                                BubbleEvent = clsSchmMaster.MenuEvent(ref pVal, oForm.UniqueID, "REMOVE");
                                break;
                            case "frmJWout":
                                BubbleEvent = clsJWOut.MenuEvent(ref pVal, oForm.UniqueID, "REMOVE");
                                break;
                            case "frmJWIn":
                                BubbleEvent = clsJWIn.MenuEvent(ref pVal, oForm.UniqueID, "REMOVE");
                                break;
                            default:
                                break;
                        }
                    }
                    // For Duplicate
                    if (pVal.MenuUID == "1287")
                    {
                        oForm = SBOMain.SBO_Application.Forms.ActiveForm;

                        switch (oForm.TypeEx)
                        {
                            case "frmPriceList":
                                clsPriceList.MenuEvent(ref pVal, oForm.UniqueID, "Duplicate");
                                break;
                            case "frmFGPRice":
                                clsFGPrice.MenuEvent(ref pVal, oForm.UniqueID, "Duplicate");
                                break;
                            default:
                                break;
                        }
                    }
                    // DELETE ROW
                    if (pVal.MenuUID == "1293")
                    {
                        oForm = SBOMain.SBO_Application.Forms.ActiveForm;
                        switch (oForm.TypeEx)
                        {
                            case "frmPriceList":
                                clsPriceList.MenuEvent(ref pVal, oForm.UniqueID, "DEL_ROW");
                                break;
                            case "frmFGPRice":
                                clsFGPrice.MenuEvent(ref pVal, oForm.UniqueID, "DEL_ROW");
                                break;
                            case "frmSCTrans":
                                clsSCTrans.MenuEvent(ref pVal, oForm.UniqueID, "DEL_ROW");
                                break;
                            case "frmETTrans":
                                BubbleEvent = clsExTrans.MenuEvent(ref pVal, oForm.UniqueID, "DEL_ROW");
                                break;
                            case "frmLCTrans":
                                ClsLCTrans.MenuEvent(ref pVal, oForm.UniqueID, "DEL_ROW");
                                break;
                            case "frmSchmMaster":
                                clsSchmMaster.MenuEvent(ref pVal, oForm.UniqueID, "DEL_ROW");
                                break;
                            case "frmJWout":
                                clsJWOut.MenuEvent(ref pVal, oForm.UniqueID, "DEL_ROW");
                                break;
                            case "frmJWIn":
                               clsJWIn.MenuEvent(ref pVal, oForm.UniqueID, "DEL_ROW");
                                break; 
                            default:
                                break;
                        }
                    }
                }
                if (pVal.BeforeAction == false)
                {

                    // Core Sute License Manager
                    if (pVal.MenuUID == "mnsmCLSA")
                    {
                        LoadXmlandSelectForm("frmLicence", "Common");
                        oForm = SBOMain.SBO_Application.Forms.ActiveForm;
                        clsLicenceManager.MenuEvent(ref pVal, oForm.UniqueID);
                    }
                    // Core Sute General Settings
                    else if (pVal.MenuUID == "mnsmGNST")
                    {
                        LoadXmlandSelectForm("frmGenSettings", "Common");
                        oForm = SBOMain.SBO_Application.Forms.ActiveForm;
                        clsGeneralSettings.MenuEvent(ref pVal, oForm.UniqueID);
                    }

                    // For Preview Layout
                    if (pVal.MenuUID == "519")
                    {
                        oForm = SBOMain.SBO_Application.Forms.ActiveForm;
                        switch (oForm.TypeEx)
                        {
                            case "frmQCSample":
                                clsSampleRequest.MenuEvent(ref pVal, oForm.UniqueID, "previewLayout");
                                break;
                            default:
                                break;
                        }
                    }

                    // FOR HEADER Find Button click :
                    if (pVal.MenuUID == "1281")
                    {
                        oForm = SBOMain.SBO_Application.Forms.ActiveForm;
                        switch (oForm.TypeEx)
                        {
                            case "frmSchmMaster":
                                clsSchmMaster.MenuEvent(ref pVal, oForm.UniqueID, "FIND");
                                break;
                            default:
                                break;
                        }
                    }

                    // For ADD NEW MENU HEADER
                    if (pVal.MenuUID == "1282")
                    {
                        oForm = SBOMain.SBO_Application.Forms.ActiveForm;
                        switch (oForm.TypeEx)
                        {
                            case "frmPriceList":
                                clsPriceList.MenuEvent(ref pVal, oForm.UniqueID, "ADDNEWFORM");
                                break;
                            case "frmFGPRice":
                                clsFGPrice.MenuEvent(ref pVal, oForm.UniqueID, "ADDNEWFORM");
                                break;
                            case "frmSCTrans":
                                clsSCTrans.MenuEvent(ref pVal, oForm.UniqueID, "ADDNEWFORM");
                                break;
                            case "frmETTrans":
                                clsExTrans.MenuEvent(ref pVal, oForm.UniqueID, "ADDNEWFORM");
                                break;
                            case "frmLCTrans":
                                ClsLCTrans.MenuEvent(ref pVal, oForm.UniqueID, "ADDNEWFORM");
                                break;
                            case "frmSchmMaster":
                                clsSchmMaster.MenuEvent(ref pVal, oForm.UniqueID, "ADDNEWFORM");
                                break;
                            case "frmPortMaster":
                                clsPortMaster.MenuEvent(ref pVal, oForm.UniqueID, "ADDNEWFORM");
                                break;
                            case "frmInctMaster":
                                clsInctMaster.MenuEvent(ref pVal, oForm.UniqueID, "ADDNEWFORM");
                                break;
                            case "frmDocMaster":
                                clsDocMaster.MenuEvent(ref pVal, oForm.UniqueID, "ADDNEWFORM");
                                break;
                            case "frmExpMaster":
                                clsExpMaster.MenuEvent(ref pVal, oForm.UniqueID, "ADDNEWFORM");
                                break;
                            case "frmQCSample":
                                clsSampleRequest.MenuEvent(ref pVal, oForm.UniqueID, "ADDNEWFORM");
                                break; 
                            case "frmJWNPM":
                                clsJWNPM.MenuEvent(ref pVal, oForm.UniqueID, "ADDNEWFORM");
                                break;
                            case "frmJWout":
                                clsJWOut.MenuEvent(ref pVal, oForm.UniqueID, "ADDNEWFORM");
                                break;
                            case "frmJWIn":
                                clsJWIn.MenuEvent(ref pVal, oForm.UniqueID, "ADDNEWFORM");
                                break;

                            default:
                                break;
                        }
                    }

                    // FOR HEADER NAVIGATION : 
                    if (pVal.MenuUID == "1290" || pVal.MenuUID == "1288" || pVal.MenuUID == "1291" || pVal.MenuUID == "1289")
                    {
                        oForm = SBOMain.SBO_Application.Forms.ActiveForm;
                        switch (oForm.TypeEx)
                        {
                            case "frmETTrans":
                                clsExTrans.MenuEvent(ref pVal, oForm.UniqueID, "navigation");
                                break;
                            case "frmSchmMaster":
                                clsSchmMaster.MenuEvent(ref pVal, oForm.UniqueID, "navigation");
                                break;
                            case "frmLCTrans":
                                ClsLCTrans.MenuEvent(ref pVal, oForm.UniqueID, "navigation");
                                break;
                            case "frmQCSample":
                                clsSampleRequest.MenuEvent(ref pVal, oForm.UniqueID, "navigation");
                                break;
                            case "frmJWNPM":
                                clsJWNPM.MenuEvent(ref pVal, oForm.UniqueID, "navigation");
                                break;
                            case "frmJWout":
                                clsJWOut.MenuEvent(ref pVal, oForm.UniqueID, "navigation");
                                break;
                            case "frmJWIn":
                                clsJWIn.MenuEvent(ref pVal, oForm.UniqueID, "navigation");
                                break;

                            default:
                                break;
                        }

                    }

                    // For ADD Row
                    if (pVal.MenuUID == "1292")
                    {
                        oForm = SBOMain.SBO_Application.Forms.ActiveForm;
                        switch (oForm.TypeEx)
                        {
                            case "frmPriceList":
                                clsPriceList.MenuEvent(ref pVal, oForm.UniqueID, "ADD_ROW");
                                break;
                            case "frmFGPRice":
                                clsFGPrice.MenuEvent(ref pVal, oForm.UniqueID, "ADD_ROW");
                                break;
                            case "frmSCTrans":
                                clsSCTrans.MenuEvent(ref pVal, oForm.UniqueID, "ADD_ROW");
                                break;
                            case "frmETTrans":
                                clsExTrans.MenuEvent(ref pVal, oForm.UniqueID, "ADD_ROW");
                                break;
                            case "frmLCTrans":
                                ClsLCTrans.MenuEvent(ref pVal, oForm.UniqueID, "ADD_ROW");
                                break;
                            case "frmSchmMaster":
                                clsSchmMaster.MenuEvent(ref pVal, oForm.UniqueID, "ADD_ROW");
                                break;
                            case "frmQCSample":
                                clsSampleRequest.MenuEvent(ref pVal, oForm.UniqueID, "ADD_ROW");
                                break;

                            case "frmJWIn":
                                clsJWIn.MenuEvent(ref pVal, oForm.UniqueID, "ADD_ROW");
                                break;
                            case "frmJWout":
                                clsJWOut.MenuEvent(ref pVal, oForm.UniqueID, "ADD_ROW");
                                break;
                            default:
                                break;
                        }
                    }

                    // DELETE ROW
                    if (pVal.MenuUID == "1293")
                    {
                        oForm = SBOMain.SBO_Application.Forms.ActiveForm;
                        switch (oForm.TypeEx)
                        {
                            case "frmPriceList":
                                clsPriceList.MenuEvent(ref pVal, oForm.UniqueID, "DEL_ROW");
                                break;
                            case "frmFGPRice":
                                clsFGPrice.MenuEvent(ref pVal, oForm.UniqueID, "DEL_ROW");
                                break;
                            case "frmSCTrans":
                                clsSCTrans.MenuEvent(ref pVal, oForm.UniqueID, "DEL_ROW");
                                break;
                            case "frmETTrans":
                                clsExTrans.MenuEvent(ref pVal, oForm.UniqueID, "DEL_ROW");
                                break;
                            case "frmLCTrans":
                                ClsLCTrans.MenuEvent(ref pVal, oForm.UniqueID, "DEL_ROW");
                                break;
                            case "frmSchmMaster":
                                clsSchmMaster.MenuEvent(ref pVal, oForm.UniqueID, "DEL_ROW");
                                break;
                            case "frmQCSample":
                                clsSampleRequest.MenuEvent(ref pVal, oForm.UniqueID, "DEL_ROW");
                                break;
                            case "frmJWIn":
                                clsJWIn.MenuEvent(ref pVal, oForm.UniqueID, "DEL_ROW");
                                break;
                            case "frmJWOut":
                                clsJWOut.MenuEvent(ref pVal, oForm.UniqueID, "DEL_ROW");
                                break;

                            default:
                                break;
                        }
                    }

                    // For Duplicate
                    if (pVal.MenuUID == "1287")
                    {
                        oForm = SBOMain.SBO_Application.Forms.ActiveForm;
                        switch (oForm.TypeEx)
                        {
                            case "frmPriceList":
                                clsPriceList.MenuEvent(ref pVal, oForm.UniqueID, "Duplicate");
                                break;
                            case "frmFGPRice":
                                clsFGPrice.MenuEvent(ref pVal, oForm.UniqueID, "Duplicate");
                                break;
                            default:
                                break;
                        }
                    }


                    //frmSCTrans
                    if (pVal.MenuUID == "mnsmEXIM007")
                    {
                        LoadXmlandSelectForm("frmPortMaster", "EXIM");
                        oForm = SBOMain.SBO_Application.Forms.ActiveForm;
                        clsPortMaster.MenuEvent(ref pVal, oForm.UniqueID);

                    }
                    else if (pVal.MenuUID == "mnsmEXIM008")
                    {
                        LoadXmlandSelectForm("frmInctMaster", "EXIM");
                        oForm = SBOMain.SBO_Application.Forms.ActiveForm;
                        clsInctMaster.MenuEvent(ref pVal, oForm.UniqueID);
                    }
                    else if (pVal.MenuUID == "mnsmEXIM009")
                    {
                        LoadXmlandSelectForm("frmDocMaster", "EXIM");
                        oForm = SBOMain.SBO_Application.Forms.ActiveForm;
                        clsDocMaster.MenuEvent(ref pVal, oForm.UniqueID);
                    }
                    else if (pVal.MenuUID == "mnsmEXIM010")
                    {
                        LoadXmlandSelectForm("frmExpMaster", "EXIM");
                        oForm = SBOMain.SBO_Application.Forms.ActiveForm;
                        clsExpMaster.MenuEvent(ref pVal, oForm.UniqueID);
                    }
                    else if (pVal.MenuUID == "mnsmEXIM011")
                    {
                        LoadXmlandSelectForm("frmSchmMaster", "EXIM");
                        oForm = SBOMain.SBO_Application.Forms.ActiveForm;
                        clsSchmMaster.MenuEvent(ref pVal, oForm.UniqueID);
                    }
                    else if (pVal.MenuUID == "mnsmEXIM012")
                    {
                        LoadXmlandSelectForm("frmLCTrans", "EXIM");
                        oForm = SBOMain.SBO_Application.Forms.ActiveForm;
                        ClsLCTrans.MenuEvent(ref pVal, oForm.UniqueID);
                    }
                    else if (pVal.MenuUID == "mnsmEXIM003")
                    {
                        LoadXmlandSelectForm("frmSCTrans", "EXIM");
                        oForm = SBOMain.SBO_Application.Forms.ActiveForm;
                        clsSCTrans.MenuEvent(ref pVal, oForm.UniqueID);
                    }
                    else if (pVal.MenuUID == "mnsmEXIM013")
                    {
                        LoadXmlandSelectForm("frmETTrans", "EXIM");
                        oForm = SBOMain.SBO_Application.Forms.ActiveForm;
                        clsExTrans.MenuEvent(ref pVal, oForm.UniqueID);
                    }
                    else if (pVal.MenuUID == "mnsmJW001")
                    {
                        LoadXmlandSelectForm("frmJWout", "JOBWORK"); 
                        oForm = SBOMain.SBO_Application.Forms.ActiveForm;
                        clsJWOut.MenuEvent(ref pVal, oForm.UniqueID);
                    } 
                    else if (pVal.MenuUID == "mnsmJW002")
                    {
                        LoadXmlandSelectForm("frmJWIn", "JOBWORK"); 
                        oForm = SBOMain.SBO_Application.Forms.ActiveForm;
                        clsJWIn.MenuEvent(ref pVal, oForm.UniqueID);
                    }
                    else if (pVal.MenuUID == "mnsmJW003")
                    {
                        LoadXmlandSelectForm("frmJWNPM", "JOBWORK");
                        oForm = SBOMain.SBO_Application.Forms.ActiveForm;
                        clsJWNPM.MenuEvent(ref pVal, oForm.UniqueID);
                    } 

                    else if (pVal.MenuUID == "mnsmQC003")
                    {
                        LoadXmlandSelectForm("frmQCMM", "QC");
                    }
                    else if (pVal.MenuUID == "mnsmQC004")
                    {
                        LoadXmlandSelectForm("frmQCSM", "QC");
                    }
                    else if (pVal.MenuUID == "mnsmQC005")
                    {
                        LoadXmlandSelectForm("frmQCPM", "QC");
                    }
                    else if (pVal.MenuUID == "mnsmQC006")
                    {
                        LoadXmlandSelectForm("frmQCPMM", "QC");
                    }
                    else if (pVal.MenuUID == "mnsmQC007")
                    {
                        LoadXmlandSelectForm("frmQCSample", "QC");
                        oForm = SBOMain.SBO_Application.Forms.ActiveForm;
                        clsSampleRequest.MenuEvent(ref pVal, oForm.UniqueID);
                    }
                    else if (pVal.MenuUID == "mnsmQC008")
                    {
                        LoadXmlandSelectForm("frmQCQC", "QC");
                    }
                    else if (pVal.MenuUID == "mnsmQC009")
                    {
                        LoadXmlandSelectForm("frmQCQA", "QC");
                    }
                    else if (pVal.MenuUID == "mnsmGP001")
                    {
                        LoadXmlandSelectForm("frmEMPS", "GATEPASS");
                    }
                    else if (pVal.MenuUID == "mnsmGP002")
                    {
                        LoadXmlandSelectForm("frmVIPS", "GATEPASS");
                    }
                    else if (pVal.MenuUID == "mnsmGP003")
                    {
                        LoadXmlandSelectForm("frmLGAP", "GATEPASS");
                    }
                    else if (pVal.MenuUID == "mnsmMB")
                    {
                        LoadXmlandSelectForm("frmMB", "MULTIBOM");
                    }
                    else if (pVal.MenuUID == "mnsmPL001")
                    {
                        LoadXmlandSelectForm("frmLMItemsList", "PRICELIST");
                        oForm = SBOMain.SBO_Application.Forms.ActiveForm;
                        clsOutwards.MenuEvent(ref pVal, oForm.UniqueID);
                    }
                    else if (pVal.MenuUID == "mnsmPL002")
                    {
                        LoadXmlandSelectForm("frmPriceList", "PRICELIST");
                        oForm = SBOMain.SBO_Application.Forms.ActiveForm;
                        clsPriceList.MenuEvent(ref pVal, oForm.UniqueID);
                        //clsPriceList.SBO_Application_MenuEvent(ref pVal,out BubbleEvent);
                    }
                    else if (pVal.MenuUID == "mnsmPL003")
                    {
                        LoadXmlandSelectForm("frmFGPrice", "PRICELIST");
                        oForm = SBOMain.SBO_Application.Forms.ActiveForm;
                        clsFGPrice.MenuEvent(ref pVal, oForm.UniqueID);
                    }

                    else if (pVal.MenuUID == "mnsmAE001")
                    {
                        LoadXmlandSelectForm("frmEmailAutomation", "AUTOEMAIL");
                        oForm = SBOMain.SBO_Application.Forms.ActiveForm;
                        clsEmailAutomation.MenuEvent(ref pVal, oForm.UniqueID);
                    } 
                }
            }
            catch (Exception ex)
            {
                //SBO_Application.MessageBox(ex.Message.ToString(), 1, "ok", "", "");
            }
            //throw new NotImplementedException();
        }

        private void SBO_Application_ItemEvent(string FormUID, ref ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true; 
            try
            {
                switch (pVal.FormTypeEx)
                {
                    // Form ID
                    case "frmLicence":
                        clsLicenceManager.ItemEvent(ref pVal, ref BubbleEvent, pVal.FormUID);
                        break;
                    case "frmGenSettings":
                        clsGeneralSettings.ItemEvent(ref pVal, ref BubbleEvent, pVal.FormUID);
                        break;
                    case "139": // Sals Order
                        clsSalesOrder.ItemEvent(ref pVal, ref BubbleEvent, pVal.FormUID);
                        break;
                    case "140": // Delivery
                        clsDelivery.ItemEvent(ref pVal, ref BubbleEvent, pVal.FormUID);
                        break;
                    case "133": // A/R Invoice
                        clsARInvoice.ItemEvent(ref pVal, ref BubbleEvent, pVal.FormUID);
                        break;
                    case "1470000200": // Purchase Request
                        clsPurchaseRequest.ItemEvent(ref pVal, ref BubbleEvent, pVal.FormUID);
                        break;
                    case "540000988": // Purchase Quotation
                        clsPurchaseQuotation.ItemEvent(ref pVal, ref BubbleEvent, pVal.FormUID);
                        break;
                    case "142": // Purchase Order
                        clsPurchaseOrder.ItemEvent(ref pVal, ref BubbleEvent, pVal.FormUID);
                        break; 
                    case "143": // GRPO
                        clsGRPO.ItemEvent(ref pVal, ref BubbleEvent, pVal.FormUID);
                        break;
                    case "141": // A/P Invoice
                        clsAPInvoice.ItemEvent(ref pVal, ref BubbleEvent, pVal.FormUID);
                        break;
                    case "720": // Goods Issue : 3079 menu
                        clsGoodsIssue.ItemEvent(ref pVal, ref BubbleEvent, pVal.FormUID);
                        break;
                    case "65214": // Receipt from Production 
                        clsReceiptFromProduction.ItemEvent(ref pVal, ref BubbleEvent, pVal.FormUID);
                        break;
                    case "721": // Goods Receipt : 3079 menu
                        clsGoodsReceipt.ItemEvent(ref pVal, ref BubbleEvent, pVal.FormUID);
                        break;
                    case "940": // Goods Receipt : 3079 menu
                        clsInvTrans.ItemEvent(ref pVal, ref BubbleEvent, pVal.FormUID);
                        break;
                    case "-134": // Business Partner Master Data
                        clsBPMaster.ItemEvent(ref pVal, ref BubbleEvent, pVal.FormUID);
                        break;

                    case "frmDB":
                        //cd.itemevent(ref pVal, pVal.FormUID);
                        break;
                    case "frmPortMaster":
                        clsPortMaster.ItemEvent(ref pVal, ref BubbleEvent, pVal.FormUID);
                        break;
                    case "frmInctMaster":
                        clsInctMaster.ItemEvent(ref pVal, ref BubbleEvent, pVal.FormUID);
                        break;
                    case "frmDocMaster":
                        clsDocMaster.ItemEvent(ref pVal, ref BubbleEvent, pVal.FormUID);
                        break;
                    case "frmExpMaster":
                        clsExpMaster.ItemEvent(ref pVal, ref BubbleEvent, pVal.FormUID);
                        break;
                    case "frmSchmMaster":
                        clsSchmMaster.ItemEvent(ref pVal, ref BubbleEvent, pVal.FormUID);
                        break;
                    case "frmSCTrans":
                        clsSCTrans.ItemEvent(ref pVal, ref BubbleEvent, pVal.FormUID);
                        break;
                    case "frmLCTrans":
                        ClsLCTrans.ItemEvent(ref pVal, ref BubbleEvent, pVal.FormUID);
                        break;
                    case "frmETTrans":
                        clsExTrans.ItemEvent(ref pVal, ref BubbleEvent, pVal.FormUID);
                        break;
                    case "frmETTransList":
                        clsETTransList.ItemEvent(ref pVal, ref BubbleEvent, pVal.FormUID);
                        break;
                    case "frmPortList":
                        clsPortList.ItemEvent(ref pVal, ref BubbleEvent, pVal.FormUID);
                        break;

                    case "frmSchmList":
                        clsSchmList.ItemEvent(ref pVal, ref BubbleEvent, pVal.FormUID);
                        break;
                    case "frmLMItemsList":
                        clsOutwards.ItemEvent(ref pVal, ref BubbleEvent, pVal.FormUID);
                        break;
                    case "frmPriceList":
                        clsPriceList.ItemEvent(ref pVal, ref BubbleEvent, pVal.FormUID);
                        break;
                    case "frmExpList":
                        clsEXPList.ItemEvent(ref pVal, ref BubbleEvent, pVal.FormUID);
                        break;
                    case "frmFGPRice":
                        clsFGPrice.ItemEvent(ref pVal, ref BubbleEvent, pVal.FormUID);
                        break; 

                        // FOR QC ADDON
                    case "frmQCSample":
                        SBOMainQC.SubItemEvent(ref pVal, ref BubbleEvent, pVal.FormUID);
                        break;
                        
                        // FOR EMAIL AUTOMATION ADDON
                   case "frmEmailAutomation":
                        SBOMainAUTMAIL.SubItemEvent(ref pVal, ref BubbleEvent, pVal.FormUID); 
                        break;

                        // FOR JOBWORK ADDON
                    case "frmJWNPM":
                        clsJWNPM.ItemEvent(ref pVal, ref BubbleEvent, pVal.FormUID);
                        break;
                    case "frmJWout":
                        clsJWOut.ItemEvent(ref pVal, ref BubbleEvent, pVal.FormUID);
                        break;
                    case "frmJWIn":
                        clsJWIn.ItemEvent(ref pVal, ref BubbleEvent, pVal.FormUID);
                        break;
                }
            }
            catch (Exception ex)
            {
                //SBO_Application.MessageBox(ex.Message.ToString(), 1, "ok", "", "");
            }
        }

        private void SBO_Application_FormDataEvent(ref BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true; 
            try
            {
                switch (BusinessObjectInfo.Type)
                {
                    // In Case Statement we are settign Object ID.

                    case "1470000113": // Purchase Request
                        clsPurchaseRequest.SBO_Application_FormDataEvent(ref BusinessObjectInfo, out BubbleEvent);
                        break;
                    case "540000006": // Purchase Quotation
                        clsPurchaseQuotation.SBO_Application_FormDataEvent(ref BusinessObjectInfo, out BubbleEvent);
                        break;
                    case "22": // Purchase Order
                        clsPurchaseOrder.SBO_Application_FormDataEvent(ref BusinessObjectInfo, out BubbleEvent);
                        break;
                    case "17": // Purchase Order
                        clsSalesOrder.SBO_Application_FormDataEvent(ref BusinessObjectInfo, out BubbleEvent);
                        break;
                    case "18": // A/P Invoice
                        clsAPInvoice.SBO_Application_FormDataEvent(ref BusinessObjectInfo, out BubbleEvent);
                        break;
                    case "60": // Goods Issue
                        clsGoodsIssue.SBO_Application_FormDataEvent(ref BusinessObjectInfo, out BubbleEvent);
                        break;
                    case "59": // Goods Receipt
                        clsGoodsReceipt.SBO_Application_FormDataEvent(ref BusinessObjectInfo, out BubbleEvent);
                        break;
                    case "67": // Inventory Transfer
                        clsInvTrans.SBO_Application_FormDataEvent(ref BusinessObjectInfo, out BubbleEvent);
                        break;
                }
            }
            catch (Exception ex)
            {  

            }
        }
        #endregion

        private void SBO_Application_RightClickEvent(ref ContextMenuInfo eventInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            RightClickLineNum = eventInfo.Row;
            RightClickItemID = eventInfo.ItemUID;
            //throw new NotImplementedException();
            BubbleEvent = true;
            try
            {
                oForm = SBOMain.SBO_Application.Forms.Item(eventInfo.FormUID);
                if ((eventInfo.FormUID == "frmETTrans") && ( 
                    (eventInfo.ItemUID == "matEXPAC") || (eventInfo.ItemUID == "matFOB")  || (eventInfo.ItemUID == "matRLF1") || (eventInfo.ItemUID == "matRLF2") ||
                    (eventInfo.ItemUID == "matETSC1") || (eventInfo.ItemUID == "matETSC2") || (eventInfo.ItemUID == "matEXVGM") || (eventInfo.ItemUID == "matEXCA") ||
                    (eventInfo.ItemUID == "matRS"))
                    )
                {
                    oForm.EnableMenu("1292", true);
                    oForm.EnableMenu("1293", true);
                } 
                else if ((eventInfo.FormUID == "frmSchmMaster") && 
                          ((eventInfo.ItemUID == "matEXOB") ||   (eventInfo.ItemUID == "matEXFL") || (eventInfo.ItemUID == "matEXIR") ||
                           (eventInfo.ItemUID == "matEXIU") ||   (eventInfo.ItemUID == "matHSN")  || (eventInfo.ItemUID == "matATTACH")))
                {
                    oForm.EnableMenu("1292", true);
                    oForm.EnableMenu("1293", true); 
                }
                else if ((eventInfo.FormUID == "frmLCTrans") && ((eventInfo.ItemUID == "matLCLD") || (eventInfo.ItemUID == "matLCDOC") ||
                         (eventInfo.ItemUID == "matLCEX") ||  (eventInfo.ItemUID == "matLCAMED") || (eventInfo.ItemUID == "matLCATT")))
                {
                    oForm.EnableMenu("1292", true);
                    oForm.EnableMenu("1293", true); 
                }
                else if ((eventInfo.FormUID == "frmSCTrans") && ((eventInfo.ItemUID == "matSTDET") || (eventInfo.ItemUID == "matDIT")))
                {
                    oForm.EnableMenu("1292", true);
                    oForm.EnableMenu("1293", true); 
                }
                else if ((eventInfo.FormUID == "frmQCSample") && (eventInfo.ItemUID == "matContent"))
                {
                    oForm.EnableMenu("1292", true);
                    oForm.EnableMenu("1293", true);
                }
                else if ((eventInfo.FormUID == "frmJWout") && ((eventInfo.ItemUID == "matFGTR") ||  (eventInfo.ItemUID == "matCMTR")))
                {
                    oForm.EnableMenu("1292", true);
                    oForm.EnableMenu("1293", true);
                }
                else if ((eventInfo.FormUID == "frmJWIn") && ((eventInfo.ItemUID == "matFGTR") || (eventInfo.ItemUID == "matCMTR")))
                {
                    oForm.EnableMenu("1292", true);
                    oForm.EnableMenu("1293", true);
                }
                else
                {
                    oForm.EnableMenu("1292", false);
                    oForm.EnableMenu("1293", false);
                }
               
            }
            catch (Exception ex)
            {

            }
        }
        
        public static void GetDocEntryFromXml(string xml, ref string DocKey)
        {
            XmlDocument xmlDoc = new XmlDocument();
            try
            {
                xmlDoc.LoadXml(xml);
                //xmlDoc.LoadXml(xml);
                DocKey = (xmlDoc.ChildNodes[xmlDoc.ChildNodes.Count - 1]).LastChild.LastChild.Value;
                XmlNode xn = (xmlDoc.ChildNodes[xmlDoc.ChildNodes.Count - 1]).LastChild.LastChild;
                //PrimaryKey = xn.ParentNode.Name;
                //PrimaryKey = (xmlDoc.ChildNodes[xmlDoc.ChildNodes.Count - 1]).LastChild.LastChild.;
            }
            catch (Exception ex)
            {

            }
            finally
            {
                xmlDoc = null;
            }
            //return string.Empty;
        }
        
        
        private void LoadXmlandSelectForm(string formID, string FolderPAth)
        {
            for (int i = 1; i < SBO_Application.Forms.Count; i++)
            {
                if (SBO_Application.Forms.Item(i).UniqueID == formID)
                {
                    SBO_Application.Forms.Item(i).Select();
                    return;
                }
            }
            LoadFromXML(formID, FolderPAth);
        }

       
        public static void LoadFromXML(string FileName, string ModulName)
        {
            try
            {
                System.Xml.XmlDocument oXmlDoc = null;
                oXmlDoc = new System.Xml.XmlDocument();
                // load the content of the XML File
                string sPath = null;
                FileName = ModulName + @"\" + FileName + ".xml";
                sPath = System.Windows.Forms.Application.StartupPath;
                oXmlDoc.Load(sPath + @"\Forms\" + FileName);
                // load the form to the SBO application in one batch
                string tmpStr;
                tmpStr = oXmlDoc.InnerXml;
                SBO_Application.LoadBatchActions(ref tmpStr);
                sPath = SBO_Application.GetLastBatchResults();
                // oForm = SBO_Application.Forms.ActiveForm;
            }
            catch (Exception ex)
            {
                //SBO_Application.MessageBox(ex.Message.ToString(), 1, "ok", "", "");
            }

        }

        #endregion


        #region ConnectToSAP
        private void SetApplication()
        {
            SAPbouiCOM.SboGuiApi SboGuiApi = null;
            string sConnectionString = null;

            SboGuiApi = new SAPbouiCOM.SboGuiApi();
            sConnectionString = System.Convert.ToString(Environment.GetCommandLineArgs().GetValue(1));
            try
            {
                // If there's no active application the connection will fail
                SboGuiApi.Connect(sConnectionString);
            }
            catch
            {   //  If Connection failed
                System.Windows.Forms.MessageBox.Show("SAP Business One Application is not running");
                System.Environment.Exit(0);
            }
            // get an initialized application object
            SBO_Application = SboGuiApi.GetApplication(-1);
        }
        public void SetConnection()
        {
            try
            {
                string sCookie;
                string sConnectionContext;

                // First initialize the Company object
                oCompany = new SAPbobsCOM.Company();
                if (oCompany.Connected == true)
                    return;

                sCookie = oCompany.GetContextCookie();
                sConnectionContext = SBO_Application.Company.GetConnectionContext(sCookie);
                oCompany.SetSboLoginContext(sConnectionContext);
                oCompany.Connect();
            }
            catch (Exception ex)
            {
                //SBO_Application.MessageBox(ex.Message.ToString(), 1, "ok", "", "");
            }
        }
        #endregion

        #region MenuCreation
        private void AddMenus()
        {
            try
            {
                oMenus = null;
                oMenuItem = null;
                oCreationPackage = null;
                oMenus = SBO_Application.Menus;
                oCreationPackage = ((SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)));

                oMenuItem = SBO_Application.Menus.Item("43524");
                AddMenu_Items(-1, "", SAPbouiCOM.BoMenuType.mt_STRING, "Core Sute License Administrator", "mnsmCLSA", oMenus);

                oMenuItem = SBO_Application.Menus.Item("8192");
                AddMenu_Items(2, "", SAPbouiCOM.BoMenuType.mt_STRING, "Core Sute General Settings", "mnsmGNST", oMenus);

                // Price List
                oMenuItem = SBO_Application.Menus.Item("2048");
                AddMenu_Items(-1, "", SAPbouiCOM.BoMenuType.mt_POPUP, "Price List", "mnsmPL", oMenus);

                oMenuItem = SBO_Application.Menus.Item("mnsmPL");
                AddMenu_Items(0, "", SAPbouiCOM.BoMenuType.mt_STRING, "Outwards", "mnsmPL001", oMenus);
                AddMenu_Items(1, "", SAPbouiCOM.BoMenuType.mt_STRING, "Price List", "mnsmPL002", oMenus);
                AddMenu_Items(2, "", SAPbouiCOM.BoMenuType.mt_STRING, "FG Price Master", "mnsmPL003", oMenus);
                 
                // Exim Addon Menu Items
                oMenuItem = SBO_Application.Menus.Item("2048");
                AddMenu_Items(-1, "", SAPbouiCOM.BoMenuType.mt_POPUP, "EXIM", "mnsmEXIM", oMenus);

                oMenuItem = SBO_Application.Menus.Item("mnsmEXIM");
                AddMenu_Items(1, "", SAPbouiCOM.BoMenuType.mt_POPUP, "Masters", "mnsmEXIM001", oMenus);
                AddMenu_Items(2, "", SAPbouiCOM.BoMenuType.mt_POPUP, "Transactions", "mnsmEXIM002", oMenus);

                oMenuItem = SBO_Application.Menus.Item("mnsmEXIM001");
                AddMenu_Items(1, "", SAPbouiCOM.BoMenuType.mt_STRING, "Port Master", "mnsmEXIM007", oMenus);
                AddMenu_Items(2, "", SAPbouiCOM.BoMenuType.mt_STRING, "Incoterm Master", "mnsmEXIM008", oMenus);
                AddMenu_Items(3, "", SAPbouiCOM.BoMenuType.mt_STRING, "Document Master", "mnsmEXIM009", oMenus);
                AddMenu_Items(4, "", SAPbouiCOM.BoMenuType.mt_STRING, "Expense Master", "mnsmEXIM010", oMenus);
                AddMenu_Items(5, "", SAPbouiCOM.BoMenuType.mt_STRING, "Scheme Master", "mnsmEXIM011", oMenus);

                oMenuItem = SBO_Application.Menus.Item("mnsmEXIM002");
                AddMenu_Items(1, "", SAPbouiCOM.BoMenuType.mt_STRING, "Latter of Credit", "mnsmEXIM012", oMenus);
                AddMenu_Items(2, "", SAPbouiCOM.BoMenuType.mt_STRING, "Exim Tracking", "mnsmEXIM013", oMenus);
                AddMenu_Items(3, "", SAPbouiCOM.BoMenuType.mt_STRING, "Scheme Transaction", "mnsmEXIM003", oMenus);

                //Auto - Email
                oMenuItem = SBO_Application.Menus.Item("3328");
                AddMenu_Items(-1, "", SAPbouiCOM.BoMenuType.mt_POPUP, "Auto-Email", "mnsmAE", oMenus);

                oMenuItem = SBO_Application.Menus.Item("mnsmAE");
                AddMenu_Items(1, "", SAPbouiCOM.BoMenuType.mt_STRING, "General Settings", "mnsmAE001", oMenus);
 
                //Jobwork Addon Menu Items
                oMenuItem = SBO_Application.Menus.Item("4352");
                AddMenu_Items(-1, "", SAPbouiCOM.BoMenuType.mt_POPUP, "Jobwork", "mnsmJW", oMenus);

               oMenuItem = SBO_Application.Menus.Item("mnsmJW");
                AddMenu_Items(1, "", SAPbouiCOM.BoMenuType.mt_STRING, "Nature of Processing Master", "mnsmJW003", oMenus);
                AddMenu_Items(2, "", SAPbouiCOM.BoMenuType.mt_STRING, "Jobwork Challan : Out", "mnsmJW001", oMenus);
                AddMenu_Items(3, "", SAPbouiCOM.BoMenuType.mt_STRING, "Jobwork Challan : In", "mnsmJW002", oMenus);
 
                // QC Addon Menu Items
                oMenuItem = SBO_Application.Menus.Item("2304");
                AddMenu_Items(-1, "", SAPbouiCOM.BoMenuType.mt_POPUP, "QC", "mnsmQC", oMenus);

                oMenuItem = SBO_Application.Menus.Item("mnsmQC");
                 AddMenu_Items(2, "", SAPbouiCOM.BoMenuType.mt_POPUP, "Transactions", "mnsmQC002", oMenus);
                 
                oMenuItem = SBO_Application.Menus.Item("mnsmQC002");
                AddMenu_Items(1, "", SAPbouiCOM.BoMenuType.mt_STRING, "Quality Control Check", "mnsmQC007", oMenus); 

                /*
                               // Gate Entry Module
                               oMenuItem = SBO_Application.Menus.Item("43544");
                               AddMenu_Items(-1, "", SAPbouiCOM.BoMenuType.mt_POPUP, "Gatepass", "mnsmGP", oMenus);
                               AddMenu_Items(-1, "", SAPbouiCOM.BoMenuType.mt_POPUP, "Admin", "mnsmAD", oMenus);

                               oMenuItem = SBO_Application.Menus.Item("mnsmGP");
                               AddMenu_Items(1, "", SAPbouiCOM.BoMenuType.mt_STRING, "Visitor Pass", "mnsmGP001", oMenus);
                               AddMenu_Items(2, "", SAPbouiCOM.BoMenuType.mt_STRING, "Employee Pass", "mnsmGP002", oMenus);
                               oMenuItem = SBO_Application.Menus.Item("mnsmAD");
                               AddMenu_Items(1, "", SAPbouiCOM.BoMenuType.mt_STRING, "Legal Applications", "mnsmGP003", oMenus);

                               oMenuItem = SBO_Application.Menus.Item("4352");
                               AddMenu_Items(0, "", SAPbouiCOM.BoMenuType.mt_STRING, "Multi Bill of Material", "mnsmMB", oMenus);


                               oMenuItem = SBO_Application.Menus.Item("mnsmEXIM003");
                                  AddMenu_Items(1, "", SAPbouiCOM.BoMenuType.mt_STRING, "Advance License", "mnsmEXIM014", oMenus);
                                  AddMenu_Items(2, "", SAPbouiCOM.BoMenuType.mt_STRING, "MEIS License", "mnsmEXIM015", oMenus);
                                  AddMenu_Items(3, "", SAPbouiCOM.BoMenuType.mt_STRING, "EPCG License", "mnsmEXIM016", oMenus);

                                  // AddMenu_Items(-1, "icon.png", SAPbouiCOM.BoMenuType.mt_POPUP, "Core Sute Connect", "mnCNC", oMenus);
                                  // oMenuItem = SBO_Application.Menus.Item("43520");
                                  /* AddMenu_Items(2, "", SAPbouiCOM.BoMenuType.mt_STRING, "Quality Control", "mnsmQC", oMenus);
                                     AddMenu_Items(3, "", SAPbouiCOM.BoMenuType.mt_STRING, "Job Work", "mnsmQC", oMenus);
                                  */
                /* oMenuItem = SBO_Application.Menus.Item("43523");
            AddMenu_Items(3, "", SAPbouiCOM.BoMenuType.mt_STRING, "Create Database", "mnuDB", oMenus);
            oMenuItem = SBO_Application.Menus.Item("43520");
            AddMenu_Items(16, "", SAPbouiCOM.BoMenuType.mt_POPUP, "Demo", "mnuPDEMO", oMenus);
            oMenuItem = SBO_Application.Menus.Item("mnuPDEMO");
            AddMenu_Items(1, "", SAPbouiCOM.BoMenuType.mt_STRING, "Demo", "mnuDEMO", oMenus);
           */
            }
            catch (Exception ex)
            {
                // SBO_Application.MessageBox(ex.Message.ToString(), 1, "ok", "", "");
                throw;
            }
        }

        //, ref string ImageName,ref SAPbouiCOM.BoMenuType Type, ref string MenuLabel,ref string MenuId, ref SAPbouiCOM.Menus oMenus
        private void AddMenu_Items(int Position, string ImageName, SAPbouiCOM.BoMenuType Type, string MenuLabel, string MenuId, SAPbouiCOM.Menus oMenus)
        {
            oCreationPackage.Type = Type;
            oCreationPackage.String = MenuLabel;
            oCreationPackage.UniqueID = MenuId;
            oCreationPackage.Enabled = true;
            oCreationPackage.Position = Position;
            if (SBO_Application.ClientType == SAPbouiCOM.BoClientType.ct_Desktop)
            {
                oCreationPackage.Image = sPath + "\\" + ImageName;
            }
            if (!oMenus.Exists(MenuId))
            {
                oMenus = oMenuItem.SubMenus;
                oMenus.AddEx(oCreationPackage);
            }
        }
        #endregion

        public static void SetCode(String FormId, String OBJ)
        {
            try
            {
                SAPbouiCOM.Form oForm;
                SAPbouiCOM.EditText oEdit;
                string Table = null;
                DateTime now = DateTime.Now;
                oForm = SBO_Application.Forms.ActiveForm;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                oEdit = oForm.Items.Item("tCode").Specific;
                Table = "@" + OBJ;
                GetNextDocNum(ref oEdit, ref Table);

                oForm.Items.Item("tCode").Specific.value = oForm.BusinessObject.GetNextSerialNumber("DocEntry", OBJ);
                oForm.Items.Item("tCode").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, (int)BoFormMode.fm_ADD_MODE, SAPbouiCOM.BoModeVisualBehavior.mvb_False);

                /*if (OBJ == "EXRU")
                {
                    Events.Series.SeriesCombo(OBJ, "tSeries");
                    oForm.Items.Item("tDocNum").Specific.value = oForm.BusinessObject.GetNextSerialNumber("DocEntry", OBJ);
                    oForm.Items.Item("tSeries").DisplayDesc = true;
                }*/
                //oForm.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized;
                //oForm.Items.Item("tDocDate").Specific.value = DateTime.Now.ToString("yyyyMMdd");
                // oForm.Items.Item("tCreator").Specific.value = oCompany.UserName;
                 
            }
            catch (Exception ex)
            {
                SBO_Application.SetStatusBarMessage("Set Code :" + ex, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }
        public static int GetNextDocNum(ref SAPbouiCOM.EditText oEdit, ref string TableName)
        {
            try
            {
                oEdit.Value = string.Empty;
                if (TableName.Trim() != string.Empty)
                {
                    SAPbobsCOM.Recordset oRec = default(SAPbobsCOM.Recordset);
                    oRec = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    oRec.DoQuery("Select Max(\"DocEntry\") From \"" + TableName + "\"");
                    int MaxCode = Convert.ToInt32(oRec.Fields.Item(0).Value);
                    if (MaxCode == 0)
                    {
                        oEdit.Value = "1";
                    }
                    else
                    {
                        if (!oRec.EoF)
                        {
                            oEdit.Value = Convert.ToString(MaxCode + 1);
                        }
                        else
                        {
                            oEdit.Value = "1";
                        }
                    }
                    oRec = null;
                }
                else
                {
                    SBO_Application.StatusBar.SetText("Error on GetNextDocNum : Invalid TableName ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                }
                return 0;
            }
            catch (Exception ex)
            {
                SBO_Application.StatusBar.SetText("Error on GetNextDocNum :" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return -1;
            }
        }

        public static string Get_Attach_Folder_Path()
        {
            string ReturnPath = string.Empty;
            try
            {
                SAPbobsCOM.Recordset oRec = default(SAPbobsCOM.Recordset);
                oRec = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRec.DoQuery("Select \"AttachPath\" from \"OADP\"");
                ReturnPath = Convert.ToString(oRec.Fields.Item(0).Value);

                oRec = null;

                return ReturnPath;
            }
            catch (Exception ex)
            {
                SBO_Application.StatusBar.SetText("Error on Get_Attach_Folder_Path :" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return string.Empty;
            }
        }
        #region ApplicationEvent
        private void SBO_Application_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {
            switch (EventType)
            {
                case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
                    {
                        Environment.Exit(-1);
                    }
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition:
                    {
                        Environment.Exit(-1);
                    }
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged:
                    {
                        Environment.Exit(-1);
                    }
                    break;
            }
        }
        #endregion
    }
}
