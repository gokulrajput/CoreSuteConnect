using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CoreSuteConnect.Events;
using SAPbouiCOM;
using SAPbobsCOM;
using CoreSuteConnect.Class.Common;
using System.Xml;
using System.Net.Http.Headers;
using CoreSuteConnect.Class.QC;

namespace CoreSuteConnect.Class.QC
{
    public class SBOMainQC
    {
        clsSampleRequest clsSampleRequest = new clsSampleRequest(null);

        public bool SubItemEvent(ref ItemEvent pVal, ref bool BubbleEvent, string FormUID)
        {
            BubbleEvent = true;
            try
            {
                switch (pVal.FormTypeEx)
                {
                    case "frmQCSample":
                        clsSampleRequest.ItemEvent(ref pVal, ref BubbleEvent, pVal.FormUID);
                        break;
                }
            }
            catch (Exception ex)
            {
                //SBO_Application.MessageBox(ex.Message.ToString(), 1, "ok", "", "");
            }
            return BubbleEvent;
        }

    }
}
