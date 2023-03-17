using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CoreSuteConnect.Class.QC.QCDB
{
    internal class QCDB
    {
        //CreateDB cd = new CreateDB();
        public void createQCDB()
        {
            if (CreateDataBase() == true)
            {
                //System.Windows.Forms.MessageBox.Show("Add-On database tables created successfully!");
            }
        }
        private bool CreateDataBase()
        {
            try
            {  
                return true;
            }
            catch (Exception)
            {
                System.Windows.Forms.MessageBox.Show("Add-On database tables creation fail!");
                return false;
            }
        }
    }
}
