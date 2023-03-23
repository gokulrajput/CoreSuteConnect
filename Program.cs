using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CoreSuteConnect
{
    internal static class Program
    {
        //internal static 
        //Selection_Port_Data oPort = new Selection_Port_Data();
        public static Selection_Port_Data lstStyleCode = new Selection_Port_Data();
        public static Seletion_LC_Data LCTransData = new Seletion_LC_Data();
        public static Seletion_EXIM_Data ExTransData = new Seletion_EXIM_Data();
        public static Selection_ETRFLLN_Data ETRFLLN_Data = new Selection_ETRFLLN_Data();

        public static Selection_EximExp_Data ExExpData = new Selection_EximExp_Data();
        public static Selection_LCExp_Data LCExpData = new Selection_LCExp_Data();

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            SBOMain obj = new SBOMain();

            // Following line will continue Application execution after creating object
            System.Windows.Forms.Application.Run();


        }
    }

    public class Selection_LCExp_Data
    {
        public string LCExpMat { get; set; }
        public int LCExpMatRow { get; set; }
        public string LCExpMatVal { get; set; } 
    }
    public class Selection_EximExp_Data
    {
        public string EXExpMat { get; set; }
        public int EXExpMatRow { get; set; }
        public string EXExpMatCol { get; set; }
        public string EXExpMatVal { get; set; } 

    }
    public class Selection_ETRFLLN_Data
    {
        public string ETMat { get; set; }
        public int ETMatRow { get; set; }
        public string ETMatCol { get; set; }
        public string ETMatVal { get; set; }
        public string ETMatItemcode { get; set; }
    }
    public class Selection_Port_Data
    {
        public string PortSelect { get; set; }
        public string PortCode { get; set; }
    }
    public class Seletion_LC_Data
    {
        public string LcNo { get; set; }
        public string DocNum { get; set; }
        public string DocEntry { get; set; }
    }
    public class Seletion_EXIM_Data
    {
        public string ExNo { get; set; }
        public string DocNum { get; set; }
        public string DocEntry { get; set; }
    }
}