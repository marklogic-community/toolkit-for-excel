using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Win32;
using System.Windows.Forms;

namespace MarkLogic_ExcelAddin
{


    class AddinConfiguration
    {
        bool debugMsg = false;
        private static AddinConfiguration instance;

        private string webUrl = "";
        private string rbnTabLbl = "";
        private string rbnBtnLbl = "";
        private string rbnGrpLbl = "";
        private string ctpTtlLbl = "";

        private string isPaneEnabled = "";
        private bool paneEnabled = false;

        private AddinConfiguration()
        {
            //MessageBox.Show("IN INITIALIZE");
            initializeConfig();
        }

        public static AddinConfiguration GetInstance()
        {

            if (instance == null)
            {
                instance = new AddinConfiguration();
            }

            return instance;
        }

        private void initializeConfig()
        {


            RegistryKey regKey1 = Registry.CurrentUser;
            regKey1 = regKey1.OpenSubKey(@"MarkLogicAddinConfiguration\Excel");

            if (regKey1 == null)
            {
                if (debugMsg)
                    MessageBox.Show("KEY IS  NULL");

            }
            else
            {
                if (debugMsg)
                    MessageBox.Show("KEY IS: " + regKey1.GetValue("URL"));

                webUrl = (string)regKey1.GetValue("URL");
                ctpTtlLbl = (string)regKey1.GetValue("CTPTitle");
                rbnTabLbl = (string)regKey1.GetValue("RbnTabLbl");
                rbnGrpLbl = (string)regKey1.GetValue("RbnGrpLbl");
                rbnBtnLbl = (string)regKey1.GetValue("RbnBtnLbl");
                isPaneEnabled = (string)regKey1.GetValue("CTPEnabled");

                if (isPaneEnabled.ToUpper().Equals("TRUE"))
                {
                    paneEnabled = true;
                }

            }

        }

        public string getWebURL()
        {
            // MessageBox.Show("URL: " + webUrl);

            return webUrl;
        }

        public string getRibbonTabLabel()
        {
            return rbnTabLbl;
        }

        public string getRibbonGroupLabel()
        {
            return rbnGrpLbl;
        }

        public string getRibbonButtonLabel()
        {
            return rbnBtnLbl;
        }

        public string getCTPTitleLabel()
        {
            return ctpTtlLbl;
        }

        public bool getPaneEnabled()
        {
            return paneEnabled;
        }
    }





}