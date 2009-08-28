/*Copyright 2009 Mark Logic Corporation

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
 * 
 * Ribbon1.cs - the Ribbon callbacks for tab and Button menu items
 * 
*/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using PPT = Microsoft.Office.Interop.PowerPoint;

namespace MarkLogic_PowerPointAddin
{
    public partial class Ribbon1 : OfficeRibbon
    {
        public Ribbon1()
        {
            InitializeComponent();
        }

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void toggleButton1_Click(object sender, RibbonControlEventArgs e)
        {  
            Globals.ThisAddIn.myPane.Visible = ((RibbonToggleButton)sender).Checked;

        }

        private void saveasbutton_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                UserControl1 uc = (UserControl1)Globals.ThisAddIn.myPane.Control;
                AddinConfiguration ac = AddinConfiguration.GetInstance();

                string url = ac.getWebURL() + ac.getRibbonXqy();
                string user = ac.getUser();
                string pwd = ac.getAuth();
                string filename = "";
                string saveasdir = uc.getTempPath();

                PPT.Presentation pptx = Globals.ThisAddIn.Application.ActivePresentation;
                string path = pptx.Path;


                if ((pptx.Name == null || pptx.Name.Equals("") || pptx.Path == null || pptx.Path.Equals("")))
                {
                    filename = uc.useSaveFileDialog();
                }
                else
                {
                    filename = pptx.Name;
                }

                //url = url + "/" + filename;

                if (!(filename.Equals("") || filename == null))
                    uc.saveActivePresentationAndImages(saveasdir, filename, url, user, pwd);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("An Error occurred when trying to save the presentation.\n\n" + ex.Message);
            }

        }

        private void savebutton_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                UserControl1 uc = (UserControl1)Globals.ThisAddIn.myPane.Control;
                AddinConfiguration ac = AddinConfiguration.GetInstance();
                string url = ac.getWebURL() + ac.getRibbonXqy();
                string user = ac.getUser();
                string pwd = ac.getAuth();

                string filename = uc.useSaveFileDialog();
                string saveasdir = uc.getTempPath();

                //url = url +"/" + filename;

                if (!(filename.Equals("") || filename == null))
                    uc.saveActivePresentationAndImages(saveasdir, filename, url, user, pwd);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("An error occurred when trying to save the presentation.\n\n" + ex.Message);
            }
        }

       /* private string useSaveFileDialog()
        {
            Prompt p = new Prompt();
            p.ShowDialog();
            string filename = p.pfilename;
            //System.Windows.Forms.MessageBox.Show(filename);
            return filename;
        }
        * */
    }


}
