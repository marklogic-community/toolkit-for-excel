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

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            UserControl1 uc = (UserControl1)Globals.ThisAddIn.myPane.Control;
            string url = "http://localhost:8023/ppt/api/upload.xqy?uid="; //add to config and get from there
            string filename="";
            string saveasdir = uc.getTempPath();

            PPT.Presentation pptx = Globals.ThisAddIn.Application.ActivePresentation;
            string path = pptx.Path;


            if ((pptx.Name == null || pptx.Name.Equals("") || pptx.Path == null || pptx.Path.Equals("")))
            {
                filename = uc.useSaveFileDialog();
                //System.Windows.Forms.MessageBox.Show("name: "+filename);
            }
            else
            {
                filename = pptx.Name;
                //System.Windows.Forms.MessageBox.Show("name2: " + filename);
            }

            url = url + "/" + filename;

            if (!(filename.Equals("") || filename == null))
                uc.saveWithImages(saveasdir,filename, url);

        }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            UserControl1 uc = (UserControl1)Globals.ThisAddIn.myPane.Control;
            string url = "http://localhost:8023/ppt/api/upload.xqy?uid=";
            
            string filename = uc.useSaveFileDialog();
            string saveasdir = uc.getTempPath();

            url = url + "/" + filename;

            if (!(filename.Equals("") || filename == null))
               uc.saveWithImages(saveasdir, filename,url);
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
