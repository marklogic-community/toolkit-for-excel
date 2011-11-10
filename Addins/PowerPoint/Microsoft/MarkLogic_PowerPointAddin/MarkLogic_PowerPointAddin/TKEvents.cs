/*Copyright 2009-2010 MarkLogic Corporation

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
 * TKEvents.cs - event handling.
 * Events caught and signals sent to functions in MarkLogicPowerPointEventSupport.js
 * 
*/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;
using Microsoft.Win32;
using PPT = Microsoft.Office.Interop.PowerPoint;
using System.Windows.Forms;

namespace MarkLogic_PowerPointAddin
{
    public partial class UserControl1 : UserControl
    {

      /*  public TKEvents()
        {

            //public PPT.ApplicationClass  ppta;
            ppta = new Microsoft.Office.Interop.PowerPoint.ApplicationClass();
            // ppta = (PPT.ApplicationClass)Globals.ThisAddIn.Application;
            ppta.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;


            System.Runtime.InteropServices.ComTypes.IConnectionPoint mConnectionPoint;
            System.Runtime.InteropServices.ComTypes.IConnectionPointContainer cpContainer;
            int mCookie;

            cpContainer =
            (System.Runtime.InteropServices.ComTypes.IConnectionPointContainer)ppta;
            Guid guid = typeof(Microsoft.Office.Interop.PowerPoint.EApplication).GUID;
            cpContainer.FindConnectionPoint(ref guid, out mConnectionPoint);
            mConnectionPoint.Advise(this, out mCookie);
        }*/

        [DispId(2001)]
        public void
        WindowSelectionChange(Microsoft.Office.Interop.PowerPoint.Selection Sel)
        {
            //System.Windows.Forms.MessageBox.Show("Window Selection Changed");

            try
            {
                string shapeName = Sel.ShapeRange.Name;
                notifyWindowSelectionChange(shapeName);
            }
            catch (Exception e)
            {
                string donothing_removewarning = e.Message;
            }

        }

        [DispId(2002)]
        public void WindowBeforeRightClick(PPT.Selection Sel, bool Cancel)
        {
            //System.Windows.Forms.MessageBox.Show("WindowBeforeRightClick"+Sel.SlideRange.SlideIndex);
            try
            {
                string slideIndex = Sel.SlideRange.SlideIndex.ToString();
                notifyWindowBeforeRightClick(slideIndex);
            }
            catch (Exception e)
            {
                string donothing_removewarning = e.Message;
            }

        }

        [DispId(2003)]
        public void WindowBeforeDoubleClick(PPT.Selection Sel, bool Cancel)
        {
            //System.Windows.Forms.MessageBox.Show("WindowBeforeDoubleClick" + Sel.SlideRange.SlideIndex);
            try
            {
                string slideIndex = Sel.SlideRange.SlideIndex.ToString();
                notifyWindowBeforeDoubleClick(slideIndex);
            }
            catch (Exception e)
            {
                string donothing_removewarning = e.Message;
            }

        }

        [DispId(2004)]
        public void PresentationClose(PPT.Presentation Pres)
        {
            // Do custom thingy here
            try
            {
                string presoName = Pres.Name;
                notifyPresentationClose(presoName);
            }
            catch (Exception e)
            {
                string donothing_removewarning = e.Message;
            }
            finally
            {

                int count = Globals.ThisAddIn.Application.Presentations.Count;
                //set flag and do not execute in case of copy
                //also check for presentation count, if gt than 1, N have been opened through the search app
                if (count == 1)
                {
                    //System.Windows.Forms.MessageBox.Show("PresentationClose");
                    try
                    {
                        
                      //ppta.Quit()
                        PPT.Presentations presentations= ppta.Presentations;
                        foreach (PPT.Presentation p in presentations)
                        {
                            PPT.Slides slides = p.Slides;
                            foreach (PPT.Slide s in slides)
                            {
                               
                                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(s);
                            }

                            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(slides);
                            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(p);
                           
                            GC.Collect();
                            GC.WaitForPendingFinalizers();
                            GC.Collect();
                            GC.WaitForPendingFinalizers();

                        }
                       
                        System.Runtime.InteropServices.Marshal.FinalReleaseComObject(presentations);
                        System.Runtime.InteropServices.Marshal.FinalReleaseComObject(ppta);
                         
                        //ppta = null;
                    
                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                        GC.Collect();
                        GC.WaitForPendingFinalizers();

                        ppta = null;

                    }
                    catch (Exception e)
                    {
                        string donothing_removewarning = e.Message;
                        MessageBox.Show("MESSAGE" + donothing_removewarning);
                        System.Runtime.InteropServices.Marshal.FinalReleaseComObject(ppta);
                        ppta = null;
                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                    }
                    finally
                    {
                        ppta = null;
                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                    }

                }
                else
                {
                    firePptCloseEvent = true;
                }

            }

        }

        [DispId(2005)]
        public void PresentationSave(PPT.Presentation Pres)
        {
            //System.Windows.Forms.MessageBox.Show("PresentationSave");
            try
            {
                string presoName = Pres.Name;
                notifyPresentationSave(presoName);
            }
            catch (Exception e)
            {
                string donothing_removewarning = e.Message;
            }

        }

        [DispId(2006)]
        public void
        PresentationOpen(PPT.Presentation Pres)
        {
            try
            {
                string presoName = Pres.Name;
                notifyPresentationOpen(presoName);
            }
            catch (Exception e)
            {
                string donothing_removewarning = e.Message;
            }

            // Pres.Application.SlideSelectionChanged += new Microsoft.Office.Interop.PowerPoint.EApplication_SlideSelectionChangedEventHandler(Application_SlideSelectionChanged);
            // Pres.Application.WindowSelectionChange += new Microsoft.Office.Interop.PowerPoint.EApplication_WindowSelectionChangeEventHandler(Application_WindowSelectionChange);
            //System.Windows.Forms.MessageBox.Show("!");

        }

        [DispId(2007)]
        public void NewPresentation(PPT.Presentation Pres)
        {
            //System.Windows.Forms.MessageBox.Show("NewPresentation");
            try
            {
                string presoName = Pres.Name;
                notifyNewPresentation(presoName);
            }
            catch (Exception e)
            {
                string donothing_removewarning = e.Message;
            }

        }

        [DispId(2008)]
        public void PresentationNewSlide(PPT.Slide Sld)
        {
            //System.Windows.Forms.MessageBox.Show("PresentationNewSlide");
            try
            {
                string slideIndex = Sld.SlideIndex.ToString();
                notifyPresentationNewSlide(slideIndex);
            }
            catch (Exception e)
            {
                string donothing_removewarning = e.Message;
            }

        }

        [DispId(2009)]
        public void WindowActivate(PPT.Presentation Pres, PPT.DocumentWindow Wn)
        {
            //System.Windows.Forms.MessageBox.Show("WindowActivate");
            try
            {
                string presoName = Wn.Presentation.Name;
                notifyWindowActivate(presoName);
            }
            catch (Exception e)
            {
                string donothing_removewarning = e.Message;
            }

        }

        [DispId(2010)]
        public void WindowDeactivate(PPT.Presentation Pres, PPT.DocumentWindow Wn)
        {
            //System.Windows.Forms.MessageBox.Show("WindowDeactivate");
            try
            {
                string presoName = Wn.Presentation.Name;
                notifyWindowDeactivate(presoName);
            }
            catch (Exception e)
            {
                string donothing_removewarning = e.Message;
              

            }
        }

        [DispId(2011)]
        public void SlideShowBegin(PPT.SlideShowWindow Wn)
        {
            //System.Windows.Forms.MessageBox.Show("SlideShowBegin");
            try
            {
                string presoName = Wn.Presentation.Name;
                notifySlideShowBegin(presoName);
            }
            catch (Exception e)
            {
                string donothing_removewarning = e.Message;
            }
        }

        [DispId(2012)]
        public void SlideShowNextBuild(PPT.SlideShowWindow Wn)
        {
            //System.Windows.Forms.MessageBox.Show("SlideShowNextBuild");
            try
            {
                string presoName = Wn.Presentation.Name;
                notifySlideShowNextBuild(presoName);
            }
            catch (Exception e)
            {
                string donothing_removewarning = e.Message;
            }
        }

        [DispId(2013)]
        public void SlideShowNextSlide(PPT.SlideShowWindow Wn)
        {
            //System.Windows.Forms.MessageBox.Show("SlideShowNextSlide");
            try
            {
                string presoName = Wn.Presentation.Name;
                notifySlideShowNextSlide(presoName);
            }
            catch (Exception e)
            {
                string donothing_removewarning = e.Message;
            }
        }

        [DispId(2014)]
        public void SlideShowEnd(PPT.Presentation Pres)
        {
            //System.Windows.Forms.MessageBox.Show("SlideShowEnd");
            try
            {
                string presoName = Pres.Name;
                notifySlideShowEnd(presoName);
            }
            catch (Exception e)
            {
                string donothing_removewarning = e.Message;
            }
        }

        [DispId(2015)]
        public void PresentationPrint(PPT.Presentation Pres)
        {

            //System.Windows.Forms.MessageBox.Show("PresentationPrint");
            try
            {
                string presoName = Pres.Name;
                notifyPresentationPrint(presoName);
            }
            catch (Exception e)
            {
                string donothing_removewarning = e.Message;
            }
        }



        [DispId(2016)]
        public void
        SlideSelectionChanged(PPT.SlideRange SldRange)
        {
            //MessageBox.Show("SlideSelectionChange");
            /*
              MessageBox.Show("slideID: " + SldRange.SlideID +          
                            "slideIdx: " + SldRange.SlideIndex + "slideNumber: "+SldRange.SlideNumber +
                            "tagCount: "+ SldRange.Tags.Count);
             */
            notifySlideSelectionChange(SldRange.SlideIndex.ToString());
            //MessageBox.Show("Slide Selection Changed");
        }


        [DispId(2017)]
        public void ColorSchemeChanged(PPT.SlideRange SldRange)
        {
            //System.Windows.Forms.MessageBox.Show("Color Changed");
            try
            {
                string shapeRangeName = SldRange.Name;
                notifyColorSchemeChange(shapeRangeName);
            }
            catch (Exception e)
            {
                string donothing_removewarning = e.Message;
            }
        }

        [DispId(2018)]
        public void PresentationBeforeSave(PPT.Presentation Pres, bool Cancel)
        {
            //System.Windows.Forms.MessageBox.Show("PresentationBeforeSave");
            try
            {
                string presoName = Pres.Name;
                notifyPresentationBeforeSave(presoName);
            }
            catch (Exception e)
            {
                string donothing_removewarning = e.Message;
            }
        }

        [DispId(2019)]
        public void SlideShowNextClick(PPT.SlideShowWindow Wn, PPT.Effect nEffect)
        {
            //System.Windows.Forms.MessageBox.Show("SlideShowNextClick");
            try
            {
                string presoName = Wn.Presentation.Name;
                notifySlideShowNextClick(presoName);
            }
            catch (Exception e)
            {
                string donothing_removewarning = e.Message;
            }
        }

        public void notifyWindowSelectionChange(string shapeName)
        {
            try
            {
                object result = webBrowser1.Document.InvokeScript("windowSelectionChange", new String[] { shapeName });
                string res = result.ToString();

                if (res.StartsWith("error"))
                {
                    MessageBox.Show("windowSelectionChangeErrorJS: " + res);
                }

            }
            catch (Exception e)
            {
                string donothing_removewarning = e.Message;
                //MessageBox.Show("windowSelectionChangeError: " + e.Message);
            }
        }

        public void notifySlideSelectionChange(string jsonSlideRange)
        {
            try
            {
                object result = webBrowser1.Document.InvokeScript("slideSelectionChange", new String[] { jsonSlideRange });
                string res = result.ToString();

                if (res.StartsWith("error"))
                {
                    //MessageBox.Show("slideSelectionChangeError:"+ res);
                }

            }
            catch (Exception e)
            {
                string donothing_removewarning = e.Message;
            }
        }

        public void notifyWindowBeforeRightClick(string slideIndex)
        {
            try
            {
                object result = webBrowser1.Document.InvokeScript("windowBeforeRightClick", new String[] { slideIndex });
                string res = result.ToString();

                if (res.StartsWith("error"))
                {
                    MessageBox.Show("windowBeforeRightClickJS: " + res);
                }

            }
            catch (Exception e)
            {
                string donothing_removewarning = e.Message;
                //MessageBox.Show(donothing_removewarning);
            }
        }

        public void notifyWindowBeforeDoubleClick(string slideIndex)
        {
            try
            {
                object result = webBrowser1.Document.InvokeScript("windowBeforeDoubleClick", new String[] { slideIndex });
                string res = result.ToString();

                if (res.StartsWith("error"))
                {
                    MessageBox.Show("windowBeforeDoubleClickJS: " + res);
                }

            }
            catch (Exception e)
            {
                string donothing_removewarning = e.Message;
                //MessageBox.Show(donothing_removewarning);
            }
        }

        public void notifyPresentationClose(string presoName)
        {
            try
            {
                object result = webBrowser1.Document.InvokeScript("presentationClose", new String[] { presoName });
                string res = result.ToString();

                if (res.StartsWith("error"))
                {
                    //MessageBox.Show("presentationCloseJS: " + res);
                }

            }
            catch (Exception e)
            {
                string donothing_removewarning = e.Message;
            }
        }

        public void notifyPresentationSave(string presoName)
        {
            try
            {
                object result = webBrowser1.Document.InvokeScript("presentationSave", new String[] { presoName });
                string res = result.ToString();

                if (res.StartsWith("error"))
                {
                    //MessageBox.Show("presentationSaveJS: " + res);
                }

            }
            catch (Exception e)
            {
                string donothing_removewarning = e.Message;
            }
        }

        public void notifyPresentationOpen(string presoName)
        {
            try
            {
                object result = webBrowser1.Document.InvokeScript("presentationOpen", new String[] { presoName });
                string res = result.ToString();

                if (res.StartsWith("error"))
                {
                    //MessageBox.Show("presentationOpenJS: " + res);
                }

            }
            catch (Exception e)
            {
                string donothing_removewarning = e.Message;
            }
        }

        public void notifyNewPresentation(string presoName)
        {
            try
            {
                object result = webBrowser1.Document.InvokeScript("newPresentation", new String[] { presoName });
                string res = result.ToString();

                if (res.StartsWith("error"))
                {
                    //MessageBox.Show("newPresentationJS: " + res);
                }

            }
            catch (Exception e)
            {
                string donothing_removewarning = e.Message;
            }
        }

        public void notifyPresentationNewSlide(string slideIndex)
        {
            try
            {
                object result = webBrowser1.Document.InvokeScript("presentationNewSlide", new String[] { slideIndex });
                string res = result.ToString();

                if (res.StartsWith("error"))
                {
                    //MessageBox.Show("presentationNewSlideJS: " + res);
                }

            }
            catch (Exception e)
            {
                string donothing_removewarning = e.Message;
            }
        }

        public void notifyWindowActivate(string presoName)
        {
            try
            {
                object result = webBrowser1.Document.InvokeScript("windowActivate", new String[] { presoName });
                string res = result.ToString();

                if (res.StartsWith("error"))
                {
                    //MessageBox.Show("windowActivateJS: " + res);
                }

            }
            catch (Exception e)
            {
                string donothing_removewarning = e.Message;
            }
        }

        public void notifyWindowDeactivate(string presoName)
        {
            try
            {
                object result = webBrowser1.Document.InvokeScript("windowDeactivate", new String[] { presoName });
                string res = result.ToString();

                if (res.StartsWith("error"))
                {
                    //MessageBox.Show("windowDeactivate: " + res);
                }

            }
            catch (Exception e)
            {
                string donothing_removewarning = e.Message;
            }
        }

        public void notifySlideShowBegin(string presoName)
        {
            try
            {
                object result = webBrowser1.Document.InvokeScript("slideShowBegin", new String[] { presoName });
                string res = result.ToString();

                if (res.StartsWith("error"))
                {
                    //MessageBox.Show("slideShowBeginJS: " + res);
                }

            }
            catch (Exception e)
            {
                string donothing_removewarning = e.Message;
            }
        }

        public void notifySlideShowNextBuild(string presoName)
        {
            try
            {
                object result = webBrowser1.Document.InvokeScript("slideShowNextBuild", new String[] { presoName });
                string res = result.ToString();

                if (res.StartsWith("error"))
                {
                    //MessageBox.Show("slideShowNextBuildJS: " + res);
                }

            }
            catch (Exception e)
            {
                string donothing_removewarning = e.Message;
            }
        }

        public void notifySlideShowNextSlide(string presoName)
        {
            try
            {
                object result = webBrowser1.Document.InvokeScript("slideShowNextSlide", new String[] { presoName });
                string res = result.ToString();

                if (res.StartsWith("error"))
                {
                    //MessageBox.Show("slideShowNextSlideJS: " + res);
                }

            }
            catch (Exception e)
            {
                string donothing_removewarning = e.Message;
            }
        }

        public void notifySlideShowEnd(string presoName)
        {
            try
            {
                object result = webBrowser1.Document.InvokeScript("slideShowEnd", new String[] { presoName });
                string res = result.ToString();

                if (res.StartsWith("error"))
                {
                    //MessageBox.Show("slideShowEndJS: " + res);
                }

            }
            catch (Exception e)
            {
                string donothing_removewarning = e.Message;
            }
        }

        public void notifyPresentationPrint(string presoName)
        {
            try
            {
                object result = webBrowser1.Document.InvokeScript("presentationPrint", new String[] { presoName });
                string res = result.ToString();

                if (res.StartsWith("error"))
                {
                    //MessageBox.Show("presentationPrintJS: " + res);
                }

            }
            catch (Exception e)
            {
                string donothing_removewarning = e.Message;
            }
        }

        public void notifyColorSchemeChange(string shapeRangeName)
        {
            try
            {
                object result = webBrowser1.Document.InvokeScript("colorSchemeChange", new String[] { shapeRangeName });
                string res = result.ToString();

                if (res.StartsWith("error"))
                {
                    //MessageBox.Show("colorSchemeChangeJS: " + res);
                }

            }
            catch (Exception e)
            {
                string donothing_removewarning = e.Message;
            }
        }

        public void notifyPresentationBeforeSave(string presoName)
        {
            try
            {
                object result = webBrowser1.Document.InvokeScript("presentationBeforeSave", new String[] { presoName });
                string res = result.ToString();

                if (res.StartsWith("error"))
                {
                    //MessageBox.Show("presentationBeforeSaveJS: " + res);
                }

            }
            catch (Exception e)
            {
                string donothing_removewarning = e.Message;
            }
        }

        public void notifySlideShowNextClick(string presoName)
        {
            try
            {
                object result = webBrowser1.Document.InvokeScript("slideShowNextClick", new String[] { presoName });
                string res = result.ToString();

                if (res.StartsWith("error"))
                {
                    //MessageBox.Show("slideShowNextClickJS: " + res);
                }

            }
            catch (Exception e)
            {
                string donothing_removewarning = e.Message;
            }
        } 

    }
}
