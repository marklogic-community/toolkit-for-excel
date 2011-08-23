using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel;
using System.Data;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Security.Permissions;
using System.IO;
using Office = Microsoft.Office.Core;
using Microsoft.Win32;
using PPT = Microsoft.Office.Interop.PowerPoint;

namespace Poller
{
    class DroidIngest
    {
        String ServerURI = @"xcc://oslo:oslol0g1c@localhost:8031";
        String WebDAV = @"http://localhost:8030";
        String StagingDir = "/staging/";
        String ProdDir = "/paven/";
        String ArchiveDir = "/archive/";

        bool checkExtension(string ext)
        {
            if (
                ext.EndsWith(".ppt")  ||
                ext.EndsWith(".pptx") ||
                ext.EndsWith(".ppsx") ||
                ext.EndsWith(".pptm") ||
                ext.EndsWith(".ppsm") ||
                ext.EndsWith(".potx") ||
                ext.EndsWith(".potm")
                )
                return true;
            else
                return false;

        }


        public void ingestDocs(String stagingUri)
        {
            //string sourceDirectory = @"C:\Users\paven\Desktop\MarkLogic\MLUC_TK_PRESENTATIONS\presos\1";//args[0];
            string openSourcePres = WebDAV + stagingUri;
            string fileName = stagingUri.Substring(StagingDir.Length);
            //string saveSourcePres = WebDAV + ProdDir + fileName;
            string extension = Path.GetExtension(openSourcePres);

            string tmpPath = System.IO.Path.GetTempPath() + @"junk\";

            //MessageBox.Show("TEMP PATH: " + tmpPath);

            //FILES TO BE LOADED BY CONTENT LOADER
            string tmpSourcePres = tmpPath + fileName;
            string tmpSinglesPath = tmpPath + fileName.Replace(extension, "_parts/");
            string tmpImagePathSmall = tmpPath + fileName.Replace(extension, "_BMP_S");
            string tmpImagePathMedium = tmpPath + fileName.Replace(extension, "_BMP_M");
            string tmpImagePathLarge = tmpPath + fileName.Replace(extension, "_BMP_L");

           // string sourceSinglesPath = WebDAV+ProdDir+ fileName.Replace(extension, "_parts/");

            bool exists =  System.IO.Directory.Exists(tmpSinglesPath);

            if (exists)
            {
                Console.WriteLine("DOING NOTHING, DIR EXISTS");
            }
            else
            {
                DirectoryInfo dir = System.IO.Directory.CreateDirectory(tmpSinglesPath);
            }


            Console.WriteLine("SINGLES PATH: " + tmpSinglesPath);
            Console.WriteLine("FILENAME"+fileName);
            Console.WriteLine("SAVESTRING"+tmpSourcePres);
         
            bool debug = true;
            object missing = System.Type.Missing;


            try
            {
                PPT.Application ppt = new PPT.Application();

                    try
                    {
                        string imgdirwithpath = "";
                       // string tmpFileName = "";
                        bool extensionCheck = checkExtension(openSourcePres);
                        if (extensionCheck)
                        {   //C:\Users\paven\Desktop\Toolkits-8031
                            PPT.Presentation pres = ppt.Presentations.Open(openSourcePres /*file.FullName*/, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue, Office.MsoTriState.msoFalse);
                            
                            if (debug)
                            {
                                Console.WriteLine("Saving images for file :   " + openSourcePres);
                            }


                            //save source deck to MarkLogic
                            //need to check type, also if 2003 or earlier,save as 2007(for search) and 2003 format
                         
                            //pres.SaveAs(saveSourcePres, PPT.PpSaveAsFileType.ppSaveAsDefault, Office.MsoTriState.msoFalse);
                            pres.SaveAs(tmpSourcePres, PPT.PpSaveAsFileType.ppSaveAsDefault, Office.MsoTriState.msoFalse);
                            //save images for deck as bmp
                            //export only works locally, so save to temp, then use content loader to load
                            //pipeline will map slides to images (have to modify pipeline for BMP, and multiple sizes of images


                            pres.Export(tmpImagePathLarge, "bmp", 960, 720);
                            //System.IO.File.Move(imgdirwithpath + @"/Slide1.BMP", imgdirwithpath + @"\Slide1_L.BMP");
                            pres.Export(tmpImagePathMedium, "bmp", 384, 288);
                            pres.Export(tmpImagePathSmall, "bmp", 192, 144);

                            //generate single .ppt for each slide, will use for deck generation, will need pipeline to map singles to slides in source

                            for (int x = 1; x <= pres.Slides.Count; x++)
                            {
                                string tmpSlideName = tmpSinglesPath + fileName.Replace(extension,"") + x + extension;
                                Console.WriteLine("PATH = " + tmpSlideName);
                                PPT.Presentation tmp = ppt.Presentations.Add(Office.MsoTriState.msoFalse);
                                int id = pres.Slides[x].SlideID;
                                pres.Slides[x].Copy();
                                tmp.Slides.Paste(1).FollowMasterBackground = Office.MsoTriState.msoFalse;

                                PPT.SlideRange sr = tmp.Slides.Range(1);
                                sr.Design = pres.Slides[x].Master.Design;
                                sr.ColorScheme = pres.Slides[x].ColorScheme;

                                try
                                {
                                    //some presentations don't have
                                    sr.BackgroundStyle = pres.Slides[x].BackgroundStyle;
                                }
                                catch (Exception e)
                                {
                                    string donothing_removewarning = e.Message;
                                }

                                sr.DisplayMasterShapes = pres.Slides[x].DisplayMasterShapes; 

                                tmp.SaveAs(tmpSlideName, PPT.PpSaveAsFileType.ppSaveAsDefault, Office.MsoTriState.msoFalse);

                                imgdirwithpath = tmpSlideName.Replace(".pptx", "_BMP");
                                Console.WriteLine("IMAGE PATH: " + imgdirwithpath);
                                tmp.Close();

                            }

                            pres.Close();



                            if (debug)
                            {
                                Console.WriteLine("Saved.");
                            }

                        }
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine("Error. Filename: " + openSourcePres + " Message: " + e.Message + " StackTrace: " + e.StackTrace);
                    }

                //}
                ppt.Quit();

            }
            catch (Exception e)
            {
                Console.WriteLine("Error: " + e.Message + " " + e.StackTrace);
            }

            Uri serverUri = new Uri(ServerURI);

           // string sourceURI = fileName.Replace(tmpPath, ProdDir).Replace("\\", "/");
            FileInfo[] sourceFile = new FileInfo[1];
            sourceFile[0] = new FileInfo(tmpSourcePres);
           // MessageBox.Show(sourceFile[0].Name + "   " + sourceFile[0].FullName);
           // MessageBox.Show("sourceURI" + fileName.Replace(tmpPath, ProdDir).Replace("\\", "/"));
            string[] sourceURI = { tmpSourcePres.Replace(tmpPath, ProdDir).Replace("\\", "/") };


            string[] singlesFilePaths = Directory.GetFiles(tmpSinglesPath);
            string[] smallImagePath = Directory.GetFiles(tmpImagePathSmall);
            string[] medImagePath = Directory.GetFiles(tmpImagePathMedium);
            string[] lrgImagePath = Directory.GetFiles(tmpImagePathLarge);

          
            FileInfo[] singlesFiles = new FileInfo[singlesFilePaths.Length];
            for (int i = 0; i < singlesFilePaths.Length; i++)
            {
                singlesFiles[i] = new FileInfo(singlesFilePaths[i]);
            }

            string[] singlesURIS = new string[singlesFilePaths.Length];
            for (int j = 0; j < singlesFilePaths.Length; j++)
            {
                singlesURIS[j] = singlesFiles[j].FullName.Replace(tmpPath, ArchiveDir).Replace("\\", "/");
                Console.WriteLine(singlesURIS[j]);
            }


           // Console.WriteLine("Press any key too continue...");
           // Console.ReadLine();

            FileInfo[] smImgFiles = new FileInfo[smallImagePath.Length];
            for (int i = 0; i < smallImagePath.Length; i++)
            {
                smImgFiles[i] = new FileInfo(smallImagePath[i]);
            }

            string[] smImgURIS = new string[smallImagePath.Length];
            for (int j = 0; j < smallImagePath.Length; j++)
            {
                smImgURIS[j] = smImgFiles[j].FullName.Replace(tmpPath, ProdDir).Replace("\\", "/");
                Console.WriteLine(smImgURIS[j]);
            }



            FileInfo[] medImgFiles = new FileInfo[medImagePath.Length];
            for (int i = 0; i < medImagePath.Length; i++)
            {
                medImgFiles[i] = new FileInfo(medImagePath[i]);
            }

            string[] medImgURIS = new string[medImagePath.Length];
            for (int j = 0; j < medImagePath.Length; j++)
            {
                medImgURIS[j] = medImgFiles[j].FullName.Replace(tmpPath, ProdDir).Replace("\\", "/");
                Console.WriteLine(medImgURIS[j]);
            }

 
            FileInfo[] lrgImgFiles = new FileInfo[lrgImagePath.Length];
            for (int i = 0; i < lrgImagePath.Length; i++)
            {
                lrgImgFiles[i] = new FileInfo(lrgImagePath[i]);
            }

            string[] lrgImgURIS = new string[lrgImagePath.Length];
            for (int j = 0; j < lrgImagePath.Length; j++)
            {
                lrgImgURIS[j] = lrgImgFiles[j].FullName.Replace(tmpPath, ProdDir).Replace("\\", "/");
                Console.WriteLine(lrgImgURIS[j]);
            }


            
            ContentLoader cl = new ContentLoader(serverUri);
            //cl.Load(singlesFiles);
            cl.Load(sourceURI, sourceFile);
            cl.Load(singlesURIS, singlesFiles);
            cl.Load(smImgURIS, smImgFiles);
            cl.Load(medImgURIS, medImgFiles);
            cl.Load(lrgImgURIS, lrgImgFiles);
            cl = null;

            

           /* if (debug)
            {
                Console.WriteLine("Press any key too continue...");
                Console.ReadLine();
            }*/
        }

    }
}
