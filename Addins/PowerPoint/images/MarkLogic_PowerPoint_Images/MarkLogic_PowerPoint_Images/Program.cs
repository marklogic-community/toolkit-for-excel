using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;


namespace MarkLogic_PowerPoint_Images
{
    class Program
    {
        static bool checkExtension(string ext)
        {
            if (ext.EndsWith(".pptx") ||
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
        static void Main(string[] args)
        {
            //parameters reqd
            //path of dir with ppt(x)s  convert, load to ML
            //path of output dir for image dirs ? (right now just create dir of images sibling to .pptx)
            //url to ML for saving the ppt and images 
            //user
            //pwd
            //could add- folder to save in ML, right now defaults to '/'

            //this assumes pipeline already installed 

            if (args.Length < 1 || args.Length >2 )
            {
                //Console.WriteLine(args.Length);
                //Console.WriteLine(args);
                foreach (string s in args)
                {
                    Console.WriteLine("arg: " + s);
                }
                Console.WriteLine("Expected Params: directory-path {debug} \n\n"+
                                  "directory-path   = path where .pptx can be found. Example: C:\\my-pptx \n"+
                                  "debug {optional} = true or false. if true, debug messages enabled \n                   and prompt entry required for exit\n\n"+
                                  "Press any key too continue...");
                Console.ReadLine();
                return;
            }

            string sourceDirectory = args[0];
            bool debug = false;

            if (args.Length==2 && args[1].ToLower().Equals("true")){

                debug = true;
            }


            object missing = System.Type.Missing;
            DirectoryInfo root = new DirectoryInfo(sourceDirectory);
            FileInfo[] rootFiles = root.GetFiles();
            DirectoryInfo[] dirs = root.GetDirectories("*", SearchOption.AllDirectories);


            try
            {
                Application ppt = new Application();

                foreach (FileInfo file in rootFiles)
                {
                    try
                    {
                        
                        string imgdirwithpath = "";
                        bool extensionCheck = checkExtension(file.Name);
                        if (extensionCheck)
                        {
                            Presentation pres = ppt.Presentations.Open(file.FullName, MsoTriState.msoFalse, MsoTriState.msoTrue, MsoTriState.msoFalse);
                            imgdirwithpath = file.FullName.Replace(".pptx", "_PNG");

                            if (debug)
                            {
                                Console.WriteLine("Saving images for file :   " + file.FullName );
                            }

                            pres.SaveAs(imgdirwithpath, PpSaveAsFileType.ppSaveAsPNG, MsoTriState.msoFalse);
                            pres.Close();

                            if (debug)
                            {
                                Console.WriteLine("Saved.");
                            }

                        }
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine("Error. Filename: " + file.FullName + " Message: " + e.Message + " StackTrace: " + e.StackTrace);
                    }

                }
                   ppt.Quit();

            }catch (Exception e)
            {
                        Console.WriteLine("Error: " + e.Message + " " + e.StackTrace);
            }

            
            foreach (DirectoryInfo d in dirs)
            {
                //if (d.Name.Equals("ug2009"))//only for testing purposes
                //{
                if (debug)
                {
                    Console.WriteLine(d.FullName);
                }
                    FileInfo[] files = d.GetFiles();

                    try
                    {
                        Application ppt = new Application();

                        foreach (FileInfo file in files)
                        {
                            try
                            {
                                string imgdirwithpath = "";
                                bool extensionCheck = checkExtension(file.Name);
                                if (extensionCheck)
                                {
                                    Presentation pres = ppt.Presentations.Open(file.FullName, MsoTriState.msoFalse, MsoTriState.msoTrue, MsoTriState.msoFalse);
                                    imgdirwithpath = file.FullName.Replace(".pptx", "_PNG");

                                    if (debug)
                                    {
                                        Console.WriteLine("Saving images for file:   " + file.FullName );
                                    }

                                    pres.SaveAs(imgdirwithpath, PpSaveAsFileType.ppSaveAsPNG, MsoTriState.msoFalse);
                                    pres.Close();

                                    if (debug)
                                    {
                                        Console.WriteLine("Saved.");
                                    }


                                }
                            }
                            catch (Exception e)
                            {
                                Console.WriteLine("Error. Filename: " + file.FullName + " Message: " + e.Message + " StackTrace: " + e.StackTrace);
                            }

                        }
                        ppt.Quit();
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine("Error: " + e.Message + " " + e.StackTrace);
                    }
                //}
            }

            if (debug)
            {
                Console.WriteLine("Press any key too continue...");
                Console.ReadLine();
            }
        }
    }
}
