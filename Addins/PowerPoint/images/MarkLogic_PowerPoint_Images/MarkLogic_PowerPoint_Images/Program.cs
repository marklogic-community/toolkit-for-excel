using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
//using System.IO.Compression;
//using System.Net;
//using System.Net.Mail;
//using System.Net.Sockets;
//using System.Runtime.InteropServices;


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

            if (args.Length != 1)
            {
                Console.WriteLine(args.Length);
                Console.WriteLine(args);
                foreach (string s in args)
                {
                    Console.WriteLine("arg: " + s);
                }
                Console.WriteLine("Expected Params: (tbd) ....\nPress any key too continue...");
                Console.ReadLine();
                return;
            }

            string sourceDirectory = args[0];


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
                        //Presentation pres = ppt.Presentations.Open(file.FullName, MsoTriState.msoFalse, MsoTriState.msoTrue, MsoTriState.msoFalse);
                        string imgdirwithpath = "";
                        bool extensionCheck = checkExtension(file.Name);
                        if (extensionCheck)
                        {
                            Presentation pres = ppt.Presentations.Open(file.FullName, MsoTriState.msoFalse, MsoTriState.msoTrue, MsoTriState.msoFalse);
                            imgdirwithpath = file.FullName.Replace(".pptx", "_PNG");

                            Console.WriteLine("BEFORE:   " + file.FullName + "|" + file.Directory + "|" + file.Name);
                            pres.SaveAs(imgdirwithpath, PpSaveAsFileType.ppSaveAsPNG, MsoTriState.msoFalse);
                            pres.Close();

                        }
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine("Error. Filename: " + file.FullName + " Message: " + e.Message + " StackTrace: " + e.StackTrace);
                    }

                    Console.WriteLine("AFTER :" + file.FullName + "|" + file.Directory + "|" + file.Name);
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
                    Console.WriteLine(d.FullName);
                    FileInfo[] files = d.GetFiles();
                    Console.WriteLine("files:");

                    try
                    {
                        Application ppt = new Application();

                        foreach (FileInfo file in files)
                        {
                            try
                            {
                                //Presentation pres = ppt.Presentations.Open(file.FullName, MsoTriState.msoFalse, MsoTriState.msoTrue, MsoTriState.msoFalse);
                                string imgdirwithpath = "";
                                bool extensionCheck = checkExtension(file.Name);
                                if (extensionCheck)
                                {
                                    Presentation pres = ppt.Presentations.Open(file.FullName, MsoTriState.msoFalse, MsoTriState.msoTrue, MsoTriState.msoFalse);
                                    imgdirwithpath = file.FullName.Replace(".pptx", "_PNG");

                                    Console.WriteLine("BEFORE:   " + file.FullName + "|" + file.Directory + "|" + file.Name);
                                    pres.SaveAs(imgdirwithpath, PpSaveAsFileType.ppSaveAsPNG, MsoTriState.msoFalse);
                                    pres.Close();

                                }
                            }
                            catch (Exception e)
                            {
                                Console.WriteLine("Error. Filename: " + file.FullName + " Message: " + e.Message + " StackTrace: " + e.StackTrace);
                            }

                            Console.WriteLine("AFTER :" + file.FullName + "|" + file.Directory + "|" + file.Name);
                        }
                        ppt.Quit();
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine("Error: " + e.Message + " " + e.StackTrace);
                    }
                //}
            }

            Console.WriteLine("Press any key too continue...");
            Console.ReadLine();
        }
    }
}
