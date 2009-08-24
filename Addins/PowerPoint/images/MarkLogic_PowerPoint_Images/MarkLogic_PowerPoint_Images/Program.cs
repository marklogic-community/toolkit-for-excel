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
 * Program.cs : simple utility to save all powerpoints in a given directory (and its subdirectories)
 *              in associated image folder with each slide in .PNG format
 * 
*/

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
            //directory-path : path of dir with ppt(x)s to save as images
            //{debug}        : true||false

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
