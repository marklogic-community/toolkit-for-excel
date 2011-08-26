using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Marklogic.Xcc;

namespace DroidPublish
{ 
    class Program
    {
        //public static Uri serverUri = new Uri("xcc://oslo:oslo@localhost:8032");
        static void Main(string[] args)
        {
            while (true)
            {
                Uri serverUri = new Uri("xcc://oslo:oslo@localhost:8032");
                //Uri serverUri = new Uri("xcc://oslo:oslo@localhost:8032");
                String galleryQueryUri = "publish-query.xqy"; //args[1];

                ModuleRunner queryRunner = new ModuleRunner(serverUri);
                //String result = runner.InvokeToSingleString(moduleUri, "\n");
                String[] result = queryRunner.InvokeToStringArray(galleryQueryUri);

                int length = result.Length;

                //if ( Int32.Parse(result) > 0)
                if (length > 0)
                {
                    Console.WriteLine("FILES FOUND: " + length);
                    for (int i = 0; i < length; i++)
                    {
                        GalleryFactory gf = new GalleryFactory();
                        Console.WriteLine("FILE NAME: " + result[i]);
                        gf.publishGallery(result[i]);
                        gf.deleteGallery(result[i]);
                    }

                }
            }
        }
    }
}
