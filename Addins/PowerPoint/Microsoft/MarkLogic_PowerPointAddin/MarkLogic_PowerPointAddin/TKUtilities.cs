/*Copyright 2009-2010 Mark Logic Corporation

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
 * TKUtilities.cs - utility functions used by UserControl1.cs.
 * 
*/

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using PwrPt = Microsoft.Office.Interop.PowerPoint;
using System.Runtime.InteropServices;
using System.Security.Permissions;
using System.IO;
//using DocumentFormat.OpenXml.Packaging; //OpenXML sdk
using Office = Microsoft.Office.Core;
using Microsoft.Win32;
using PPT = Microsoft.Office.Interop.PowerPoint;
//using OX = DocumentFormat.OpenXml.Packaging;
using System.Web.Script.Serialization;

namespace MarkLogic_PowerPointAddin
{
    public class TKUtilities
    {
        static TKUtilities()
        {
        }

        public static string convertFilenameToImageDir(string filename)
        {
            //MessageBox.Show("filename: " + filename);
            string imgDir = "";
            string tmpDir = "";
            string fname = "";

            string[] split = filename.Split(new Char[] { '\\' });
            fname = split.Last();
            //MessageBox.Show("fname: " + fname);
            tmpDir = filename.Replace(fname, "");

            fname = fname.Replace(".pptx", "_PNG");
            //MessageBox.Show("imgdir: " + fname);
            imgDir = fname; //getTempPath() + fname;
            return imgDir;

        }

        public static Office.MsoTextOrientation getTextOrientation(string textOrientation)
        {
            Office.MsoTextOrientation addOrientation;

            if (textOrientation.Equals("msoTextOrientationDownward"))
            {
                addOrientation = Office.MsoTextOrientation.msoTextOrientationDownward;
            }
            else if (textOrientation.Equals("msoTextOrientationHorizontal"))
            {
                addOrientation = Office.MsoTextOrientation.msoTextOrientationHorizontal;
            }
            else if (textOrientation.Equals("msoTextOrientationHorizontalRotatedFarEast"))
            {
                addOrientation = Office.MsoTextOrientation.msoTextOrientationHorizontalRotatedFarEast;
            }
            else if (textOrientation.Equals("msoTextOrientationMixed"))
            {
                addOrientation = Office.MsoTextOrientation.msoTextOrientationMixed;
            }
            else if (textOrientation.Equals("msoTextOrientationUpward"))
            {
                addOrientation = Office.MsoTextOrientation.msoTextOrientationUpward;
            }
            else if (textOrientation.Equals("msoTextOrientationVertical"))
            {
                addOrientation = Office.MsoTextOrientation.msoTextOrientationVertical;
            }
            else if (textOrientation.Equals("msoTextOrientationVerticalFarEast"))
            {
                addOrientation = Office.MsoTextOrientation.msoTextOrientationVerticalFarEast;
            }
            else
            {
                addOrientation = Office.MsoTextOrientation.msoTextOrientationDownward;
            }


            return addOrientation;

        }

        public static PPT.PpParagraphAlignment getParagraphAlignment(string paragraphAlignment)
        {
            PPT.PpParagraphAlignment addAlignment;

            if (paragraphAlignment.Equals("ppAlignCenter"))
            {
                addAlignment = PPT.PpParagraphAlignment.ppAlignCenter;
            }
            else if (paragraphAlignment.Equals("ppAlignDistribute"))
            {
                addAlignment = PPT.PpParagraphAlignment.ppAlignDistribute;
            }
            else if (paragraphAlignment.Equals("ppAlignJustify"))
            {
                addAlignment = PPT.PpParagraphAlignment.ppAlignJustify;
            }
            else if (paragraphAlignment.Equals("ppAlignJustifyLow"))
            {
                addAlignment = PPT.PpParagraphAlignment.ppAlignJustifyLow;
            }
            else if (paragraphAlignment.Equals("ppAlignLeft"))
            {
                addAlignment = PPT.PpParagraphAlignment.ppAlignLeft;
            }
            else if (paragraphAlignment.Equals("ppAlignMixed"))
            {
                addAlignment = PPT.PpParagraphAlignment.ppAlignmentMixed;
            }
            else if (paragraphAlignment.Equals("ppAlignRight"))
            {
                addAlignment = PPT.PpParagraphAlignment.ppAlignRight;
            }
            else if (paragraphAlignment.Equals("ppAlignThaiDistribute"))
            {
                addAlignment = PPT.PpParagraphAlignment.ppAlignThaiDistribute;
            }
            else
            {
                addAlignment = PPT.PpParagraphAlignment.ppAlignLeft;
            }


            return addAlignment;
        }

        public static PPT.PpBulletType getParagraphBulletType(string paragraphBulletType)
        {
            PPT.PpBulletType addBullet;
            if (paragraphBulletType.Equals("ppBulletMixed"))
            {
                addBullet = PPT.PpBulletType.ppBulletMixed;
            }
            else if (paragraphBulletType.Equals("ppBulletNone"))
            {
                addBullet = PPT.PpBulletType.ppBulletNone;
            }
            else if (paragraphBulletType.Equals("ppBulletNumbered"))
            {
                addBullet = PPT.PpBulletType.ppBulletNumbered;
            }
            else if (paragraphBulletType.Equals("ppBulletPicture"))
            {
                addBullet = PPT.PpBulletType.ppBulletPicture;
            }
            else if (paragraphBulletType.Equals("ppBulletUnnumbered"))
            {
                addBullet = PPT.PpBulletType.ppBulletUnnumbered;
            }
            else
            {
                addBullet = PPT.PpBulletType.ppBulletNone;
            }

            return addBullet;
        }

        public static Office.MsoPictureColorType getColorType(string colorType)
        {
            Office.MsoPictureColorType picType;

            if (colorType.Equals("msoPictureAutomatic"))
            {
                picType = Microsoft.Office.Core.MsoPictureColorType.msoPictureAutomatic;
            }
            else if (colorType.Equals("msoPictureBlackAndWhite"))
            {
                picType = Microsoft.Office.Core.MsoPictureColorType.msoPictureBlackAndWhite;
            }
            else if (colorType.Equals("msoPictureGrayscale"))
            {
                picType = Microsoft.Office.Core.MsoPictureColorType.msoPictureGrayscale;
            }
            else if (colorType.Equals("msoPictureMixed"))
            {
                picType = Microsoft.Office.Core.MsoPictureColorType.msoPictureMixed;
            }
            else if (colorType.Equals("msoPictureWatermark"))
            {
                picType = Microsoft.Office.Core.MsoPictureColorType.msoPictureWatermark;
            }
            else
            {
                picType = Microsoft.Office.Core.MsoPictureColorType.msoPictureAutomatic;
            }

            return picType;
        }

        public static Office.MsoTriState getTriState(string triState)
        {
            Office.MsoTriState state;

            if (triState.Equals("msoCTrue"))
            {
                state = Office.MsoTriState.msoCTrue;
            }
            else if (triState.Equals("msoFalse"))
            {
                state = Office.MsoTriState.msoFalse;
            }
            else if (triState.Equals("msoTriStateMixed"))
            {
                state = Office.MsoTriState.msoTriStateMixed;
            }
            else if (triState.Equals("msoTriStateToggle"))
            {
                state = Office.MsoTriState.msoTriStateToggle;
            }
            else if (triState.Equals("msoTrue"))
            {
                state = Office.MsoTriState.msoTrue;
            }
            else
            {
                state = Office.MsoTriState.msoFalse;
            }

            return state;
        }
//FINISH
        public static PPT.PpSlideLayout getSlideLayout(string customLayout)
        {
            PPT.PpSlideLayout layout;

            if (customLayout.Equals("ppLayoutBlank"))
            {
                layout = PPT.PpSlideLayout.ppLayoutBlank;
            }
            else if (customLayout.Equals("ppLayoutChart"))
            {
                layout = PPT.PpSlideLayout.ppLayoutChart;
            }
            else if (customLayout.Equals("ppLayoutChartAndText"))
            {
                layout = PPT.PpSlideLayout.ppLayoutChartAndText;
            }
            else if (customLayout.Equals("ppLayoutClipartAndText"))
            {
                layout = PPT.PpSlideLayout.ppLayoutClipartAndText;
            }
            else if (customLayout.Equals("ppLayoutClipArtAndVerticalText"))
            {
                layout = PPT.PpSlideLayout.ppLayoutClipArtAndVerticalText;
            }
            else if (customLayout.Equals("ppLayoutComparison"))
            {
                layout = PPT.PpSlideLayout.ppLayoutComparison;
            }
            else if (customLayout.Equals("ppLayoutContentWithCaption"))
            {
                layout = PPT.PpSlideLayout.ppLayoutContentWithCaption;
            }
            else if (customLayout.Equals("ppLayoutCustom"))
            {
                layout = PPT.PpSlideLayout.ppLayoutCustom;
            }
            else if (customLayout.Equals("ppLayoutFourObjects"))
            {
                layout = PPT.PpSlideLayout.ppLayoutFourObjects;
            }
            else if (customLayout.Equals("ppLayoutLargeObject"))
            {
                layout = PPT.PpSlideLayout.ppLayoutLargeObject;
            }
            else if (customLayout.Equals("ppLayoutMediaClipAndText"))
            {
                layout = PPT.PpSlideLayout.ppLayoutMediaClipAndText;
            }
            else if (customLayout.Equals("ppLayoutMixed"))
            {
                layout = PPT.PpSlideLayout.ppLayoutMixed;
            }
            else if (customLayout.Equals("ppLayoutObject"))
            {
                layout = PPT.PpSlideLayout.ppLayoutObject;
            }
            else if (customLayout.Equals("ppLayoutObjectAndText"))
            {
                layout = PPT.PpSlideLayout.ppLayoutObjectAndText;
            }
            else if (customLayout.Equals("ppLayoutObjectAndTwoObjects"))
            {
                layout = PPT.PpSlideLayout.ppLayoutObjectAndTwoObjects;
            }
            else if (customLayout.Equals("ppLayoutObjectOverText"))
            {
                layout = PPT.PpSlideLayout.ppLayoutObjectOverText;
            }
            else if (customLayout.Equals("ppLayoutOrgchart"))
            {
                layout = PPT.PpSlideLayout.ppLayoutOrgchart;
            }
            else if (customLayout.Equals("ppLayoutPictureWithCaption"))
            {
                layout = PPT.PpSlideLayout.ppLayoutPictureWithCaption;
            }
            else if (customLayout.Equals("ppLayoutSectionHeader"))
            {
                layout = PPT.PpSlideLayout.ppLayoutSectionHeader;
            }
            else if (customLayout.Equals("ppLayoutTable"))
            {
                layout = PPT.PpSlideLayout.ppLayoutTable;
            }
            else if (customLayout.Equals("ppLayoutText"))
            {
                layout = PPT.PpSlideLayout.ppLayoutText;
            }
            else if (customLayout.Equals("ppLayoutTextAndChart"))
            {
                layout = PPT.PpSlideLayout.ppLayoutTextAndChart;
            }
            else if (customLayout.Equals("ppLayoutTextAndClipart"))
            {
                layout = PPT.PpSlideLayout.ppLayoutTextAndClipart;
            }
            else if (customLayout.Equals("ppLayoutTextAndMediaClip"))
            {
                layout = PPT.PpSlideLayout.ppLayoutTextAndMediaClip;
            }
            else if (customLayout.Equals("ppLayoutTextAndObject"))
            {
                layout = PPT.PpSlideLayout.ppLayoutTextAndObject;
            }
            else if (customLayout.Equals("ppLayoutTextAndTwoObjects"))
            {
                layout = PPT.PpSlideLayout.ppLayoutTextAndTwoObjects;
            }
            else if (customLayout.Equals("ppLayoutTextOverObject"))
            {
                layout = PPT.PpSlideLayout.ppLayoutTextOverObject;
            }
            else if (customLayout.Equals("ppLayoutTitle"))
            {
                layout = PPT.PpSlideLayout.ppLayoutTitle;
            }
            else if (customLayout.Equals("ppLayoutTitleOnly"))
            {
                layout = PPT.PpSlideLayout.ppLayoutTitleOnly;
            }
            else if (customLayout.Equals("ppLayoutTwoColumnText"))
            {
                layout = PPT.PpSlideLayout.ppLayoutTwoColumnText;
            }
            else if (customLayout.Equals("ppLayoutTwoObjects"))
            {
                layout = PPT.PpSlideLayout.ppLayoutTwoObjects;
            }
            else if (customLayout.Equals("ppLayoutTwoObjectsAndObject"))
            {
                layout = PPT.PpSlideLayout.ppLayoutTwoObjectsAndObject;
            }
            else if (customLayout.Equals("ppLayoutTwoObjectsAndText"))
            {
                layout = PPT.PpSlideLayout.ppLayoutTwoObjectsAndText;
            }
            else if (customLayout.Equals("ppLayoutTwoObjectsOverText"))
            {
                layout = PPT.PpSlideLayout.ppLayoutTwoObjectsOverText;
            }
            else if (customLayout.Equals("ppLayoutVerticalText"))
            {
                layout = PPT.PpSlideLayout.ppLayoutVerticalText;
            }
            else if (customLayout.Equals("ppLayoutVerticalTitleAndText"))
            {
                layout = PPT.PpSlideLayout.ppLayoutVerticalTitleAndText;
            }
            else if (customLayout.Equals("ppLayoutVerticalTitleAndTextOverChart"))
            {
                layout = PPT.PpSlideLayout.ppLayoutVerticalTitleAndTextOverChart;
            }
            else
            {
                layout = PPT.PpSlideLayout.ppLayoutBlank;
            }

            return layout;


        }

        public static float getFloatFromString(string number)
        {
            return (float)Convert.ToDouble(number);
        }

        public static int getInt32FromString(string number)
        {
            return Convert.ToInt32(number);
        }

        public static Image byteArrayToImage(byte[] byteArrayIn)
        {
            try
            {
                MemoryStream ms = new MemoryStream(byteArrayIn);
                Image returnImage = Image.FromStream(ms);
                return returnImage;
            }
            catch (Exception e)
            {
                throw (e);
            }
        }

        public static byte[] imageToByteArray(System.Drawing.Image imageIn)
        {
            try
            {
                MemoryStream ms = new MemoryStream();
                imageIn.Save(ms, System.Drawing.Imaging.ImageFormat.Gif);
                return ms.ToArray();
            }
            catch (Exception e)
            {
                throw (e);
            }
        }

        public static void downloadFile(string url, string sourcefile, string user, string pwd)
        {
            try
            {
                System.Net.WebClient Client = new System.Net.WebClient();
                Client.Credentials = new System.Net.NetworkCredential(user, pwd);
                Client.DownloadFile(url, sourcefile);
                Client.Dispose();
            }
            catch (Exception e)
            {
                throw (e);
            }
        }

        public static void uploadData(string url, byte[] content, string user, string pwd)
        {
            try
            {
                System.Net.WebClient Client = new System.Net.WebClient();
                Client.Headers.Add("enctype", "multipart/form-data");
                Client.Headers.Add("Content-Type", "application/octet-stream");
                Client.Credentials = new System.Net.NetworkCredential(user, pwd);

                Client.UploadData(url, "POST", content);
                Client.Dispose();
            }
            catch (Exception e)
            {
                throw (e);
            }
        }

        public static byte[] downloadData(string url, string user, string pwd)
        {
            byte[] bytearray;
            try
            {
                System.Net.WebClient Client = new System.Net.WebClient();
                Client.Credentials = new System.Net.NetworkCredential(user, pwd);
                bytearray = Client.DownloadData(url);
                Client.Dispose();
            }
            catch (Exception e)
            {
                throw (e);
            }
            return bytearray;
        }

        public static bool FileInUse(string path)
        {
            string __message = "";
            try
            {
                //Just opening the file as open/create
                using (FileStream fs = new FileStream(path, FileMode.OpenOrCreate))
                {
                    //If required we can check for read/write by using fs.CanRead or fs.CanWrite
                }
                return false;
            }
            catch (IOException ex)
            {
                //check if message is for a File IO
                __message = ex.Message.ToString();
                if (__message.Contains("The process cannot access the file"))
                    return true;
                else
                    throw;
            }
        }

    }


}