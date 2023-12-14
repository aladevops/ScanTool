using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Appium;
using System;
using System.Threading;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Collections.ObjectModel;
using OpenQA.Selenium.Interactions;
using System.Collections.Generic;
using System.IO;

namespace ScanTool
{
    [TestClass]
    public class ScanTool
    {
        System.Diagnostics.Process winAppDriverProcess;
        AppiumOptions options;
        WindowsDriver<WindowsElement> Scant;

        [TestInitialize] 
        public void Initialize()
        {
            // Initiate WinAppDriver
            winAppDriverProcess = System.Diagnostics.Process.Start(@"C:\\Program Files (x86)\\Windows Application Driver\\WinAppDriver.exe");

            // Launch Scan Tool 
            options = new AppiumOptions();
            options.AddAdditionalCapability("app", @"C:\gmTestBed\ScanToolAla\proj_vbn\deploy\ScanToolUI\bin\ScanToolUI.exe");

            Scant = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), options);
        }

        [TestMethod]
        public void Help()
        {
            //Help
            Scant.FindElementByName("Help").Click();
            Scant.FindElementByName("OK").Click();

            Scant.Close();
        }

        [TestMethod]
        public void Selectafile()
        {
            //Select a file 
            Scant.FindElementByName("File").Click();
            Scant.FindElementByName("Open").Click();
            Thread.Sleep(1000);
            //Scant.FindElementByName("Name").Click(); // Select a file from open dialog 
            //Scant.FindElementByAccessibilityId("System.ItemNameDisplay").Click() ;
            Scant.FindElementByName("RemoveUnused.Scantool.xml").Click(); // Select a file from open dialog
            Scant.FindElementByAccessibilityId("1").Click(); // Select Open button 
            Thread.Sleep(2000);
            Scant.FindElementByName("Cancel").Click();// As one more open dialog window display, click on cancel
            Thread.Sleep(2000);
            
            var fname = Scant.FindElementByAccessibilityId("txtStatus").Text; // Fetch the filename from display panel
            Console.WriteLine(fname);
            Thread.Sleep(1000);
            Scant.FindElementByName("Create Report").Click();
            var rdisplay = Scant.FindElementByAccessibilityId("txtStatus").Text;
            Console.WriteLine(rdisplay);
            Thread.Sleep(1000);



            Scant.Close();

        }

        [TestMethod]
        public void ReportModeElements()
        {
            // Verify the list of elements in Report Mode
            Scant.FindElementByName("Reporting Mode").Click();
            var rmode = Scant.FindElementByName("Reporting Mode").Text;
            Console.WriteLine(rmode);

            // Find the frame containing the radio buttons
            //WindowsElement frame = Scant.FindElementByName("Reporting Mode");

            // Find all radio buttons within the Reprting mode frame
            ReadOnlyCollection<WindowsElement> radioButtons = Scant.FindElementsByClassName("WindowsForms10.BUTTON.app.0.141b42a_r8_ad1");

            // Output radio button names to the console
            Console.WriteLine("Reporting Mode Options: ");
            foreach (WindowsElement radioButton in radioButtons)
            {
                
               // Console.WriteLine("Reporting Mode Options: " + radioButton.GetAttribute("Name"));
                Console.WriteLine("" + radioButton.GetAttribute("Name"));
            }

            Scant.Close();

        }

        [TestMethod]
        public void CreateaReport()
        {

            var SearchRoot = "C:\\gmTestBed\\ScanToolAla\\proj_vbn\\deploy\\ScanToolUI";
            // Scant.FindElementByAccessibilityId("txtRootPath").Text = "1";
            WindowsElement rootpath = Scant.FindElementByAccessibilityId("txtRootPath");
            // Set the text of the element using SendKeys
            rootpath.Click(); // Optional, to ensure the element has focus before sending keys
            rootpath.Clear();
            rootpath.SendKeys(SearchRoot);
            Thread.Sleep(1000);

            //VBPCount report
            Scant.FindElementByName("VBPCount").Click();
            Scant.FindElementByName("Create Report").Click();
            var VBreport = Scant.FindElementByAccessibilityId("txtStatus").Text;
            Console.WriteLine(VBreport);
            
            var SaveTo = Scant.FindElementByAccessibilityId("txtSaveTo").Text; // Find the text in Save To field
            var SaveTofile = "VBPCount.tab"; // Save to 
            
            // Verify saved to correct file
            if (SaveTo == SaveTofile)
            {
                Console.WriteLine("Report saved to the file : {0}", SaveTo);
            }
            else
            {
                Console.WriteLine("Report is not saved: {0}" + SaveTofile , SaveTo);
            }
            Thread.Sleep(2000);

            //Verfiy wether the outfile created 
            string folderPath = @"C:\gmTestBed\ScanToolAla\proj_vbn\deploy\ScanToolUI\bin";
            string filePath = Path.Combine(folderPath, SaveTofile);
            if (File.Exists(filePath))
            {
                Console.WriteLine("File exists: {0} ", SaveTofile);
            }
            else
            {
                Console.WriteLine("File does not exist: {0} ", SaveTofile);
            }


            //VBPMod report
            Scant.FindElementByName("VBPMod").Click();
            Scant.FindElementByName("Create Report").Click();
            var VBpmodreport = Scant.FindElementByAccessibilityId("txtStatus").Text;
            Console.WriteLine(VBpmodreport);

            var vbpmodSaveTo = Scant.FindElementByAccessibilityId("txtSaveTo").Text; // Find the text in Save To field
            var vbpmodSaveTofile = "VBPMod.tab"; // Save to 
                                                     // Verify saved to correct file
                if (vbpmodSaveTo == vbpmodSaveTofile)
                {
                    Console.WriteLine("Report saved to the file : {0}", vbpmodSaveTo);
                }
                else
                {
                    Console.WriteLine("Report is not saved: {0}" + vbpmodSaveTofile, vbpmodSaveTo);
                }
                Thread.Sleep(1000);     


            //COMGuid report
            Scant.FindElementByName("COMGuid").Click();
            Scant.FindElementByName("Create Report").Click();
            var COMGuidreport = Scant.FindElementByAccessibilityId("txtStatus").Text;
            Console.WriteLine(COMGuidreport);

            var COMGuidSaveTo = Scant.FindElementByAccessibilityId("txtSaveTo").Text; // Find the text in Save To field
            var COMGuidSaveTofile = "COMGuid.tab"; // Save to 
                                                       // Verify saved to correct file
                if (COMGuidSaveTo == COMGuidSaveTofile)
                {
                    Console.WriteLine("Report saved to the file : {0}", COMGuidSaveTo);
                }
                else
                {
                    Console.WriteLine("Report is not saved: {0}" + COMGuidSaveTofile, COMGuidSaveTo);
                }
                Thread.Sleep(1000);



              // All type of reports
             string[] reportTypes = { "DirList", "VBPMod", "COMGuid", "VBPBin", "VBPRef", "FolderCnt", "VBPSource", "VBPCount" };

              foreach (string reportType in reportTypes)

              {
                    // Find the frame containing the list and interact with it to select the specific type
                    // WindowsElement frame = Scant.FindElementsByClassName("WindowsForms10.BUTTON.app.0.141b42a_r8_ad1");

                  WindowsElement listItem = Scant.FindElementByName(reportType);
                  listItem.Click();

                  Thread.Sleep(1000);

                  Scant.FindElementByName("Create Report").Click();
                  Thread.Sleep(2000);
                  var report = Scant.FindElementByAccessibilityId("txtStatus").Text;
                  Console.WriteLine(report);

                  var saveTo = Scant.FindElementByAccessibilityId("txtSaveTo").Text;

                  var saveToFile = $"{reportType}.tab"; // file naming convention to save

                    if (saveTo == saveToFile)
                    {
                        Console.WriteLine("Report saved to the file: {0}", saveTo);
                    }
                    else
                    {
                        Console.WriteLine("Report is not saved: {0}", saveTo);
                    }

                    Thread.Sleep(1000);

                //Verfiy wether the output file created 
                string folderPath1 = @"C:\gmTestBed\ScanToolAla\proj_vbn\deploy\ScanToolUI\bin";
                string filePath1 = Path.Combine(folderPath1, saveToFile);
                   if (File.Exists(filePath1))
                   {
                       Console.WriteLine("File exists: {0} ", saveToFile);
                   }
                   else
                   {
                       Console.WriteLine("File does not exist: {0} ", saveToFile);
                   }

              }

                Scant.Close();
            
        }

         [TestMethod]
        public void SelectthePathCreateReport()
        {

           // Select a path from C Drive 
            Scant.FindElementByAccessibilityId("txtRootPath").Clear();

            Scant.FindElementByAccessibilityId("Drive1").Click();
            Scant.FindElementByName("c:").Click();
            Thread.Sleep(1000);
            Scant.FindElementByAccessibilityId("Drive1").Click();
            Scant.FindElementByName("d:").Click();
            Thread.Sleep(1000);
            Scant.FindElementByAccessibilityId("Drive1").Click();
            Scant.FindElementByName("c:").Click();
            Thread.Sleep(1000);
            //Focus
            Scant.FindElementByAccessibilityId("Dir1").Click();

            
            var FolderNames = Scant.FindElementByAccessibilityId("Dir1"); // Locate the folders in the panel
            var FindFolderNames = FolderNames.FindElementsByTagName("ListItem"); // Find all folders

            List<string> FolderList = new List<string>(); // Create a list to store folder names 

            foreach (var fileElement in FindFolderNames) // Extract folder names and add them to the list
            {
                string Fname = fileElement.GetAttribute("Name");
                FolderList.Add(Fname);
                if(Fname == "gmTestBed")
                {
                   Scant.FindElementByName(Fname).Click();
                    Thread.Sleep(3000);
                }
            }

            Console.WriteLine("Folder Names:");// Print the list of folder names 
            foreach (var Fname in FolderList)
            {
                Console.WriteLine(Fname);
            }

            //VBPCount report
            Scant.FindElementByName("VBPCount").Click();
            Scant.FindElementByName("Create Report").Click();
            Thread.Sleep(1000);
            var VBreport1 = Scant.FindElementByAccessibilityId("txtStatus").Text;
            Console.WriteLine(VBreport1);

            Thread.Sleep(1000);
           

            Scant.Close();

        }

        [TestMethod]
        public void ScanFolderSaveReport()
        {
            // Search the folder and create a report ( Save as )

            var SearchRoot = "C:\\gmTestBed\\ScanToolAla\\proj_vbn\\deploy\\ScanToolUI";
            WindowsElement rootpath = Scant.FindElementByAccessibilityId("txtRootPath");
            // Set the text of the element using SendKeys
            rootpath.Click(); // Optional, to ensure the element has focus before sending keys
            rootpath.Clear();
            rootpath.SendKeys(SearchRoot);
            Thread.Sleep(1000);

            Scant.FindElementByName("VBPMod").Click();
            Scant.FindElementByName("Create Report").Click();
            Scant.FindElementByName("File").Click();
            Scant.FindElementByName("Save As").Click();
            Thread.Sleep(1000);

            //var FileName = Scant.FindElementByName("File name:");
            var FileName = Scant.FindElementByAccessibilityId("1001"); // File Name field in Save as dialog
            FileName.Click();
            FileName.SendKeys("VBPMod");
            Thread.Sleep(1000);
            Scant.FindElementByName("Save").Click();
            // If file name already exists then replace with new file
            Scant.FindElementByName("Yes").Click();

            Scant.Close();



        }
    }
}
