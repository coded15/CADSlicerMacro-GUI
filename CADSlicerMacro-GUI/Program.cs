using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorks.Interop.swpublished;
using SolidWorks.Interop.swcommands;
using SolidWorks.Interop.swdocumentmgr;
using System;

namespace SolidWorksAutomation
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create an instance of SolidWorks application
            SldWorks swApp = null;

            try
            {
                // Try to connect to an existing instance of SolidWorks
                swApp = (SldWorks)Activator.CreateInstance(Type.GetTypeFromProgID("SldWorks.Application"));
            }
            catch (Exception)
            {
                // If SolidWorks is not running, create a new instance
                swApp = new SldWorks();
                swApp.Visible = true; // Make SolidWorks visible
            }

            // Now SolidWorks should be open and visible
            Console.WriteLine("SolidWorks is now open.");

            // Open a part document
            ModelDoc2 swModel = swApp.OpenDoc6(@"E:\Desktop\car seat\worm gear 2.SLDPRT", (int)swDocumentTypes_e.swDocPART, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", 0, 0);
            if (swModel != null)
            {
                Console.WriteLine("Part opened successfully!");
                // Switch to the Weldments environment
                // Run the SolidWorks macro
                RunSolidWorksMacro(swApp, @"E:\Desktop\BTP\macros\startingWeldment.swp");
            }
            else
            {
                Console.WriteLine("Failed to open part!");
            }

            // Wait for user input before exiting
            Console.WriteLine("Press any key to exit SolidWorks...");
            Console.ReadKey();

            // Don't forget to release SolidWorks instance when done
            swApp.ExitApp();
            swApp = null;
        }
        static void RunSolidWorksMacro(SldWorks swApp, string macroPath)
        {
            int error = 0; // Define an error variable
            int options = 0; // Set options to 0 for default behavior

            // Run the SolidWorks macro
            swApp.RunMacro2(macroPath, "", "", options, out error);

            // Check the value of the error variable to handle any errors
            if (error != 0)
            {
                Console.WriteLine("Error occurred while executing the macro. Error code: " + error);
            }
            else
            {
                Console.WriteLine("Macro executed successfully.");
            }
        }


    }
}
