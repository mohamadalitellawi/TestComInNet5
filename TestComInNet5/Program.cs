using Autodesk.AutoCAD.Interop;
using System;

namespace TestComInNet5
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");

            AcadApplication AcadApp = default;


			try
			{
				// Getting running AutoCAD instance by Marshalling by passing Programmatic ID as a string, AutoCAD.Application is the Programmatic ID for AutoCAD.
				AcadApp = (AcadApplication)Marshal3.GetActiveObject("AutoCAD.Application");

				AcadApp.Visible = true;

				AcadApp.ActiveDocument.Utility.Prompt($"Hello World\n{DateTime.Now.ToLongDateString()}\n");
				
			}
			catch (Exception ex)
			{
                Console.WriteLine(ex); 
			}

            finally
            {
                // System.Runtime.InteropServices.Marshal.ReleaseComObject(AcadApp);
                AcadApp = null;
            }
        }
    }
}
