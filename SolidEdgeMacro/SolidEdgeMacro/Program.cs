using SolidEdgeCommunity;
using SolidEdgeFramework;
using SolidEdgeFrameworkSupport;
using SolidEdgePart;
using System;

namespace SolidEdgeMacro
{
    internal class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            try
            {
                OleMessageFilter.Register();
                Application application = SolidEdgeUtils.Connect(true);
                if (application.ActiveDocumentType != DocumentTypeConstants.igPartDocument) throw new Exception("Current file is not .par!!!");

                var activeDocument = (PartDocument)application.ActiveDocument;
                var dimensions = (Dimensions)activeDocument.ProfileSets.Item(2).Profiles.Item(1).Dimensions;

                var firstDim = dimensions.Item(1);
                var secondDim = dimensions.Item(2);

                firstDim.Value = 0.07;
                secondDim.Value = 0.03;

                firstDim.VariableTableName = "Dimension1";
                secondDim.VariableTableName = "Dimension2";
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
            finally
            {
                Console.WriteLine("Completed!!!");
                OleMessageFilter.Unregister();
                Console.ReadLine();
            }
        }
    }
}
