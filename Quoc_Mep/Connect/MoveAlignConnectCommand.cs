using System;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace Quoc_MEP
{
    /// <summary>
    /// Move, Align and Connect MEP elements - with alignment check
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    public class MoveAlignConnectCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                // Step 1: Pick destination element
                Reference destRef = uidoc.Selection.PickObject(
                    ObjectType.Element,
                    new MEPSelectionFilter(),
                    "Select destination MEP element (will stay in place)");

                Element destElement = doc.GetElement(destRef);

                // Step 2: Pick source element to move
                Reference srcRef = uidoc.Selection.PickObject(
                    ObjectType.Element,
                    new MEPSelectionFilter(),
                    "Select source MEP element (will be moved and connected)");

                Element srcElement = doc.GetElement(srcRef);

                // Validate elements
                if (srcElement.Id == destElement.Id)
                {
                    TaskDialog.Show("Error", "Cannot connect an element to itself!");
                    return Result.Failed;
                }

                // Execute move, align and connect with alignment enforcement
                bool success = ConnectionHelper.MoveAlignAndConnect(doc, srcElement, destElement);

                if (success)
                {
                    TaskDialog.Show("Success", 
                        "Elements moved, aligned and connected successfully!");
                    return Result.Succeeded;
                }
                else
                {
                    TaskDialog.Show("Failed", 
                        "Could not move, align and connect the elements. " +
                        "Check if elements are compatible and have available connectors.");
                    return Result.Failed;
                }
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return Result.Cancelled;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }
    }
}
