using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace Quoc_MEP
{
    [Transaction(TransactionMode.Manual)]
    public class DisconnectCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                // Chọn phần tử MEP
                Reference pickedRef = uidoc.Selection.PickObject(
                    ObjectType.Element,
                    new MEPSelectionFilter(),
                    "Select MEP element to disconnect");

                Element element = doc.GetElement(pickedRef);

                // Kiểm tra xem có phải MEP element không
                if (!IsMEPElement(element))
                {
                    TaskDialog.Show("Error", "Selected element is not a valid MEP element.");
                    return Result.Failed;
                }

                // Disconnect tất cả connectors
                int disconnectedCount = DisconnectAllConnectors(doc, element);

                if (disconnectedCount > 0)
                {
                    TaskDialog.Show("Success", 
                        $"Disconnected {disconnectedCount} connector(s) from the selected element.");
                    return Result.Succeeded;
                }
                else
                {
                    TaskDialog.Show("Info", "No connected connectors found on the selected element.");
                    return Result.Succeeded;
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

        private bool IsMEPElement(Element element)
        {
            return element is Pipe || 
                   element is Duct || 
                   element is CableTray ||
                   element is Conduit ||
                   element is FamilyInstance;
        }

        private int DisconnectAllConnectors(Document doc, Element element)
        {
            int count = 0;
            ConnectorManager connectorManager = null;

            // Get connector manager based on element type
            if (element is MEPCurve mepCurve)
            {
                connectorManager = mepCurve.ConnectorManager;
            }
            else if (element is FamilyInstance familyInstance)
            {
                connectorManager = familyInstance.MEPModel?.ConnectorManager;
            }

            if (connectorManager == null)
                return 0;

            using (Transaction trans = new Transaction(doc, "Disconnect MEP Elements"))
            {
                trans.Start();

                foreach (Connector connector in connectorManager.Connectors)
                {
                    if (connector.IsConnected)
                    {
                        // Get all connected connectors
                        ConnectorSet connectedSet = connector.AllRefs;
                        List<Connector> connectorsToDisconnect = new List<Connector>();

                        foreach (Connector connectedConnector in connectedSet)
                        {
                            if (connectedConnector.Owner.Id != element.Id)
                            {
                                connectorsToDisconnect.Add(connectedConnector);
                            }
                        }

                        // Disconnect each connector
                        foreach (Connector connectedConnector in connectorsToDisconnect)
                        {
                            try
                            {
                                connector.DisconnectFrom(connectedConnector);
                                count++;
                            }
                            catch
                            {
                                // Continue if disconnect fails
                            }
                        }
                    }
                }

                trans.Commit();
            }

            return count;
        }
    }
}
