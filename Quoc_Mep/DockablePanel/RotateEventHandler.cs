using System;
using System.Collections.Generic;
using System.Diagnostics;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace Quoc_MEP
{
    /// <summary>
    /// ExternalEvent Handler cho Rotate từ DockPanel
    /// COPY CHÍNH XÁC từ ImprovedRotationEventHandler trong Rotate\ImprovedRotationEventHandler.cs
    /// </summary>
    public class RotateEventHandler : IExternalEventHandler
    {
        public double PendingAngleDegrees { get; set; }

        public void Execute(UIApplication uiapp)
        {
            Logger.StartOperation($"RotateEventHandler.Execute - Angle: {PendingAngleDegrees}°");

            try
            {
                UIDocument uidoc = uiapp.ActiveUIDocument;
                Document doc = uidoc.Document;

                // Check selected elements
                ICollection<ElementId> selectedIds = uidoc.Selection.GetElementIds();
                
                if (selectedIds.Count == 0)
                {
                    Logger.Warning("No selection");
                    TaskDialog.Show("Warning", "Please select elements to rotate first.\nVui lòng chọn các đối tượng cần xoay trước.");
                    return;
                }

                Logger.Info($"Selected {selectedIds.Count} elements to rotate");

                // Select rotation axis with improved method
                Line rotationAxis = SelectRotationAxisImproved(uidoc);
                if (rotationAxis == null)
                {
                    Logger.Info("User cancelled axis selection or could not determine axis");
                    return;
                }

                Logger.Info($"Rotation axis determined successfully");

                // Convert angle to radians
                double angleInRadians = PendingAngleDegrees * (Math.PI / 180.0);

                // Perform rotation
                using (Transaction trans = new Transaction(doc, $"Rotate {PendingAngleDegrees}° from Panel"))
                {
                    trans.Start();

                    // Checkout elements if workshared
                    if (doc.IsWorkshared)
                    {
                        try
                        {
                            WorksharingUtils.CheckoutElements(doc, selectedIds);
                        }
                        catch (Exception ex)
                        {
                            Logger.Warning($"Checkout warning: {ex.Message}");
                        }
                    }

                    // Rotate all selected elements at once
                    ElementTransformUtils.RotateElements(doc, selectedIds, rotationAxis, angleInRadians);

                    trans.Commit();
                }

                Logger.Info($"✅ Successfully rotated {selectedIds.Count} elements by {PendingAngleDegrees}°");

                // Save angle for next time
                AngleMemory.SaveLastAngle(PendingAngleDegrees);

                // Ask for confirmation
                bool isCorrect = ConfirmRotationDirection(selectedIds.Count, PendingAngleDegrees);

                if (!isCorrect)
                {
                    // Undo by rotating back with -2x angle (because we already rotated once)
                    using (Transaction trans = new Transaction(doc, "Undo Rotation"))
                    {
                        trans.Start();
                        ElementTransformUtils.RotateElements(doc, selectedIds, rotationAxis, -2 * angleInRadians);
                        trans.Commit();
                    }
                    
                    Logger.Info("Rotation undone by user");
                    TaskDialog.Show("Info", "Rotation was undone.\nĐã hoàn tác xoay.");
                }
                else
                {
                    TaskDialog.Show("Success", $"Rotated {selectedIds.Count} element(s) by {PendingAngleDegrees}°");
                }
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                Logger.Info("Operation cancelled by user");
            }
            catch (Exception ex)
            {
                Logger.Error("Error in rotation", ex);
                TaskDialog.Show("Error", $"Rotation failed: {ex.Message}");
            }
            finally
            {
                Logger.EndOperation("RotateEventHandler.Execute");
            }
        }

        public string GetName()
        {
            return "RotateEventHandler";
        }

        /// <summary>
        /// Improved axis selection that handles both pipes and pipe fittings with connectors
        /// </summary>
        private Line SelectRotationAxisImproved(UIDocument uidoc)
        {
            try
            {
                ISelectionFilter filter = new PipeAndFittingSelectionFilter();
                
                Reference reference = uidoc.Selection.PickObject(
                    ObjectType.Element, 
                    filter,
                    "Select a pipe or fitting as rotation axis / Chọn ống hoặc phụ kiện làm trục xoay:");
                
                if (reference == null)
                    return null;

                Element axisElement = uidoc.Document.GetElement(reference);
                Logger.Info($"Selected axis element: {axisElement.Id}, Category: {axisElement.Category?.Name}");

                // Method 1: Try to get axis from LocationCurve (for pipes)
                if (axisElement.Location is LocationCurve locationCurve)
                {
                    Curve curve = locationCurve.Curve;
                    if (curve is Line line)
                    {
                        Logger.Info("Got axis from LocationCurve (Line)");
                        return line;
                    }
                    else
                    {
                        XYZ startPoint = curve.GetEndPoint(0);
                        XYZ endPoint = curve.GetEndPoint(1);
                        Logger.Info("Got axis from LocationCurve (Curve converted to Line)");
                        return Line.CreateBound(startPoint, endPoint);
                    }
                }

                // Method 2: Try to get axis from connectors (for fittings)
                ConnectorManager connectorManager = GetConnectorManager(axisElement);
                
                if (connectorManager != null)
                {
                    List<Connector> connectors = new List<Connector>();
                    foreach (Connector conn in connectorManager.Connectors)
                    {
                        connectors.Add(conn);
                    }

                    if (connectors.Count < 2)
                    {
                        TaskDialog.Show("Warning", 
                            $"Pipe fitting only has {connectors.Count} connector(s). Need at least 2 connectors.\n" +
                            $"Phụ kiện chỉ có {connectors.Count} connector. Cần ít nhất 2 connectors.");
                        return null;
                    }

                    XYZ point1 = connectors[0].Origin;
                    XYZ point2 = connectors[1].Origin;

                    // Check if points are not the same
                    if (point1.DistanceTo(point2) < 0.001)
                    {
                        TaskDialog.Show("Warning", 
                            "The first 2 connectors are at the same position.\n" +
                            "2 connector đầu tiên ở cùng vị trí.");
                        return null;
                    }

                    Line axisLine = Line.CreateBound(point1, point2);
                    
                    if (connectors.Count > 2)
                    {
                        TaskDialog.Show("Info", 
                            "Using first 2 connectors as rotation axis.\n" +
                            "Sử dụng 2 connector đầu tiên làm trục xoay.");
                    }

                    Logger.Info($"Got axis from {connectors.Count} connectors");
                    return axisLine;
                }

                // Method 3: Fallback to vertical axis through center
                if (axisElement.Location is LocationPoint locationPoint)
                {
                    XYZ center = locationPoint.Point;
                    XYZ axisDirection = XYZ.BasisZ;
                    XYZ startPoint = center - axisDirection * 10;
                    XYZ endPoint = center + axisDirection * 10;
                    
                    TaskDialog.Show("Info", 
                        "Using vertical Z-axis through element center.\n" +
                        "Sử dụng trục Z dọc qua tâm element.");
                    
                    Logger.Info("Got axis from LocationPoint (vertical Z-axis)");
                    return Line.CreateBound(startPoint, endPoint);
                }

                TaskDialog.Show("Error", 
                    "Cannot get rotation axis from selected element.\n" +
                    "Không thể xác định trục xoay từ element đã chọn.");
                return null;
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                Logger.Info("User cancelled axis selection");
                return null;
            }
            catch (Exception ex)
            {
                Logger.Error("Error in SelectRotationAxisImproved", ex);
                TaskDialog.Show("Error", $"Error selecting axis: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Get ConnectorManager from element (works for pipes, fittings, and MEP elements)
        /// </summary>
        private ConnectorManager GetConnectorManager(Element element)
        {
            try
            {
                // Try for FamilyInstance with MEPModel
                if (element is FamilyInstance fi && fi.MEPModel != null)
                {
                    return fi.MEPModel.ConnectorManager;
                }

                // Try direct property access using reflection for other MEP elements
                var connManagerProp = element.GetType().GetProperty("ConnectorManager");
                if (connManagerProp != null)
                {
                    return connManagerProp.GetValue(element) as ConnectorManager;
                }
            }
            catch (Exception ex)
            {
                Logger.Error("Error getting ConnectorManager", ex);
            }

            return null;
        }

        private bool ConfirmRotationDirection(int elementCount, double angle)
        {
            TaskDialog dialog = new TaskDialog("Rotation Confirmation");
            dialog.MainInstruction = $"Rotated {elementCount} elements by {angle}°";
            dialog.MainContent = "Các đối tượng đã được xoay. Vị trí có đúng không?\nIs the rotation direction correct?";
            dialog.CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No;
            dialog.DefaultButton = TaskDialogResult.Yes;
            
            TaskDialogResult result = dialog.Show();
            
            Logger.Info($"User confirmation result: {result}");
            return result == TaskDialogResult.Yes;
        }

        // Selection Filter class
        public class PipeAndFittingSelectionFilter : ISelectionFilter
        {
            public bool AllowElement(Element elem)
            {
                if (elem.Category != null)
                {
                    string categoryName = elem.Category.Name;
                    return categoryName == "Pipes" || 
                           categoryName == "Pipe Fittings" || 
                           categoryName == "Ducts" || 
                           categoryName == "Duct Fittings" ||
                           categoryName == "Pipe Accessories" ||
                           categoryName == "Duct Accessories";
                }
                return false;
            }

            public bool AllowReference(Reference reference, XYZ position)
            {
                return false;
            }
        }
    }
}
