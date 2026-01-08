using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Windows;

namespace Quoc_MEP
{
    /// <summary>
    /// L?nh thay d?i chi?u d�i ?ng/?ng gi�
    /// Change pipe/duct length command
    /// Version: 1.0
    /// Author: Quoc.Nguyen
    /// Date: 19.04.2025
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    public class ChangeLengthcmd : IExternalCommand
    {
        // H?ng s? chuy?n d?i / Conversion constants
        private const double MM_TO_FEET = 304.8;
        private const double MIN_LENGTH_MM = 1;
        private const double MAX_LENGTH_MM = 100000;

        private static PipeLengthWindow _window;
        private static ExternalEvent _lengthChangeEvent;
        private static LengthChangeEventHandler _eventHandler;
        private static bool _traceListenerAdded = false;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Th�m OutputDebugString listener d? xu?t ra DebugView
            if (!_traceListenerAdded)
            {
                Trace.Listeners.Add(new DefaultTraceListener());
                _traceListenerAdded = true;
            }

            Logger.StartOperation("Execute");
            
            try
            {
                // Get UIDocument v� Document
                UIDocument uidoc = commandData.Application.ActiveUIDocument;
                Document doc = uidoc.Document;
                UIApplication uiApp = commandData.Application;

                Logger.Info($"UIApp: {uiApp != null}, UIDoc: {uidoc != null}, Doc: {doc != null}");
                
                // ===== LOG BRIDGE STATE =====
                Logger.Info("=== CHECKING PANEL DATA BRIDGE ===");
                Logger.Info($"Bridge.IsCalledFromPanel: {PanelDataBridge.IsCalledFromPanel}");
                Logger.Info($"Bridge.ChangeLengthValue: {PanelDataBridge.ChangeLengthValue}");
                Logger.Info($"Bridge.ChangeLengthValue.HasValue: {PanelDataBridge.ChangeLengthValue.HasValue}");

                // ===== CHECK: N?u g?i t? Panel =====
                if (PanelDataBridge.IsCalledFromPanel && PanelDataBridge.ChangeLengthValue.HasValue)
                {
                    Logger.Info($"? Called from Panel with length: {PanelDataBridge.ChangeLengthValue.Value} mm");
                    Logger.Info("Executing direct mode (NO FORM)...");
                    
                    // Th?c thi tr?c ti?p KH�NG hi?n th? form
                    double lengthMm = PanelDataBridge.ChangeLengthValue.Value;
                    var selectedIds = uidoc.Selection.GetElementIds();
                    
                    if (selectedIds.Count == 0)
                    {
                        TaskDialog.Show("Warning", "Please select pipe or duct elements!");
                        PanelDataBridge.Reset();
                        return Result.Cancelled;
                    }
                    
                    // Th?c thi change length
                    double lengthFeet = lengthMm / 304.8;
                    int successCount = 0;
                    int skipCount = 0;
                    
                    using (Transaction trans = new Transaction(doc, "Change Length from Panel"))
                    {
                        trans.Start();
                        
                        foreach (ElementId id in selectedIds)
                        {
                            Element elem = doc.GetElement(id);
                            
                            if (elem is Autodesk.Revit.DB.Plumbing.Pipe pipe)
                            {
                                Parameter lengthParam = pipe.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH);
                                if (lengthParam != null && !lengthParam.IsReadOnly)
                                {
                                    lengthParam.Set(lengthFeet);
                                    successCount++;
                                }
                                else skipCount++;
                            }
                            else if (elem is Autodesk.Revit.DB.Mechanical.Duct duct)
                            {
                                Parameter lengthParam = duct.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH);
                                if (lengthParam != null && !lengthParam.IsReadOnly)
                                {
                                    lengthParam.Set(lengthFeet);
                                    successCount++;
                                }
                                else skipCount++;
                            }
                            else skipCount++;
                        }
                        
                        trans.Commit();
                    }
                    
                    Logger.Info($"Panel execution: {successCount} success, {skipCount} skipped");
                    
                    string msg = $"Changed {successCount} element(s) to {lengthMm}mm";
                    if (skipCount > 0) msg += $"\n{skipCount} skipped.";
                    TaskDialog.Show("Success", msg);
                    
                    // Reset Bridge
                    PanelDataBridge.Reset();
                    
                    Logger.EndOperation("Execute");
                    return Result.Succeeded;
                }
                
                // ===== G?i t? Ribbon: Hi?n th? Form =====
                Logger.Info("? Called from Ribbon - Showing form");
                Logger.Info($"Bridge was NOT set (IsCalledFromPanel={PanelDataBridge.IsCalledFromPanel})");
                
                // Kh?i t?o window v� event handler n?u chua c� (singleton pattern)
                // Initialize window and event handler if not exists (singleton pattern)
                if (_window == null)
                {
                    Logger.Info("Creating new PipeLengthWindow and ExternalEvent");
                    _eventHandler = new LengthChangeEventHandler();
                    _lengthChangeEvent = ExternalEvent.Create(_eventHandler);
                    _window = new PipeLengthWindow(uiApp);
                    
                    // Subscribe to length change event
                    _window.LengthChangeRequested += (sender, e) =>
                    {
                        Logger.Info($"LengthChangeRequested: {e.Length} mm");
                        _eventHandler.LengthMm = e.Length;
                        _eventHandler.UIDoc = uidoc;
                        _eventHandler.ParentWindow = _window; // Pass window reference
                        _lengthChangeEvent.Raise();
                    };
                }
                else
                {
                    // Reset state khi m? l?i window
                    Logger.Info("Resetting window state for reuse");
                    _window.ResetState();
                }

                // Hi?n th? window (kh�ng block)
                // Show window (non-blocking)
                Logger.Info("Showing PipeLengthWindow");
                _window.Show();
                _window.Activate();

                Logger.EndOperation("Execute");
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                PanelDataBridge.Reset(); // �?m b?o reset khi c� l?i
                return Result.Failed;
            }
        }

        /// <summary>
        /// L?y ConnectorSet t? Pipe ho?c Duct
        /// Get ConnectorSet from Pipe or Duct
        /// </summary>
        private ConnectorSet GetConnectors(Element element)
        {
            if (element is Pipe pipe)
            {
                return pipe.ConnectorManager?.Connectors;
            }
            else if (element is Duct duct)
            {
                return duct.ConnectorManager?.Connectors;
            }
            return null;
        }

        /// <summary>
        /// T�m connector g?n nh?t v?i di?m cho tru?c
        /// Find closest connector to given point
        /// </summary>
        private Connector FindClosestConnector(ConnectorSet connectors, XYZ targetPoint)
        {
            if (connectors == null) return null;

            Connector closestConnector = null;
            double minDistance = double.MaxValue;

            foreach (Connector conn in connectors)
            {
                double distance = conn.Origin.DistanceTo(targetPoint);
                if (distance < minDistance)
                {
                    minDistance = distance;
                    closestConnector = conn;
                }
            }

            return closestConnector;
        }

        /// <summary>
        /// Thu th?p t?t c? c�c element k?t n?i v?i connector
        /// Collect all elements connected to connector
        /// </summary>
        private List<Element> CollectConnectedElements(Connector connector, HashSet<int> processedIds = null)
        {
            if (processedIds == null)
            {
                processedIds = new HashSet<int>();
            }

            List<Element> result = new List<Element>();
            int ownerId = connector.Owner.Id.IntegerValue;

            if (processedIds.Contains(ownerId))
            {
                return result;
            }

            processedIds.Add(ownerId);

            foreach (Connector refConnector in connector.AllRefs)
            {
                Element refOwner = refConnector.Owner;
                int refOwnerId = refOwner.Id.IntegerValue;

                if (refOwnerId == ownerId || processedIds.Contains(refOwnerId))
                {
                    continue;
                }

                result.Add(refOwner);
                processedIds.Add(refOwnerId);

                // N?u l� fitting, ti?p t?c thu th?p c�c element k?t n?i
                // If it's a fitting, continue collecting connected elements
                if (!(refOwner is Pipe) && !(refOwner is Duct))
                {
                    ConnectorSet refConnectors = GetConnectors(refOwner);
                    if (refConnectors != null)
                    {
                        foreach (Connector otherConn in refConnectors)
                        {
                            if (otherConn.Id != refConnector.Id)
                            {
                                result.AddRange(CollectConnectedElements(otherConn, processedIds));
                            }
                        }
                    }
                }
            }

            return result;
        }
    }

    /// <summary>
    /// External Event Handler d? th?c hi?n thay d?i chi?u d�i
    /// External Event Handler to perform length change
    /// </summary>
    public class LengthChangeEventHandler : IExternalEventHandler
    {
        public UIDocument UIDoc { get; set; }
        public double LengthMm { get; set; }
        public PipeLengthWindow ParentWindow { get; set; }
        
        /// <summary>
        /// Callback du?c g?i khi operation ho�n th�nh (success ho?c cancelled)
        /// Parameters: (bool success, string message)
        /// </summary>
        public Action<bool, string> OnCompleted { get; set; }

        private const double MM_TO_FEET = 304.8;

        public void Execute(UIApplication app)
        {
            Logger.StartOperation($"EventHandler.Execute - Length: {LengthMm}mm");

            if (UIDoc == null || UIDoc.Document == null)
            {
                Logger.Error("UIDoc or Document is null");
                MessageBox.Show("L?i: Document kh�ng h?p l?\nError: Invalid Document", "L?i / Error");
                return;
            }

            Document doc = UIDoc.Document;
            double lengthFt = LengthMm / MM_TO_FEET;

            try
            {
                // L?y t?t c? pipes v� ducts trong view
                Logger.Info("Collecting all pipes and ducts in view");
                FilteredElementCollector collector = new FilteredElementCollector(doc, doc.ActiveView.Id);
                
                List<Element> allPipes = collector.OfClass(typeof(Pipe)).ToList();
                List<Element> allDucts = collector.OfClass(typeof(Duct)).ToList();
                List<Element> allElements = new List<Element>();
                allElements.AddRange(allPipes);
                allElements.AddRange(allDucts);
                
                Logger.Info($"Found {allPipes.Count} pipes and {allDucts.Count} ducts");

                if (allElements.Count == 0)
                {
                    Logger.Warning("No pipes or ducts found in active view");
                    TaskDialog.Show("Th�ng b�o / Notice", 
                        "Kh�ng t�m th?y ?ng ho?c ?ng gi� trong view.\nNo pipes or ducts found in view.");
                    return;
                }

                // Pick point g?n ?ng - �I?M N�Y S? L� HU?NG DI CHUY?N (END POINT M?I)
                Logger.Info("Prompting user to pick point near pipe");
                XYZ pickPoint = UIDoc.Selection.PickPoint("Click v�o ?ng t?i hu?ng mu?n k�o d�i (di?m n�y l� end point m?i) / Click on pipe direction to extend (this will be the new end point)");
                Logger.Info($"Picked point: {pickPoint}");

                // T�m ?ng g?n nh?t v� connector g?n nh?t
                Element nearestElement = null;
                Connector nearestConnector = null;
                double minDistance = double.MaxValue;

                foreach (Element elem in allElements)
                {
                    Connector closestConn = GetNearestConnector(elem, pickPoint);
                    if (closestConn != null)
                    {
                        // Chi?u pickPoint l�n c�ng d? cao v?i connector
                        XYZ projectedPoint = new XYZ(pickPoint.X, pickPoint.Y, closestConn.Origin.Z);
                        double distance = projectedPoint.DistanceTo(closestConn.Origin);
                        
                        if (distance < minDistance)
                        {
                            minDistance = distance;
                            nearestElement = elem;
                            nearestConnector = closestConn;
                        }
                    }
                }

                if (nearestElement == null || nearestConnector == null)
                {
                    Logger.Error("Could not find nearest pipe/duct or connector");
                    TaskDialog.Show("L?i / Error", 
                        "Kh�ng t�m th?y ?ng g?n di?m pick.\nCould not find pipe near picked point.");
                    return;
                }

                Logger.Info($"Nearest element: {nearestElement.Name} (ID: {nearestElement.Id})");
                Logger.Info($"Nearest connector at: {nearestConnector.Origin}");

                // L?y LocationCurve
                LocationCurve locationCurve = nearestElement.Location as LocationCurve;
                if (locationCurve == null || !(locationCurve.Curve is Line))
                {
                    Logger.Error("Element does not have valid LocationCurve or not a straight line");
                    TaskDialog.Show("L?i / Error",
                        "Ch? h? tr? du?ng th?ng.\nOnly straight lines are supported.");
                    return;
                }

                Line originalLine = locationCurve.Curve as Line;
                XYZ startPoint = originalLine.GetEndPoint(0);
                XYZ endPoint = originalLine.GetEndPoint(1);
                XYZ direction = (endPoint - startPoint).Normalize();

                Logger.Info($"Original line: Start={startPoint}, End={endPoint}");

                // X�c d?nh di?m di chuy?n d?a tr�n connector g?n pickPoint nh?t
                // The connector nearest to pickPoint will be the moving direction
                bool moveStartPoint = startPoint.DistanceTo(nearestConnector.Origin) < endPoint.DistanceTo(nearestConnector.Origin);
                
                Logger.Info($"Move decision: moveStartPoint={moveStartPoint}");

                // Start Transaction
                Logger.Info("Starting transaction");
                using (Transaction trans = new Transaction(doc, "Thay d?i chi?u d�i ?ng / Change Pipe Length"))
                {
                    trans.Start();

                    try
                    {
                        XYZ newStartPoint, newEndPoint;
                        XYZ oldMovingPoint, newMovingPoint;

                        if (moveStartPoint)
                        {
                            // Di chuy?n start, c? d?nh end
                            newEndPoint = endPoint;
                            newStartPoint = newEndPoint - direction * lengthFt;
                            oldMovingPoint = startPoint;
                            newMovingPoint = newStartPoint;
                        }
                        else
                        {
                            // Di chuy?n end, c? d?nh start
                            newStartPoint = startPoint;
                            newEndPoint = newStartPoint + direction * lengthFt;
                            oldMovingPoint = endPoint;
                            newMovingPoint = newEndPoint;
                        }

                        Logger.Info($"New line: Start={newStartPoint}, End={newEndPoint}");

                        // L?y connector t?i di?m di chuy?n tru?c khi thay d?i
                        ConnectorSet connectors = null;
                        if (nearestElement is Pipe pipe)
                        {
                            connectors = pipe.ConnectorManager?.Connectors;
                        }
                        else if (nearestElement is Duct duct)
                        {
                            connectors = duct.ConnectorManager?.Connectors;
                        }

                        Connector movingConnector = null;
                        if (connectors != null)
                        {
                            foreach (Connector conn in connectors)
                            {
                                if (conn.Origin.DistanceTo(oldMovingPoint) < 0.01) // tolerance
                                {
                                    movingConnector = conn;
                                    break;
                                }
                            }
                        }

                        // Thu th?p elements k?t n?i v?i di?m di chuy?n
                        List<ElementId> connectedElementIds = new List<ElementId>();
                        if (movingConnector != null && movingConnector.IsConnected)
                        {
                            Logger.Info("Collecting connected elements");
                            foreach (Connector refConn in movingConnector.AllRefs)
                            {
                                if (refConn.Owner.Id != nearestElement.Id)
                                {
                                    connectedElementIds.Add(refConn.Owner.Id);
                                    Logger.Info($"Found connected element: {refConn.Owner.Name} (ID: {refConn.Owner.Id})");
                                }
                            }
                        }

                        // Thay d?i chi?u d�i ?ng
                        Line newLine = Line.CreateBound(newStartPoint, newEndPoint);
                        locationCurve.Curve = newLine;
                        Logger.Info("Pipe length changed");

                        // Di chuy?n c�c elements k?t n?i
                        if (connectedElementIds.Count > 0)
                        {
                            XYZ translationVector = newMovingPoint - oldMovingPoint;
                            Logger.Info($"Moving {connectedElementIds.Count} connected elements by vector: {translationVector}");
                            ElementTransformUtils.MoveElements(doc, connectedElementIds, translationVector);
                        }

                        trans.Commit();
                        Logger.Info("Transaction committed successfully");
                        
                        // Notify success
                        OnCompleted?.Invoke(true, $"? Changed length to {LengthMm}mm successfully!");


                    }
                    catch (Exception ex)
                    {
                        Logger.Error("Error in transaction", ex);
                        trans.RollBack();
                        OnCompleted?.Invoke(false, $"? Transaction failed: {ex.Message}");
                        throw;
                    }
                }
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                Logger.Info("User cancelled pick operation");
                OnCompleted?.Invoke(false, "? Operation cancelled by user");
            }
            catch (Exception ex)
            {
                Logger.Error("Unexpected error", ex);
                MessageBox.Show(
                    $"L?i: {ex.Message}\nError: {ex.Message}",
                    "L?i / Error");
                OnCompleted?.Invoke(false, $"? Error: {ex.Message}");
            }
            finally
            {
                // Reset window state v� show l?i d? c� th? s? d?ng l?i
                if (ParentWindow != null)
                {
                    ParentWindow.Dispatcher.Invoke(() =>
                    {
                        ParentWindow.ResetState();
                        ParentWindow.Show();
                        ParentWindow.Activate();
                        Logger.Info("Window reset and shown for next use");
                    });
                }
                
                Logger.EndOperation("EventHandler.Execute");
            }
        }

        /// <summary>
        /// T�m connector g?n nh?t d?n di?m pickpoint
        /// Find nearest connector to pickpoint
        /// </summary>
        private Connector GetNearestConnector(Element element, XYZ pickPoint)
        {
            ConnectorSet connectors = null;

            if (element is Pipe pipe)
            {
                connectors = pipe.ConnectorManager?.Connectors;
            }
            else if (element is Duct duct)
            {
                connectors = duct.ConnectorManager?.Connectors;
            }

            if (connectors == null || connectors.Size == 0)
            {
                return null;
            }

            Connector nearestConnector = null;
            double minDistance = double.MaxValue;

            foreach (Connector conn in connectors)
            {
                double distance = conn.Origin.DistanceTo(pickPoint);
                if (distance < minDistance)
                {
                    minDistance = distance;
                    nearestConnector = conn;
                }
            }

            return nearestConnector;
        }

        public string GetName()
        {
            return "LengthChangeEventHandler";
        }
    }

    /// <summary>
    /// Filter d? ch? cho ph�p ch?n Pipe ho?c Duct
    /// Filter to only allow selecting Pipe or Duct
    /// </summary>
    public class PipeOrDuctSelectionFilter : ISelectionFilter
    {
        public bool AllowElement(Element elem)
        {
            return elem is Pipe || elem is Duct;
        }

        public bool AllowReference(Reference reference, XYZ position)
        {
            return false;
        }
    }
}

