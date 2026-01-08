using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;

namespace Quoc_MEP
{
    /// <summary>
    /// ExternalEvent Handler cho Change Length từ DockPanel
    /// COPY CHÍNH XÁC từ LengthChangeEventHandler trong ChangeLengthcmd.cs
    /// </summary>
    public class ChangeLengthEventHandler : IExternalEventHandler
    {
        public double PendingLengthMm { get; set; }
        private const double MM_TO_FEET = 304.8;

        public void Execute(UIApplication app)
        {
            Logger.StartOperation($"ChangeLengthEventHandler.Execute - Length: {PendingLengthMm}mm");

            var UIDoc = app.ActiveUIDocument;
            if (UIDoc == null || UIDoc.Document == null)
            {
                Logger.Error("UIDoc or Document is null");
                MessageBox.Show("Lỗi: Document không hợp lệ\nError: Invalid Document", "Lỗi / Error");
                return;
            }

            Document doc = UIDoc.Document;
            double lengthFt = PendingLengthMm / MM_TO_FEET;

            try
            {
                // Lấy tất cả pipes và ducts trong view
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
                    TaskDialog.Show("Thông báo / Notice", 
                        "Không tìm thấy ống hoặc ống gió trong view.\nNo pipes or ducts found in view.");
                    return;
                }

                // Pick point gần ống - ĐIỂM NÀY SẼ LÀ HƯỚNG DI CHUYỂN (END POINT MỚI)
                Logger.Info("Prompting user to pick point near pipe");
                XYZ pickPoint = UIDoc.Selection.PickPoint("Click vào ống tại hướng muốn kéo dài (điểm này là end point mới) / Click on pipe direction to extend (this will be the new end point)");
                Logger.Info($"Picked point: {pickPoint}");

                // Tìm ống gần nhất và connector gần nhất
                Element nearestElement = null;
                Connector nearestConnector = null;
                double minDistance = double.MaxValue;

                foreach (Element elem in allElements)
                {
                    Connector closestConn = GetNearestConnector(elem, pickPoint);
                    if (closestConn != null)
                    {
                        // Chiếu pickPoint lên cùng độ cao với connector
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
                    TaskDialog.Show("Lỗi / Error", 
                        "Không tìm thấy ống gần điểm pick.\nCould not find pipe near picked point.");
                    return;
                }

                Logger.Info($"Nearest element: {nearestElement.Name} (ID: {nearestElement.Id})");
                Logger.Info($"Nearest connector at: {nearestConnector.Origin}");

                // Lấy LocationCurve
                LocationCurve locationCurve = nearestElement.Location as LocationCurve;
                if (locationCurve == null || !(locationCurve.Curve is Line))
                {
                    Logger.Error("Element does not have valid LocationCurve or not a straight line");
                    TaskDialog.Show("Lỗi / Error",
                        "Chỉ hỗ trợ đường thẳng.\nOnly straight lines are supported.");
                    return;
                }

                Line originalLine = locationCurve.Curve as Line;
                XYZ startPoint = originalLine.GetEndPoint(0);
                XYZ endPoint = originalLine.GetEndPoint(1);
                XYZ direction = (endPoint - startPoint).Normalize();

                Logger.Info($"Original line: Start={startPoint}, End={endPoint}");

                // Xác định điểm di chuyển dựa trên connector gần pickPoint nhất
                bool moveStartPoint = startPoint.DistanceTo(nearestConnector.Origin) < endPoint.DistanceTo(nearestConnector.Origin);
                
                Logger.Info($"Move decision: moveStartPoint={moveStartPoint}");

                // Start Transaction
                Logger.Info("Starting transaction");
                using (Transaction trans = new Transaction(doc, "Change Pipe Length from Panel"))
                {
                    trans.Start();

                    try
                    {
                        XYZ newStartPoint, newEndPoint;
                        XYZ oldMovingPoint, newMovingPoint;

                        if (moveStartPoint)
                        {
                            // Di chuyển start, cố định end
                            newEndPoint = endPoint;
                            newStartPoint = newEndPoint - direction * lengthFt;
                            oldMovingPoint = startPoint;
                            newMovingPoint = newStartPoint;
                        }
                        else
                        {
                            // Di chuyển end, cố định start
                            newStartPoint = startPoint;
                            newEndPoint = newStartPoint + direction * lengthFt;
                            oldMovingPoint = endPoint;
                            newMovingPoint = newEndPoint;
                        }

                        Logger.Info($"New line: Start={newStartPoint}, End={newEndPoint}");

                        // Lấy connector tại điểm di chuyển trước khi thay đổi
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

                        // Thu thập elements kết nối với điểm di chuyển
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

                        // Thay đổi chiều dài ống
                        Line newLine = Line.CreateBound(newStartPoint, newEndPoint);
                        locationCurve.Curve = newLine;
                        Logger.Info("Pipe length changed");

                        // Di chuyển các elements kết nối
                        if (connectedElementIds.Count > 0)
                        {
                            XYZ translationVector = newMovingPoint - oldMovingPoint;
                            Logger.Info($"Moving {connectedElementIds.Count} connected elements by vector: {translationVector}");
                            ElementTransformUtils.MoveElements(doc, connectedElementIds, translationVector);
                        }

                        trans.Commit();
                        Logger.Info("Transaction committed successfully");
                        
                        TaskDialog.Show("Success", $"Changed length to {PendingLengthMm}mm successfully!");
                    }
                    catch (Exception ex)
                    {
                        Logger.Error("Error in transaction", ex);
                        trans.RollBack();
                        throw;
                    }
                }
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                Logger.Info("User cancelled pick operation");
            }
            catch (Exception ex)
            {
                Logger.Error("Unexpected error", ex);
                TaskDialog.Show("Error", $"Lỗi: {ex.Message}\nError: {ex.Message}");
            }
            finally
            {
                Logger.EndOperation("ChangeLengthEventHandler.Execute");
            }
        }

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
            return "ChangeLengthEventHandler";
        }
    }
}
