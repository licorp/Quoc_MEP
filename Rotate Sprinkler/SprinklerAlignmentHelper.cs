using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;

namespace Quoc_MEP
{
    /// <summary>
    /// Helper class để xử lý việc căn chỉnh Sprinkler thẳng hàng với Pap
    /// Kiểm tra theo trục Z (vertical) - tức là X-Y phải khớp nhau
    /// </summary>
    public static class SprinklerAlignmentHelper
    {
        // Tolerance mặc định: 1mm = 1/304.8 feet
        private const double DEFAULT_TOLERANCE = 1.0 / 304.8;

        /// <summary>
        /// Kiểm tra element có thẳng hàng với Pap theo trục Z không
        /// (X-Y của element phải khớp với X-Y của Pap connector)
        /// </summary>
        /// <param name="papConnectorOrigin">Vị trí connector của Pap</param>
        /// <param name="elementConnectorOrigin">Vị trí connector của element</param>
        /// <param name="tolerance">Sai số chấp nhận được (feet)</param>
        /// <returns>True nếu thẳng hàng</returns>
        public static bool IsAlignedVertically(XYZ papConnectorOrigin, XYZ elementConnectorOrigin, double tolerance = DEFAULT_TOLERANCE)
        {
            // Kiểm tra X và Y có khớp nhau không (chỉ cho phép sai lệch trong tolerance)
            double deltaX = Math.Abs(papConnectorOrigin.X - elementConnectorOrigin.X);
            double deltaY = Math.Abs(papConnectorOrigin.Y - elementConnectorOrigin.Y);

            return deltaX <= tolerance && deltaY <= tolerance;
        }

        /// <summary>
        /// Lấy tất cả element kết nối với Pap (bao gồm Pipe và Sprinkler)
        /// </summary>
        /// <param name="pap">Element Pap (Pipe Accessory)</param>
        /// <returns>Danh sách các element kết nối</returns>
        public static List<Element> GetConnectedElements(Element pap)
        {
            var connectedElements = new List<Element>();

            try
            {
                ConnectorManager connectorManager = GetConnectorManager(pap);
                if (connectorManager == null) return connectedElements;

                foreach (Connector connector in connectorManager.Connectors)
                {
                    if (connector.IsConnected)
                    {
                        foreach (Connector connectedConnector in connector.AllRefs)
                        {
                            Element connectedElement = connectedConnector.Owner;
                            if (connectedElement != null && connectedElement.Id != pap.Id)
                            {
                                // Tránh thêm trùng lặp
                                if (!connectedElements.Any(e => e.Id == connectedElement.Id))
                                {
                                    connectedElements.Add(connectedElement);
                                }
                            }
                        }
                    }
                }
            }
            catch
            {
                // Silent fail
            }

            return connectedElements;
        }

        /// <summary>
        /// Lấy tất cả element trong chuỗi kết nối từ Pap (đệ quy)
        /// </summary>
        /// <param name="pap">Element Pap gốc</param>
        /// <returns>Danh sách tất cả element trong chuỗi (không bao gồm Pap)</returns>
        public static List<Element> GetAllConnectedElementsChain(Element pap)
        {
            var allElements = new List<Element>();
            var visitedIds = new HashSet<ElementId>();
            visitedIds.Add(pap.Id);

            GetConnectedElementsRecursive(pap, allElements, visitedIds);

            return allElements;
        }

        private static void GetConnectedElementsRecursive(Element element, List<Element> result, HashSet<ElementId> visitedIds)
        {
            var directConnected = GetConnectedElements(element);

            foreach (var connected in directConnected)
            {
                if (!visitedIds.Contains(connected.Id))
                {
                    visitedIds.Add(connected.Id);
                    result.Add(connected);
                    GetConnectedElementsRecursive(connected, result, visitedIds);
                }
            }
        }

        /// <summary>
        /// Lấy connector kết nối giữa hai element
        /// </summary>
        public static (Connector element1Connector, Connector element2Connector) GetConnectingConnectors(Element element1, Element element2)
        {
            try
            {
                ConnectorManager cm1 = GetConnectorManager(element1);
                if (cm1 == null) return (null, null);

                foreach (Connector c1 in cm1.Connectors)
                {
                    if (c1.IsConnected)
                    {
                        foreach (Connector connectedC in c1.AllRefs)
                        {
                            if (connectedC.Owner.Id == element2.Id)
                            {
                                return (c1, connectedC);
                            }
                        }
                    }
                }
            }
            catch
            {
                // Silent fail
            }

            return (null, null);
        }

        /// <summary>
        /// Tìm connector của Pap hướng xuống (theo trục Z âm) - connector kết nối với ống/sprinkler phía dưới
        /// </summary>
        public static Connector GetDownwardConnector(Element pap)
        {
            try
            {
                ConnectorManager connectorManager = GetConnectorManager(pap);
                if (connectorManager == null) return null;

                Connector downwardConnector = null;
                double lowestZ = double.MaxValue;

                foreach (Connector connector in connectorManager.Connectors)
                {
                    // Tìm connector có Z thấp nhất (hướng xuống)
                    if (connector.Origin.Z < lowestZ)
                    {
                        lowestZ = connector.Origin.Z;
                        downwardConnector = connector;
                    }
                }

                return downwardConnector;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Căn chỉnh element theo trục Z để thẳng hàng với Pap
        /// Di chuyển element sao cho X-Y của nó khớp với X-Y của Pap connector
        /// </summary>
        /// <param name="doc">Document</param>
        /// <param name="papConnector">Connector của Pap (làm chuẩn)</param>
        /// <param name="element">Element cần căn chỉnh</param>
        /// <param name="elementConnector">Connector của element gần Pap nhất</param>
        /// <returns>True nếu căn chỉnh thành công</returns>
        public static bool AlignElementToPap(Document doc, Connector papConnector, Element element, Connector elementConnector)
        {
            try
            {
                if (papConnector == null || elementConnector == null) return false;

                // Tính vector di chuyển chỉ theo X-Y (giữ nguyên Z)
                double deltaX = papConnector.Origin.X - elementConnector.Origin.X;
                double deltaY = papConnector.Origin.Y - elementConnector.Origin.Y;

                // Nếu đã thẳng hàng thì không cần di chuyển
                if (Math.Abs(deltaX) <= DEFAULT_TOLERANCE && Math.Abs(deltaY) <= DEFAULT_TOLERANCE)
                {
                    return true;
                }

                XYZ moveVector = new XYZ(deltaX, deltaY, 0);

                // Unpin element nếu cần
                ConnectionHelper.UnpinElementIfPinned(doc, element);

                // Di chuyển element
                ElementTransformUtils.MoveElement(doc, element.Id, moveVector);

                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Xử lý toàn bộ quá trình căn chỉnh cho một element
        /// 1. Disconnect
        /// 2. Di chuyển để thẳng hàng
        /// 3. Connect lại
        /// </summary>
        /// <param name="doc">Document</param>
        /// <param name="pap">Pap element (cố định)</param>
        /// <param name="element">Element cần căn chỉnh</param>
        /// <returns>True nếu thành công</returns>
        public static bool DisconnectAlignReconnect(Document doc, Element pap, Element element)
        {
            try
            {
                // Lấy connector kết nối giữa Pap và element
                var (papConnector, elementConnector) = GetConnectingConnectors(pap, element);
                
                if (papConnector == null || elementConnector == null)
                {
                    // Không có kết nối trực tiếp, bỏ qua
                    return false;
                }

                // Kiểm tra đã thẳng hàng chưa
                if (IsAlignedVertically(papConnector.Origin, elementConnector.Origin))
                {
                    // Đã thẳng hàng, không cần xử lý
                    return true;
                }

                // 1. Disconnect
                ConnectionHelper.UnpinElementIfPinned(doc, element);
                bool disconnected = ConnectionHelper.DisconnectTwoElements(doc, pap, element);
                if (!disconnected)
                {
                    return false;
                }

                // 2. Tính vector di chuyển
                double deltaX = papConnector.Origin.X - elementConnector.Origin.X;
                double deltaY = papConnector.Origin.Y - elementConnector.Origin.Y;
                XYZ moveVector = new XYZ(deltaX, deltaY, 0);

                // Di chuyển element
                ElementTransformUtils.MoveElement(doc, element.Id, moveVector);

                // 3. Connect lại
                // Lấy lại connector sau khi di chuyển
                var elementConnectorManager = GetConnectorManager(element);
                if (elementConnectorManager == null) return false;

                Connector newElementConnector = null;
                double minDistance = double.MaxValue;

                foreach (Connector c in elementConnectorManager.Connectors)
                {
                    if (!c.IsConnected)
                    {
                        double dist = c.Origin.DistanceTo(papConnector.Origin);
                        if (dist < minDistance)
                        {
                            minDistance = dist;
                            newElementConnector = c;
                        }
                    }
                }

                if (newElementConnector != null)
                {
                    // Move to exact position
                    XYZ finalMove = papConnector.Origin - newElementConnector.Origin;
                    ElementTransformUtils.MoveElement(doc, element.Id, finalMove);

                    // Get updated connector
                    elementConnectorManager = GetConnectorManager(element);
                    foreach (Connector c in elementConnectorManager.Connectors)
                    {
                        if (!c.IsConnected && c.Origin.DistanceTo(papConnector.Origin) < 0.001)
                        {
                            c.ConnectTo(papConnector);
                            return true;
                        }
                    }
                }

                return false;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Xử lý căn chỉnh cho chuỗi element (Pipe -> Sprinkler)
        /// Element gần Pap nhất sẽ được căn chỉnh trước, sau đó đến các element tiếp theo
        /// </summary>
        public static AlignmentResult AlignConnectedChain(Document doc, Element pap)
        {
            var result = new AlignmentResult();

            try
            {
                // Lấy connector hướng xuống của Pap
                Connector papDownConnector = GetDownwardConnector(pap);
                if (papDownConnector == null)
                {
                    result.ErrorMessage = "Không tìm thấy connector của Pap";
                    return result;
                }

                // Lấy element kết nối trực tiếp với Pap qua connector này
                Element directConnected = null;
                Connector directConnectedConnector = null;

                if (papDownConnector.IsConnected)
                {
                    foreach (Connector connectedC in papDownConnector.AllRefs)
                    {
                        if (connectedC.Owner.Id != pap.Id)
                        {
                            directConnected = connectedC.Owner;
                            directConnectedConnector = connectedC;
                            break;
                        }
                    }
                }

                if (directConnected == null)
                {
                    result.ErrorMessage = "Không tìm thấy element kết nối với Pap";
                    return result;
                }

                // Kiểm tra và căn chỉnh element kết nối trực tiếp
                if (!IsAlignedVertically(papDownConnector.Origin, directConnectedConnector.Origin))
                {
                    result.ElementsNeedAlignment.Add(directConnected);

                    // Disconnect
                    ConnectionHelper.UnpinElementIfPinned(doc, directConnected);
                    ConnectionHelper.DisconnectTwoElements(doc, pap, directConnected);

                    // Tính vector di chuyển
                    double deltaX = papDownConnector.Origin.X - directConnectedConnector.Origin.X;
                    double deltaY = papDownConnector.Origin.Y - directConnectedConnector.Origin.Y;
                    XYZ moveVector = new XYZ(deltaX, deltaY, 0);

                    // Di chuyển element (và tất cả element con kết nối với nó)
                    var chainElements = GetAllConnectedElementsChain(directConnected);
                    
                    // Di chuyển element chính
                    ElementTransformUtils.MoveElement(doc, directConnected.Id, moveVector);
                    result.ElementsAligned.Add(directConnected);

                    // Di chuyển các element trong chuỗi
                    foreach (var chainElement in chainElements)
                    {
                        ConnectionHelper.UnpinElementIfPinned(doc, chainElement);
                        ElementTransformUtils.MoveElement(doc, chainElement.Id, moveVector);
                        result.ElementsAligned.Add(chainElement);
                    }

                    // Kết nối lại với Pap
                    var updatedConnectorManager = GetConnectorManager(directConnected);
                    if (updatedConnectorManager != null)
                    {
                        foreach (Connector c in updatedConnectorManager.Connectors)
                        {
                            if (!c.IsConnected)
                            {
                                double dist = c.Origin.DistanceTo(papDownConnector.Origin);
                                if (dist < 0.01) // Gần đủ để kết nối
                                {
                                    // Move nhỏ để khớp chính xác
                                    XYZ finalMove = papDownConnector.Origin - c.Origin;
                                    ElementTransformUtils.MoveElement(doc, directConnected.Id, finalMove);
                                    
                                    // Di chuyển chain elements theo
                                    foreach (var chainElement in chainElements)
                                    {
                                        ElementTransformUtils.MoveElement(doc, chainElement.Id, finalMove);
                                    }

                                    // Kết nối
                                    updatedConnectorManager = GetConnectorManager(directConnected);
                                    foreach (Connector c2 in updatedConnectorManager.Connectors)
                                    {
                                        if (!c2.IsConnected && c2.Origin.DistanceTo(papDownConnector.Origin) < 0.001)
                                        {
                                            c2.ConnectTo(papDownConnector);
                                            result.Success = true;
                                            break;
                                        }
                                    }
                                    break;
                                }
                            }
                        }
                    }
                }
                else
                {
                    result.Success = true;
                    result.AlreadyAligned = true;
                }
            }
            catch (Exception ex)
            {
                result.ErrorMessage = ex.Message;
            }

            return result;
        }

        /// <summary>
        /// Kiểm tra vector từ Pap đến Sprinkler có song song với trục Z không
        /// </summary>
        /// <param name="papPosition">Vị trí Pap (điểm đầu vector)</param>
        /// <param name="sprinklerPosition">Vị trí Sprinkler (điểm cuối vector)</param>
        /// <param name="tolerance">Góc cho phép (radians)</param>
        /// <returns>True nếu song song với Z</returns>
        public static bool IsParallelToZAxis(XYZ papPosition, XYZ sprinklerPosition, double tolerance = 0.001)
        {
            XYZ vector = sprinklerPosition - papPosition;
            if (vector.IsZeroLength()) return true;

            XYZ normalizedVector = vector.Normalize();
            XYZ zAxis = XYZ.BasisZ; // (0, 0, 1)
            XYZ negativeZAxis = -XYZ.BasisZ; // (0, 0, -1)

            // Kiểm tra song song với Z+ hoặc Z-
            double dotProductPositive = Math.Abs(normalizedVector.DotProduct(zAxis));
            double dotProductNegative = Math.Abs(normalizedVector.DotProduct(negativeZAxis));

            // Nếu dot product gần 1 thì song song
            return dotProductPositive >= (1.0 - tolerance) || dotProductNegative >= (1.0 - tolerance);
        }

        /// <summary>
        /// Tính góc giữa vector Pap-Sprinkler và trục Z
        /// </summary>
        /// <param name="papPosition">Vị trí Pap (điểm đầu)</param>
        /// <param name="sprinklerPosition">Vị trí Sprinkler (điểm cuối)</param>
        /// <returns>Góc tính bằng radians</returns>
        public static double GetAngleWithZAxis(XYZ papPosition, XYZ sprinklerPosition)
        {
            XYZ vector = sprinklerPosition - papPosition;
            if (vector.IsZeroLength()) return 0;

            XYZ normalizedVector = vector.Normalize();
            
            // Vector hướng xuống (Z âm) vì sprinkler thường ở dưới pap
            XYZ zAxisDown = -XYZ.BasisZ;

            double dotProduct = normalizedVector.DotProduct(zAxisDown);
            // Clamp để tránh lỗi acos do làm tròn số
            dotProduct = Math.Max(-1.0, Math.Min(1.0, dotProduct));

            return Math.Acos(dotProduct);
        }

        /// <summary>
        /// Tìm Sprinkler trong chuỗi kết nối từ Pap
        /// </summary>
        public static Element FindSprinklerInChain(Element pap, Document doc)
        {
            var chainElements = GetAllConnectedElementsChain(pap);
            
            foreach (var element in chainElements)
            {
                if (element is FamilyInstance fi)
                {
                    // Kiểm tra xem có phải Sprinkler không (category)
                    if (fi.Category != null && 
                        fi.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Sprinklers)
                    {
                        return element;
                    }
                }
            }
            return null;
        }

        /// <summary>
        /// Tìm ống DN65 kết nối trực tiếp với Pap
        /// </summary>
        public static Pipe FindConnectedPipe(Element pap)
        {
            var connectedElements = GetConnectedElements(pap);
            foreach (var element in connectedElements)
            {
                if (element is Pipe pipe)
                {
                    return pipe;
                }
            }
            return null;
        }

        /// <summary>
        /// Lấy connector của Pipe kết nối với Pap
        /// </summary>
        public static Connector GetPipeConnectorToPap(Element pap, Pipe pipe)
        {
            var (papConnector, pipeConnector) = GetConnectingConnectors(pap, pipe);
            return pipeConnector;
        }

        /// <summary>
        /// Quay cả cụm Pap-Pipe-Sprinkler quanh tâm connector của ống DN65
        /// để vector Pap-Sprinkler song song với trục Z
        /// </summary>
        /// <param name="doc">Document</param>
        /// <param name="pap">Pap element</param>
        /// <param name="rotationCenter">Tâm quay (connector của pipe với pap)</param>
        /// <returns>True nếu quay thành công</returns>
        public static bool RotateAssemblyToAlignWithZ(Document doc, Element pap, XYZ rotationCenter)
        {
            try
            {
                // Tìm Sprinkler trong chuỗi
                Element sprinkler = FindSprinklerInChain(pap, doc);
                if (sprinkler == null)
                {
                    return false;
                }

                // Lấy vị trí của Pap và Sprinkler
                XYZ papPosition = GetElementLocation(pap);
                XYZ sprinklerPosition = GetElementLocation(sprinkler);

                if (papPosition == null || sprinklerPosition == null)
                {
                    return false;
                }

                // Kiểm tra đã song song Z chưa
                if (IsParallelToZAxis(papPosition, sprinklerPosition))
                {
                    return true; // Đã song song, không cần quay
                }

                // Tính vector hiện tại từ Pap đến Sprinkler
                XYZ currentVector = sprinklerPosition - papPosition;
                if (currentVector.IsZeroLength()) return true;

                // Project vector lên mặt phẳng XY để tính góc quay
                XYZ currentVectorXY = new XYZ(currentVector.X, currentVector.Y, 0);
                
                // Nếu vector đã thẳng đứng trong XY thì không cần quay quanh Z
                if (currentVectorXY.IsZeroLength())
                {
                    return true;
                }

                // Tính góc cần quay để vector XY về 0 (thẳng đứng)
                // Vector mục tiêu là (0, 0, -1) hoặc (0, 0, 1)
                XYZ targetVector = currentVector.Z < 0 ? -XYZ.BasisZ : XYZ.BasisZ;
                
                // Tính góc quay quanh trục Z
                double angleAroundZ = Math.Atan2(currentVectorXY.Y, currentVectorXY.X);
                
                // Tạo trục quay (trục Z đi qua tâm quay)
                Line rotationAxis = Line.CreateBound(rotationCenter, rotationCenter + XYZ.BasisZ);

                // Thu thập tất cả element cần quay
                var elementsToRotate = new List<ElementId>();
                elementsToRotate.Add(pap.Id);
                
                var chainElements = GetAllConnectedElementsChain(pap);
                foreach (var element in chainElements)
                {
                    ConnectionHelper.UnpinElementIfPinned(doc, element);
                    elementsToRotate.Add(element.Id);
                }

                // Unpin pap
                ConnectionHelper.UnpinElementIfPinned(doc, pap);

                // Quay tất cả element quanh trục Z
                // Góc quay = -angleAroundZ để đưa vector về song song Z
                ElementTransformUtils.RotateElements(doc, elementsToRotate, rotationAxis, -angleAroundZ);

                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Lấy vị trí của element (LocationPoint hoặc trung điểm LocationCurve)
        /// </summary>
        public static XYZ GetElementLocation(Element element)
        {
            if (element.Location is LocationPoint locPoint)
            {
                return locPoint.Point;
            }
            else if (element.Location is LocationCurve locCurve)
            {
                // Lấy trung điểm của curve
                return (locCurve.Curve.GetEndPoint(0) + locCurve.Curve.GetEndPoint(1)) / 2.0;
            }
            
            // Fallback: lấy connector đầu tiên
            var connectorManager = GetConnectorManager(element);
            if (connectorManager != null)
            {
                foreach (Connector c in connectorManager.Connectors)
                {
                    return c.Origin;
                }
            }
            
            return null;
        }

        /// <summary>
        /// Xử lý hoàn chỉnh: Căn chỉnh XY + Quay để song song Z
        /// </summary>
        public static AlignmentResult AlignAndRotateToZ(Document doc, Element pap)
        {
            var result = new AlignmentResult();

            try
            {
                // Bước 1: Căn chỉnh XY trước
                var alignResult = AlignConnectedChain(doc, pap);
                result.ElementsAligned = alignResult.ElementsAligned;
                result.ElementsNeedAlignment = alignResult.ElementsNeedAlignment;

                if (!alignResult.Success && !alignResult.AlreadyAligned)
                {
                    result.ErrorMessage = alignResult.ErrorMessage;
                    return result;
                }

                // Bước 2: Tìm ống pipe kết nối với Pap để lấy tâm quay
                Pipe connectedPipe = FindConnectedPipe(pap);
                if (connectedPipe == null)
                {
                    result.ErrorMessage = "Không tìm thấy ống kết nối với Pap";
                    return result;
                }

                // Lấy connector của pipe với pap làm tâm quay
                Connector pipeConnector = GetPipeConnectorToPap(pap, connectedPipe);
                if (pipeConnector == null)
                {
                    result.ErrorMessage = "Không tìm thấy connector của ống";
                    return result;
                }

                XYZ rotationCenter = pipeConnector.Origin;

                // Bước 3: Kiểm tra và quay nếu cần
                Element sprinkler = FindSprinklerInChain(pap, doc);
                if (sprinkler != null)
                {
                    XYZ papPosition = GetElementLocation(pap);
                    XYZ sprinklerPosition = GetElementLocation(sprinkler);

                    if (papPosition != null && sprinklerPosition != null)
                    {
                        double angle = GetAngleWithZAxis(papPosition, sprinklerPosition);
                        double angleDegrees = angle * 180.0 / Math.PI;

                        if (angleDegrees > 0.1) // Góc > 0.1 độ thì quay
                        {
                            result.RotationApplied = true;
                            result.RotationAngle = angleDegrees;

                            bool rotated = RotateAssemblyToAlignWithZ(doc, pap, rotationCenter);
                            if (!rotated)
                            {
                                result.ErrorMessage = "Không thể quay cụm element";
                                return result;
                            }
                        }
                    }
                }

                result.Success = true;
            }
            catch (Exception ex)
            {
                result.ErrorMessage = ex.Message;
            }

            return result;
        }

        /// <summary>
        /// Lấy ConnectorManager từ element
        /// </summary>
        private static ConnectorManager GetConnectorManager(Element element)
        {
            if (element is MEPCurve mepCurve)
                return mepCurve.ConnectorManager;

            if (element is FamilyInstance familyInstance)
                return familyInstance.MEPModel?.ConnectorManager;

            return null;
        }
    }

    /// <summary>
    /// Kết quả của quá trình căn chỉnh
    /// </summary>
    public class AlignmentResult
    {
        public bool Success { get; set; } = false;
        public bool AlreadyAligned { get; set; } = false;
        public bool RotationApplied { get; set; } = false;
        public double RotationAngle { get; set; } = 0; // Góc đã quay (degrees)
        public List<Element> ElementsNeedAlignment { get; set; } = new List<Element>();
        public List<Element> ElementsAligned { get; set; } = new List<Element>();
        public string ErrorMessage { get; set; } = "";
    }
}
