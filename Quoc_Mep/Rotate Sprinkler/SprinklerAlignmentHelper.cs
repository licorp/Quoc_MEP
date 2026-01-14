using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;

namespace Quoc_MEP
{
    /// <summary>
    /// Helper class để xử lý việc căn chỉnh Sprinkler thẳng hàng với Pap
    /// Kiểm tra theo trục Z (vertical) - tức là X-Y phải khớp nhau
    /// </summary>
    public static partial class SprinklerAlignmentHelper
    {
        // Tolerance mặc định: 1mm = 1/304.8 feet
        private const double DEFAULT_TOLERANCE = 1.0 / 304.8;

        /// <summary>
        /// Xóa tất cả dimensions liên quan đến các element
        /// </summary>
        private static List<ElementId> DeleteDimensionsForElements(Document doc, ICollection<ElementId> elementIds)
        {
            var deletedDimensions = new List<ElementId>();
            
            try
            {
                // Tìm tất cả dimensions trong document
                var collector = new FilteredElementCollector(doc)
                    .OfClass(typeof(Dimension));

                foreach (Dimension dim in collector)
                {
                    bool shouldDelete = false;
                    
                    // Kiểm tra xem dimension có reference đến bất kỳ element nào trong danh sách không
                    foreach (Reference dimRef in dim.References)
                    {
                        if (dimRef != null && dimRef.ElementId != null && elementIds.Contains(dimRef.ElementId))
                        {
                            shouldDelete = true;
                            break;
                        }
                    }
                    
                    if (shouldDelete)
                    {
                        try
                        {
                            doc.Delete(dim.Id);
                            deletedDimensions.Add(dim.Id);
                            LogHelper.Log($"[DIMENSION] Đã xóa dimension {dim.Id} để tránh lỗi parallel");
                        }
                        catch (Exception ex)
                        {
                            LogHelper.Log($"[DIMENSION] Không thể xóa dimension {dim.Id}: {ex.Message}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogHelper.Log($"[DIMENSION] Lỗi khi xóa dimensions: {ex.Message}");
            }
            
            return deletedDimensions;
        }

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
        /// Tìm connector của Pap hướng lên (theo trục Z dương) - connector kết nối với ống chính phía trên
        /// </summary>
        public static Connector GetUpwardConnector(Element pap)
        {
            try
            {
                ConnectorManager connectorManager = GetConnectorManager(pap);
                if (connectorManager == null) return null;

                Connector upwardConnector = null;
                double highestZ = double.MinValue;

                foreach (Connector connector in connectorManager.Connectors)
                {
                    // Tìm connector có Z cao nhất (hướng lên)
                    if (connector.Origin.Z > highestZ)
                    {
                        highestZ = connector.Origin.Z;
                        upwardConnector = connector;
                    }
                }

                return upwardConnector;
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
        /// Tìm ống kết nối trực tiếp với Pap
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
        /// Tìm ống chính (main pipe) mà Pap gán vào (ống được khoan lỗ, ví dụ ống 65)
        /// Hỗ trợ 2 loại:
        /// - Loại 1: Pap → Sprinkler (không có pipe nhỏ)
        /// - Loại 2: Pap → Pipe → Sprinkler (có pipe nhỏ nối xuống)
        /// Ống chính là ống mà Pap kết nối vào theo hướng ngang (không phải hướng xuống)
        /// </summary>
        public static Pipe FindMainPipe(Element pap)
        {
            var connectedElements = GetConnectedElements(pap);
            Pipe mainPipe = null;
            Pipe verticalPipe = null;
            
            // Lấy connectors của Pap để xác định hướng kết nối
            ConnectorManager papCM = GetConnectorManager(pap);
            
            foreach (var element in connectedElements)
            {
                if (element is Pipe pipe)
                {
                    // Lấy hướng của pipe
                    LocationCurve locCurve = pipe.Location as LocationCurve;
                    if (locCurve != null)
                    {
                        XYZ pipeDirection = (locCurve.Curve.GetEndPoint(1) - locCurve.Curve.GetEndPoint(0)).Normalize();
                        
                        // Kiểm tra pipe có nằm ngang không (Z component nhỏ)
                        bool isHorizontal = Math.Abs(pipeDirection.Z) < 0.5;
                        
                        if (isHorizontal)
                        {
                            // Đây có thể là ống chính
                            if (mainPipe == null)
                            {
                                mainPipe = pipe;
                            }
                            else
                            {
                                // So sánh đường kính, chọn ống lớn hơn
                                double currentDiameter = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM)?.AsDouble() ?? 0;
                                double existingDiameter = mainPipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM)?.AsDouble() ?? 0;
                                if (currentDiameter > existingDiameter)
                                {
                                    mainPipe = pipe;
                                }
                            }
                        }
                        else
                        {
                            // Đây là ống thẳng đứng (pipe nhỏ nối xuống sprinkler)
                            verticalPipe = pipe;
                        }
                    }
                }
            }
            
            // Nếu tìm thấy ống nằm ngang, đó là ống chính
            if (mainPipe != null)
            {
                return mainPipe;
            }
            
            // Nếu không có ống nằm ngang, có thể Pap kết nối trực tiếp với Sprinkler
            // Trong trường hợp này, cần tìm ống chính thông qua connector
            // Kiểm tra connector của Pap có kết nối với ống nào không thông qua hướng connector
            if (papCM != null)
            {
                Connector horizontalConnector = null;
                
                foreach (Connector c in papCM.Connectors)
                {
                    // Connector nằm ngang (hướng X hoặc Y lớn hơn Z)
                    XYZ dir = c.CoordinateSystem.BasisZ; // Hướng của connector
                    if (Math.Abs(dir.Z) < 0.5) // Connector hướng ngang
                    {
                        horizontalConnector = c;
                        
                        // Kiểm tra connector này kết nối với gì
                        if (c.IsConnected)
                        {
                            foreach (Connector connectedC in c.AllRefs)
                            {
                                if (connectedC.Owner is Pipe foundPipe && foundPipe.Id != pap.Id)
                                {
                                    return foundPipe;
                                }
                            }
                        }
                    }
                }
            }
            
            // Fallback: trả về pipe thẳng đứng nếu có
            return verticalPipe;
        }

        /// <summary>
        /// Lấy center line của ống (đường thẳng đi qua tâm ống)
        /// </summary>
        public static Line GetPipeCenterLine(Pipe pipe)
        {
            if (pipe.Location is LocationCurve locCurve)
            {
                Curve curve = locCurve.Curve;
                if (curve is Line line)
                {
                    return line;
                }
                // Nếu là curve, tạo line từ 2 đầu
                return Line.CreateBound(curve.GetEndPoint(0), curve.GetEndPoint(1));
            }
            return null;
        }

        /// <summary>
        /// Lấy vector/hướng của Pap (hướng từ connector này sang connector kia)
        /// </summary>
        public static XYZ GetPapDirection(Element pap)
        {
            ConnectorManager cm = GetConnectorManager(pap);
            if (cm == null) return null;

            var connectors = new List<Connector>();
            foreach (Connector c in cm.Connectors)
            {
                connectors.Add(c);
            }

            if (connectors.Count >= 2)
            {
                // LUÔN trả về vector từ connector cao hơn (Z lớn hơn) đến connector thấp hơn (Z nhỏ hơn)
                // Đảm bảo hướng LUÔN là từ trên xuống dưới
                Connector higherConnector = connectors[0].Origin.Z > connectors[1].Origin.Z ? connectors[0] : connectors[1];
                Connector lowerConnector = connectors[0].Origin.Z > connectors[1].Origin.Z ? connectors[1] : connectors[0];
                
                // Vector từ trên xuống dưới
                return (lowerConnector.Origin - higherConnector.Origin).Normalize();
            }
            else if (connectors.Count == 1)
            {
                // Nếu chỉ có 1 connector, dùng hướng Z hướng xuống
                return -XYZ.BasisZ;
            }

            return null;
        }

        /// <summary>
        /// Tính giao điểm giữa center line của ống chính và vector của Pap
        /// Đây là tâm quay chính xác
        /// </summary>
        /// <param name="pap">Pap element</param>
        /// <param name="mainPipe">Ống chính (ống được khoan lỗ)</param>
        /// <returns>Điểm giao nhau làm tâm quay</returns>
        public static XYZ GetRotationCenter(Element pap, Pipe mainPipe)
        {
            // Lấy center line của ống chính
            Line pipeCenterLine = GetPipeCenterLine(mainPipe);
            if (pipeCenterLine == null) return null;

            // Lấy vị trí và hướng của Pap
            XYZ papLocation = GetElementLocation(pap);
            XYZ papDirection = GetPapDirection(pap);
            
            if (papLocation == null || papDirection == null) return null;

            // Tạo đường thẳng từ Pap theo hướng của nó
            // Kéo dài đường thẳng đủ xa để cắt ống
            XYZ papLineStart = papLocation - papDirection * 100; // 100 feet về phía sau
            XYZ papLineEnd = papLocation + papDirection * 100;   // 100 feet về phía trước
            Line papLine = Line.CreateBound(papLineStart, papLineEnd);

            // Tìm giao điểm giữa 2 đường thẳng
            // Vì 2 đường có thể không giao nhau chính xác (skew lines), 
            // ta tìm điểm gần nhất trên pipe center line đến pap line
            XYZ closestPointOnPipe = GetClosestPointOnLine(pipeCenterLine, papLocation);
            
            if (closestPointOnPipe != null)
            {
                return closestPointOnPipe;
            }

            // Fallback: dùng connector của pap với ống chính
            var (papConnector, pipeConnector) = GetConnectingConnectors(pap, mainPipe);
            if (pipeConnector != null)
            {
                return pipeConnector.Origin;
            }

            return papLocation;
        }

        /// <summary>
        /// Tìm điểm gần nhất trên đường thẳng đến một điểm cho trước
        /// </summary>
        public static XYZ GetClosestPointOnLine(Line line, XYZ point)
        {
            XYZ lineStart = line.GetEndPoint(0);
            XYZ lineEnd = line.GetEndPoint(1);
            XYZ lineDirection = (lineEnd - lineStart).Normalize();
            
            // Vector từ line start đến point
            XYZ startToPoint = point - lineStart;
            
            // Project point lên line
            double projectionLength = startToPoint.DotProduct(lineDirection);
            
            // Điểm gần nhất trên đường thẳng (không giới hạn trong đoạn thẳng)
            XYZ closestPoint = lineStart + lineDirection * projectionLength;
            
            return closestPoint;
        }

        /// <summary>
        /// Quay Pap để hướng của nó thẳng đứng (song song với trục Z)
        /// Đơn giản: Chỉ kiểm tra hướng Pap với trục Z, nếu lệch thì quay
        /// </summary>
        /// <param name="doc">Document</param>
        /// <param name="pap">Pap element</param>
        /// <param name="rotationCenter">Tâm quay (giao điểm của ống 65 với Pap)</param>
        /// <param name="mainPipe">Ống chính (ống 65) để lấy trục quay</param>
        /// <param name="dimensionsDeleted">Output: số dimensions đã xóa</param>
        /// <returns>True nếu quay thành công</returns>
        public static bool RotatePapToVertical(Document doc, Element pap, XYZ rotationCenter, Pipe mainPipe, out int dimensionsDeleted)
        {
            dimensionsDeleted = 0;
            try
            {
                LogHelper.Log($"[ROTATE] Bắt đầu quay Pap {pap.Id}");
                
                // Lấy hướng của Pap
                XYZ papDirection = GetPapDirection(pap);
                if (papDirection == null || papDirection.IsZeroLength())
                {
                    LogHelper.Log("[ROTATE] Không xác định được hướng Pap");
                    return false;
                }

                LogHelper.Log($"[ROTATE] Hướng Pap ban đầu: ({papDirection.X:F4}, {papDirection.Y:F4}, {papDirection.Z:F4})");

                // Xác định trục quay = trục của ống 65 (ống chính)
                XYZ rotationAxisDirection;
                Line pipeCenterLine = null;
                if (mainPipe != null)
                {
                    pipeCenterLine = GetPipeCenterLine(mainPipe);
                    if (pipeCenterLine != null)
                    {
                        rotationAxisDirection = (pipeCenterLine.GetEndPoint(1) - pipeCenterLine.GetEndPoint(0)).Normalize();
                        LogHelper.Log($"[ROTATE] Trục xoay (center line ống 65): ({rotationAxisDirection.X:F4}, {rotationAxisDirection.Y:F4}, {rotationAxisDirection.Z:F4})");
                        LogHelper.Log($"[ROTATE] Điểm giao (tâm xoay): ({rotationCenter.X:F4}, {rotationCenter.Y:F4}, {rotationCenter.Z:F4})");
                    }
                    else
                    {
                        LogHelper.Log("[ROTATE] CẢNH BÁO: Không lấy được center line ống, dùng trục Z");
                        rotationAxisDirection = XYZ.BasisZ;
                    }
                }
                else
                {
                    LogHelper.Log("[ROTATE] CẢNH BÁO: Không có ống chính, dùng trục Z");
                    rotationAxisDirection = XYZ.BasisZ;
                }
                
                // Vector mục tiêu = trục thẳng đứng hướng XUỐNG (negative Z)
                XYZ targetVector = -XYZ.BasisZ; // Hướng xuống dưới
                
                // Tính góc lệch CHÍNH XÁC của Pap này
                double dotProduct = papDirection.DotProduct(targetVector);
                double angleDegrees = Math.Acos(Math.Max(-1.0, Math.Min(1.0, Math.Abs(dotProduct)))) * 180.0 / Math.PI;
                
                LogHelper.Log($"[ROTATE] ========================================");
                LogHelper.Log($"[ROTATE] Pap {pap.Id} - GÓC LỆCH BAN ĐẦU: {angleDegrees:F4}°");
                LogHelper.Log($"[ROTATE] dotProduct với trục xuống: {dotProduct:F4}");
                LogHelper.Log($"[ROTATE] ========================================");

                // Nếu đã thẳng đứng VÀ hướng xuống (góc < 0.1 độ và dotProduct > 0.999), không cần quay
                if (angleDegrees < 0.1 && dotProduct > 0.999)
                {
                    LogHelper.Log($"[ROTATE] Pap {pap.Id} đã thẳng đứng và hướng xuống, không cần quay");
                    return true;
                }
                
                LogHelper.Log($"[ROTATE] Pap {pap.Id} CẦN XOAY để căn chỉnh");

                // Lặp tối đa 3 lần để đạt độ chính xác cao
                int maxIterations = 3;
                
                // CHỈ QUAY PAP - không disconnect vì Pap là TAP connection
                var elementsToRotate = new List<ElementId> { pap.Id };
                
                // Unpin pap (chỉ làm 1 lần)
                ConnectionHelper.UnpinElementIfPinned(doc, pap);

                // Xóa dimensions trước khi quay (chỉ làm 1 lần)
                LogHelper.Log("[ROTATE] Đang xóa dimensions liên quan...");
                var deletedDimensions = DeleteDimensionsForElements(doc, elementsToRotate);
                dimensionsDeleted = deletedDimensions.Count;
                if (dimensionsDeleted > 0)
                {
                    LogHelper.Log($"[ROTATE] Đã xóa {dimensionsDeleted} dimensions");
                }
                
                for (int iteration = 0; iteration < maxIterations; iteration++)
                {
                    LogHelper.Log($"[ROTATE] === Pap {pap.Id} - Vòng lặp {iteration + 1}/{maxIterations} ===");
                    
                    // Cập nhật lại hướng Pap sau mỗi lần xoay
                    if (iteration > 0)
                    {
                        papDirection = GetPapDirection(pap);
                        if (papDirection == null || papDirection.IsZeroLength())
                        {
                            LogHelper.Log("[ROTATE] CẢNH BÁO: Không xác định được hướng Pap sau xoay");
                            break;
                        }
                        
                        // Tính lại góc lệch CHÍNH XÁC sau khi xoay
                        dotProduct = papDirection.DotProduct(targetVector);
                        angleDegrees = Math.Acos(Math.Max(-1.0, Math.Min(1.0, Math.Abs(dotProduct)))) * 180.0 / Math.PI;
                        
                        LogHelper.Log($"[ROTATE] Pap {pap.Id} sau iteration {iteration}:");
                        LogHelper.Log($"[ROTATE]   - Hướng mới: ({papDirection.X:F4}, {papDirection.Y:F4}, {papDirection.Z:F4})");
                        LogHelper.Log($"[ROTATE]   - Góc lệch còn lại: {angleDegrees:F4}°");
                        LogHelper.Log($"[ROTATE]   - dotProduct: {dotProduct:F4}");
                        
                        if (angleDegrees < 0.1 && dotProduct > 0.999)
                        {
                            LogHelper.Log($"[ROTATE] Pap {pap.Id} đã đạt độ chính xác mong muốn!");
                            break;
                        }
                    }

                // Project hướng Pap lên mặt phẳng vuông góc với trục quay
                XYZ projectedPapDir = papDirection - rotationAxisDirection * papDirection.DotProduct(rotationAxisDirection);
                if (projectedPapDir.IsZeroLength())
                {
                    LogHelper.Log("[ROTATE] Hướng Pap đã song song với trục quay");
                    return true;
                }
                projectedPapDir = projectedPapDir.Normalize();

                // Project target vector lên cùng mặt phẳng
                XYZ projectedTargetDir = targetVector - rotationAxisDirection * targetVector.DotProduct(rotationAxisDirection);
                if (projectedTargetDir.IsZeroLength())
                {
                    LogHelper.Log("[ROTATE] Target vector song song với trục quay");
                    return true;
                }
                projectedTargetDir = projectedTargetDir.Normalize();

                // Tính góc cần quay (từ Pap về Target) - CHÍNH XÁC cho Pap này
                double dot = projectedPapDir.DotProduct(projectedTargetDir);
                dot = Math.Max(-1.0, Math.Min(1.0, dot));
                double angle = Math.Acos(dot);
                
                LogHelper.Log($"[ROTATE] === TÍNH GÓC XOAY CHO PAP {pap.Id} ===");
                LogHelper.Log($"[ROTATE] Góc tính được (radians): {angle:F6}");
                LogHelper.Log($"[ROTATE] Góc tính được (degrees): {angle * 180 / Math.PI:F4}°");

                // QUAN TRỌNG: Không được xoay quá 90 độ
                // Nếu góc > 90°, xoay ngược chiều (180° - angle)
                if (Math.Abs(angle) > Math.PI / 2)
                {
                    angle = Math.PI - angle;
                    LogHelper.Log($"[ROTATE] Góc > 90°, điều chỉnh thành: {angle * 180 / Math.PI:F4}°");
                }

                // Xác định chiều quay (từ Pap về Target)
                XYZ crossProduct = projectedPapDir.CrossProduct(projectedTargetDir);
                double crossDotAxis = crossProduct.DotProduct(rotationAxisDirection);
                if (crossDotAxis < 0)
                {
                    angle = -angle;
                    LogHelper.Log($"[ROTATE] Đảo chiều quay (cross product: {crossDotAxis:F4})");
                }
                
                LogHelper.Log($"[ROTATE] GÓC XOAY CUỐI CÙNG cho Pap {pap.Id}: {angle * 180 / Math.PI:F4}°");
                LogHelper.Log($"[ROTATE] ====================================");

                // Nếu góc quá nhỏ, không cần quay nữa
                if (Math.Abs(angle) < 0.0001) // ~0.0057 độ
                {
                    LogHelper.Log($"[ROTATE] Pap {pap.Id}: Góc quá nhỏ ({Math.Abs(angle) * 180 / Math.PI:F4}°), kết thúc iteration");
                    break;
                }

                // Tạo trục quay qua tâm quay (trục của ống 65)
                Line rotationAxis = Line.CreateBound(rotationCenter, rotationCenter + rotationAxisDirection);

                    // Thực hiện xoay Pap quanh center line của ống 65
                    LogHelper.Log($"[ROTATE] >>> Thực hiện xoay Pap {pap.Id}: {angle * 180 / Math.PI:F4}° quanh tâm ({rotationCenter.X:F3}, {rotationCenter.Y:F3}, {rotationCenter.Z:F3})");
                    ElementTransformUtils.RotateElements(doc, elementsToRotate, rotationAxis, angle);
                    LogHelper.Log($"[ROTATE] >>> Xoay Pap {pap.Id} hoàn thành!");
                } // end iteration loop

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
        /// Tìm pipe hoặc pipe fitting size 40mm gần Pap
        /// </summary>
        /// <param name="doc">Document</param>
        /// <param name="pap">Pap element</param>
        /// <param name="searchRadius">Bán kính tìm kiếm (feet), mặc định 10 feet</param>
        /// <returns>Danh sách pipe/fitting 40mm trong bán kính</returns>
        public static List<Element> Find40mmElementsNearPap(Document doc, Element pap, double searchRadius = 10.0)
        {
            var result = new List<Element>();
            
            try
            {
                XYZ papLocation = GetElementLocation(pap);
                if (papLocation == null)
                {
                    LogHelper.Log("[FIND_40MM] Không xác định được vị trí Pap");
                    return result;
                }

                LogHelper.Log($"[FIND_40MM] Tìm kiếm pipe/fitting 40mm quanh Pap {pap.Id} (bán kính: {searchRadius * 304.8:F0}mm)");

                // Tạo bounding box quanh Pap
                XYZ minPoint = papLocation - new XYZ(searchRadius, searchRadius, searchRadius);
                XYZ maxPoint = papLocation + new XYZ(searchRadius, searchRadius, searchRadius);
                Outline outline = new Outline(minPoint, maxPoint);
                BoundingBoxIntersectsFilter bbFilter = new BoundingBoxIntersectsFilter(outline);

                // Tìm tất cả pipe và pipe fitting trong bounding box
                var pipeCollector = new FilteredElementCollector(doc)
                    .OfClass(typeof(Pipe))
                    .WherePasses(bbFilter);

                var fittingCollector = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_PipeFitting)
                    .WherePasses(bbFilter);

                // Kiểm tra pipe
                foreach (Pipe pipe in pipeCollector)
                {
                    if (pipe.Id == pap.Id) continue;

                    // Kiểm tra khoảng cách
                    XYZ pipeLocation = GetElementLocation(pipe);
                    if (pipeLocation != null && pipeLocation.DistanceTo(papLocation) <= searchRadius)
                    {
                        // Kiểm tra đường kính
                        Parameter diamParam = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM);
                        if (diamParam != null)
                        {
                            double diameterMM = diamParam.AsDouble() * 304.8;
                            // Cho phép sai số ±3mm
                            if (Math.Abs(diameterMM - 40) < 3)
                            {
                                result.Add(pipe);
                                LogHelper.Log($"[FIND_40MM] Tìm thấy Pipe 40mm: {pipe.Id}, khoảng cách: {pipeLocation.DistanceTo(papLocation) * 304.8:F0}mm");
                            }
                        }
                    }
                }

                // Kiểm tra pipe fitting
                foreach (Element fitting in fittingCollector)
                {
                    if (fitting.Id == pap.Id) continue;

                    XYZ fittingLocation = GetElementLocation(fitting);
                    if (fittingLocation != null && fittingLocation.DistanceTo(papLocation) <= searchRadius)
                    {
                        // Kiểm tra size của fitting (thông qua connector)
                        ConnectorManager cm = GetConnectorManager(fitting);
                        if (cm != null)
                        {
                            foreach (Connector conn in cm.Connectors)
                            {
                                double diameterMM = conn.Radius * 2 * 304.8;
                                if (Math.Abs(diameterMM - 40) < 3)
                                {
                                    result.Add(fitting);
                                    LogHelper.Log($"[FIND_40MM] Tìm thấy Fitting 40mm: {fitting.Id}, khoảng cách: {fittingLocation.DistanceTo(papLocation) * 304.8:F0}mm");
                                    break;
                                }
                            }
                        }
                    }
                }

                LogHelper.Log($"[FIND_40MM] Tổng cộng tìm thấy {result.Count} đối tượng 40mm");
            }
            catch (Exception ex)
            {
                LogHelper.Log($"[FIND_40MM] Lỗi: {ex.Message}");
            }

            return result;
        }

        /// <summary>
        /// Tìm chuỗi kết nối từ Pap xuống Sprinkler
        /// Trả về: Pipe 40mm (nếu có), Fitting, Sprinkler (nếu có)
        /// </summary>
        public static (Pipe pipe40, Element fitting, Element sprinkler) FindPapChain(Document doc, Element pap)
        {
            try
            {
                LogHelper.Log($"[FIND_CHAIN] === Tìm chuỗi kết nối từ Pap {pap.Id} ===");
                
                Pipe pipe40 = null;
                Element fitting = null;
                Element sprinkler = null;

                // Lấy connector hướng xuống của Pap
                Connector papDownConnector = GetDownwardConnector(pap);
                if (papDownConnector == null)
                {
                    LogHelper.Log("[FIND_CHAIN] Pap không có connector hướng xuống");
                    return (null, null, null);
                }

                LogHelper.Log($"[FIND_CHAIN] Connector Pap IsConnected: {papDownConnector.IsConnected}");

                // Tìm element kết nối trực tiếp với Pap (nếu đang connected)
                Element connectedToPap = null;
                if (papDownConnector.IsConnected)
                {
                    foreach (Connector c in papDownConnector.AllRefs)
                    {
                        if (c.Owner.Id != pap.Id)
                        {
                            connectedToPap = c.Owner;
                            break;
                        }
                    }
                }

                // Nếu không connected, tìm theo khoảng cách
                if (connectedToPap == null)
                {
                    LogHelper.Log("[FIND_CHAIN] Connector không connected, tìm theo khoảng cách (5mm)");
                    XYZ papConnPos = papDownConnector.Origin;
                    double searchRadius = 5.0 / 304.8; // 5mm
                    
                    // Tìm pipe/fitting gần connector
                    List<Element> nearElements = Find40mmElementsNearPap(doc, pap, searchRadius);
                    
                    if (nearElements.Count > 0)
                    {
                        // Tìm element gần connector nhất
                        double minDist = double.MaxValue;
                        foreach (var elem in nearElements)
                        {
                            ConnectorManager cm = GetConnectorManager(elem);
                            if (cm != null)
                            {
                                foreach (Connector conn in cm.Connectors)
                                {
                                    double dist = conn.Origin.DistanceTo(papConnPos);
                                    if (dist < minDist)
                                    {
                                        minDist = dist;
                                        connectedToPap = elem;
                                    }
                                }
                            }
                        }
                        
                        if (connectedToPap != null)
                        {
                            LogHelper.Log($"[FIND_CHAIN] Tìm thấy element gần nhất: {connectedToPap.Id} (khoảng cách: {minDist * 304.8:F1}mm)");
                        }
                    }
                }

                if (connectedToPap == null)
                {
                    LogHelper.Log("[FIND_CHAIN] Không tìm thấy element kết nối với Pap (cả connected và khoảng cách)");
                    return (null, null, null);
                }

                LogHelper.Log($"[FIND_CHAIN] Element kết nối trực tiếp với Pap: {connectedToPap.Id} - {connectedToPap.Category?.Name}");

                // TRƯỜNG HỢP 1: Pap → Pipe 40mm → Fitting → Sprinkler
                if (connectedToPap is Pipe)
                {
                    Pipe pipe = connectedToPap as Pipe;
                    Parameter diamParam = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM);
                    if (diamParam != null)
                    {
                        double diameterMM = diamParam.AsDouble() * 304.8;
                        if (Math.Abs(diameterMM - 40) < 3)
                        {
                            pipe40 = pipe;
                            LogHelper.Log($"[FIND_CHAIN] ✓ Trường hợp 1: Tìm thấy Pipe 40mm: {pipe40.Id}");

                            // Tìm fitting kết nối với pipe này
                            ConnectorManager pipeCM = GetConnectorManager(pipe40);
                            if (pipeCM != null)
                            {
                                foreach (Connector pipeConn in pipeCM.Connectors)
                                {
                                    if (pipeConn.IsConnected)
                                    {
                                        foreach (Connector connectedConn in pipeConn.AllRefs)
                                        {
                                            Element connectedElem = connectedConn.Owner;
                                            if (connectedElem.Id != pipe40.Id && connectedElem.Id != pap.Id)
                                            {
                                                // Kiểm tra xem có phải fitting không
                                                if (connectedElem.Category?.Id.IntegerValue == (int)BuiltInCategory.OST_PipeFitting)
                                                {
                                                    fitting = connectedElem;
                                                    LogHelper.Log($"[FIND_CHAIN] ✓ Tìm thấy Fitting: {fitting.Id}");

                                                    // Tìm sprinkler nối với fitting
                                                    sprinkler = FindSprinklerConnectedTo(fitting);
                                                    if (sprinkler != null)
                                                    {
                                                        LogHelper.Log($"[FIND_CHAIN] ✓ Tìm thấy Sprinkler: {sprinkler.Id}");
                                                    }
                                                    break;
                                                }
                                            }
                                        }
                                    }
                                    if (fitting != null) break;
                                }
                            }
                        }
                    }
                }
                // TRƯỜNG HỢP 2: Pap → Fitting → Sprinkler (không có pipe 40mm)
                else if (connectedToPap.Category?.Id.IntegerValue == (int)BuiltInCategory.OST_PipeFitting)
                {
                    fitting = connectedToPap;
                    LogHelper.Log($"[FIND_CHAIN] ✓ Trường hợp 2: Fitting kết nối trực tiếp: {fitting.Id}");

                    // Tìm sprinkler
                    sprinkler = FindSprinklerConnectedTo(fitting);
                    if (sprinkler != null)
                    {
                        LogHelper.Log($"[FIND_CHAIN] ✓ Tìm thấy Sprinkler: {sprinkler.Id}");
                    }
                }

                LogHelper.Log($"[FIND_CHAIN] === Kết quả: Pipe40={pipe40?.Id}, Fitting={fitting?.Id}, Sprinkler={sprinkler?.Id} ===");
                return (pipe40, fitting, sprinkler);
            }
            catch (Exception ex)
            {
                LogHelper.Log($"[FIND_CHAIN] Lỗi: {ex.Message}");
                return (null, null, null);
            }
        }

        /// <summary>
        /// Tìm sprinkler kết nối với element (qua connector)
        /// </summary>
        private static Element FindSprinklerConnectedTo(Element element)
        {
            try
            {
                ConnectorManager cm = GetConnectorManager(element);
                if (cm == null) return null;

                foreach (Connector conn in cm.Connectors)
                {
                    if (conn.IsConnected)
                    {
                        foreach (Connector connectedConn in conn.AllRefs)
                        {
                            Element connectedElem = connectedConn.Owner;
                            if (connectedElem.Id != element.Id)
                            {
                                // Kiểm tra xem có phải sprinkler không
                                var category = connectedElem.Category;
                                if (category != null && 
                                    (category.Id.IntegerValue == (int)BuiltInCategory.OST_Sprinklers ||
                                     category.Id.IntegerValue == (int)BuiltInCategory.OST_FireAlarmDevices))
                                {
                                    return connectedElem;
                                }
                            }
                        }
                    }
                }
            }
            catch { }
            return null;
        }

        /// <summary>
        /// Tìm sprinkler gần element theo khoảng cách
        /// </summary>
        private static Element FindSprinklerNearElement(Document doc, Element element, double searchRadius = 3.0)
        {
            try
            {
                XYZ elemLocation = GetElementLocation(element);
                if (elemLocation == null) return null;

                LogHelper.Log($"[FIND_SPRINKLER] Tìm sprinkler gần element {element.Id} (bán kính: {searchRadius * 304.8:F0}mm)");

                // Tạo bounding box
                XYZ minPoint = elemLocation - new XYZ(searchRadius, searchRadius, searchRadius);
                XYZ maxPoint = elemLocation + new XYZ(searchRadius, searchRadius, searchRadius);
                Outline outline = new Outline(minPoint, maxPoint);
                BoundingBoxIntersectsFilter bbFilter = new BoundingBoxIntersectsFilter(outline);

                // Tìm sprinklers
                var sprinklerCollector = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_Sprinklers)
                    .WherePasses(bbFilter);

                Element nearestSprinkler = null;
                double minDistance = double.MaxValue;

                foreach (Element spr in sprinklerCollector)
                {
                    XYZ sprLocation = GetElementLocation(spr);
                    if (sprLocation != null)
                    {
                        double dist = sprLocation.DistanceTo(elemLocation);
                        if (dist <= searchRadius && dist < minDistance)
                        {
                            minDistance = dist;
                            nearestSprinkler = spr;
                        }
                    }
                }

                if (nearestSprinkler != null)
                {
                    LogHelper.Log($"[FIND_SPRINKLER] Tìm thấy sprinkler {nearestSprinkler.Id} (khoảng cách: {minDistance * 304.8:F0}mm)");
                }

                return nearestSprinkler;
            }
            catch (Exception ex)
            {
                LogHelper.Log($"[FIND_SPRINKLER] Lỗi: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Kiểm tra xem element có align 3D với Pap không
        /// (X, Y, Z đều phải gần khớp nhau)
        /// </summary>
        /// <param name="pap">Pap element</param>
        /// <param name="element">Element cần kiểm tra</param>
        /// <param name="tolerance">Sai số cho phép (feet), mặc định 3mm</param>
        /// <returns>True nếu đã align 3D</returns>
        public static bool IsAligned3D(Element pap, Element element, double tolerance = 3.0 / 304.8)
        {
            try
            {
                // Lấy connector của Pap hướng xuống (connector kết nối với pipe/sprinkler)
                Connector papConnector = GetDownwardConnector(pap);
                if (papConnector == null)
                {
                    LogHelper.Log("[ALIGN_3D] Không tìm thấy connector Pap hướng xuống");
                    return false;
                }

                // Lấy connector gần nhất của element với Pap
                Connector elementConnector = null;
                double minDistance = double.MaxValue;
                
                ConnectorManager cm = GetConnectorManager(element);
                if (cm == null) return false;

                foreach (Connector conn in cm.Connectors)
                {
                    double distance = conn.Origin.DistanceTo(papConnector.Origin);
                    if (distance < minDistance)
                    {
                        minDistance = distance;
                        elementConnector = conn;
                    }
                }

                if (elementConnector == null) return false;

                // Kiểm tra X, Y, Z
                double deltaX = Math.Abs(papConnector.Origin.X - elementConnector.Origin.X);
                double deltaY = Math.Abs(papConnector.Origin.Y - elementConnector.Origin.Y);
                double deltaZ = Math.Abs(papConnector.Origin.Z - elementConnector.Origin.Z);

                bool aligned = deltaX <= tolerance && deltaY <= tolerance && deltaZ <= tolerance;

                LogHelper.Log($"[ALIGN_3D] Pap {pap.Id} vs Element {element.Id}:");
                LogHelper.Log($"[ALIGN_3D]   ΔX: {deltaX * 304.8:F2}mm, ΔY: {deltaY * 304.8:F2}mm, ΔZ: {deltaZ * 304.8:F2}mm");
                LogHelper.Log($"[ALIGN_3D]   Aligned: {aligned} (tolerance: {tolerance * 304.8:F1}mm)");

                return aligned;
            }
            catch (Exception ex)
            {
                LogHelper.Log($"[ALIGN_3D] Lỗi: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Align, move và connect element với Pap (hoặc element khác)
        /// </summary>
        /// <param name="doc">Document</param>
        /// <param name="referenceElement">Element tham chiếu (Pap hoặc Pipe)</param>
        /// <param name="elementToMove">Element cần align và di chuyển (pipe hoặc fitting 40mm)</param>
        /// <returns>True nếu thành công</returns>
        public static bool AlignMoveConnectWithPap(Document doc, Element referenceElement, Element elementToMove)
        {
            try
            {
                LogHelper.Log($"[ALIGN_MOVE_CONNECT] Bắt đầu align {elementToMove.Category?.Name} {elementToMove.Id} với {referenceElement.Category?.Name} {referenceElement.Id}");

                // Lấy connector của reference element hướng xuống
                Connector refConnector = GetDownwardConnector(referenceElement);
                if (refConnector == null)
                {
                    LogHelper.Log("[ALIGN_MOVE_CONNECT] Không tìm thấy connector reference element");
                    return false;
                }

                LogHelper.Log($"[ALIGN_MOVE_CONNECT] Connector reference tại: ({refConnector.Origin.X * 304.8:F1}, {refConnector.Origin.Y * 304.8:F1}, {refConnector.Origin.Z * 304.8:F1})mm");

                LogHelper.Log($"[ALIGN_MOVE_CONNECT] Connector reference tại: ({refConnector.Origin.X * 304.8:F1}, {refConnector.Origin.Y * 304.8:F1}, {refConnector.Origin.Z * 304.8:F1})mm");

                // Lấy connector gần nhất của element cần di chuyển
                Connector elementConnector = null;
                double minDistance = double.MaxValue;
                
                ConnectorManager cm = GetConnectorManager(elementToMove);
                if (cm == null)
                {
                    LogHelper.Log("[ALIGN_MOVE_CONNECT] Element không có connector");
                    return false;
                }

                foreach (Connector conn in cm.Connectors)
                {
                    double distance = conn.Origin.DistanceTo(refConnector.Origin);
                    if (distance < minDistance)
                    {
                        minDistance = distance;
                        elementConnector = conn;
                    }
                }

                if (elementConnector == null)
                {
                    LogHelper.Log("[ALIGN_MOVE_CONNECT] Không tìm thấy connector của element");
                    return false;
                }

                LogHelper.Log($"[ALIGN_MOVE_CONNECT] Connector element tại: ({elementConnector.Origin.X * 304.8:F1}, {elementConnector.Origin.Y * 304.8:F1}, {elementConnector.Origin.Z * 304.8:F1})mm");
                LogHelper.Log($"[ALIGN_MOVE_CONNECT] Khoảng cách hiện tại: {minDistance * 304.8:F1}mm");

                // Disconnect nếu đang kết nối
                if (elementConnector.IsConnected)
                {
                    LogHelper.Log("[ALIGN_MOVE_CONNECT] Disconnect element...");
                    try
                    {
                        var refs = new List<Connector>();
                        foreach (Connector c in elementConnector.AllRefs)
                        {
                            if (c.Owner.Id != elementToMove.Id)
                            {
                                refs.Add(c);
                            }
                        }
                        foreach (var c in refs)
                        {
                            elementConnector.DisconnectFrom(c);
                        }
                        LogHelper.Log($"[ALIGN_MOVE_CONNECT] Đã disconnect {refs.Count} connector");
                    }
                    catch (Exception ex)
                    {
                        LogHelper.Log($"[ALIGN_MOVE_CONNECT] Lỗi disconnect: {ex.Message}");
                    }
                }

                // Unpin element
                ConnectionHelper.UnpinElementIfPinned(doc, elementToMove);

                // Tính vector di chuyển để align 3D
                XYZ moveVector = refConnector.Origin - elementConnector.Origin;
                LogHelper.Log($"[ALIGN_MOVE_CONNECT] Di chuyển vector: ({moveVector.X * 304.8:F1}, {moveVector.Y * 304.8:F1}, {moveVector.Z * 304.8:F1})mm (độ dài: {moveVector.GetLength() * 304.8:F1}mm)");

                // Di chuyển element
                if (moveVector.GetLength() > 0.001) // > ~0.3mm
                {
                    ElementTransformUtils.MoveElement(doc, elementToMove.Id, moveVector);
                    LogHelper.Log("[ALIGN_MOVE_CONNECT] ✓ Đã di chuyển element");
                }
                else
                {
                    LogHelper.Log("[ALIGN_MOVE_CONNECT] Element đã ở đúng vị trí, không cần di chuyển");
                }

                // Connect lại
                try
                {
                    LogHelper.Log($"[ALIGN_MOVE_CONNECT] Kết nối: refConn.IsConnected={refConnector.IsConnected}, elemConn.IsConnected={elementConnector.IsConnected}");
                    
                    if (!refConnector.IsConnected || !elementConnector.IsConnected)
                    {
                        refConnector.ConnectTo(elementConnector);
                        LogHelper.Log("[ALIGN_MOVE_CONNECT] ✓ Đã connect element với reference");
                    }
                    else
                    {
                        LogHelper.Log("[ALIGN_MOVE_CONNECT] Cả 2 connector đã connected với element khác");
                    }
                }
                catch (Exception ex)
                {
                    LogHelper.Log($"[ALIGN_MOVE_CONNECT] Không thể connect (có thể đã auto-connect hoặc lỗi): {ex.Message}");
                    // Không return false vì element đã được align và di chuyển rồi
                }

                LogHelper.Log($"[ALIGN_MOVE_CONNECT] ✓ Hoàn thành align {elementToMove.Category?.Name} {elementToMove.Id} với {referenceElement.Category?.Name} {referenceElement.Id}");
                return true;
            }
            catch (Exception ex)
            {
                LogHelper.Log($"[ALIGN_MOVE_CONNECT] Lỗi: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Align toàn bộ chuỗi Pipe 40mm + Fitting với Pap
        /// </summary>
        public static (int alignedCount, string details) AlignChainWithPap(Document doc, Element pap)
        {
            int alignedCount = 0;
            var details = new List<string>();

            try
            {
                LogHelper.Log($"[ALIGN_CHAIN] === Bắt đầu align chuỗi với Pap {pap.Id} ===");

                // Tìm chuỗi kết nối qua connector
                var (pipe40, fitting, sprinkler) = FindPapChain(doc, pap);

                // Nếu không tìm thấy qua connector, tìm theo khoảng cách
                if (pipe40 == null && fitting == null)
                {
                    LogHelper.Log("[ALIGN_CHAIN] Không tìm thấy qua connector, tìm theo khoảng cách...");
                    
                    // Tìm tất cả pipe/fitting 40mm gần Pap (bán kính 3 feet)
                    List<Element> nearElements = Find40mmElementsNearPap(doc, pap, searchRadius: 3.0);
                    
                    if (nearElements.Count > 0)
                    {
                        LogHelper.Log($"[ALIGN_CHAIN] Tìm thấy {nearElements.Count} pipe/fitting 40mm gần Pap");
                        
                        // Phân loại pipe và fitting
                        foreach (var elem in nearElements)
                        {
                            if (elem is Pipe && pipe40 == null)
                            {
                                pipe40 = elem as Pipe;
                                LogHelper.Log($"[ALIGN_CHAIN] Chọn Pipe 40mm gần nhất: {pipe40.Id}");
                            }
                            else if (elem.Category?.Id.IntegerValue == (int)BuiltInCategory.OST_PipeFitting && fitting == null)
                            {
                                fitting = elem;
                                LogHelper.Log($"[ALIGN_CHAIN] Chọn Fitting 40mm gần nhất: {fitting.Id}");
                            }
                        }
                    }
                }

                // Luôn tìm sprinkler nếu có fitting (không phân biệt tìm qua connector hay distance)
                if (fitting != null && sprinkler == null)
                {
                    sprinkler = FindSprinklerNearElement(doc, fitting, searchRadius: 3.0);
                    if (sprinkler != null)
                    {
                        LogHelper.Log($"[ALIGN_CHAIN] Tìm thấy Sprinkler gần Fitting: {sprinkler.Id}");
                    }
                    else
                    {
                        LogHelper.Log("[ALIGN_CHAIN] Không tìm thấy Sprinkler gần Fitting trong bán kính 3 feet");
                    }
                }

                if (pipe40 == null && fitting == null)
                {
                    LogHelper.Log("[ALIGN_CHAIN] Không tìm thấy pipe 40mm hoặc fitting (cả connector và khoảng cách)");
                    details.Add("Không tìm thấy pipe/fitting 40mm trong bán kính 3 feet");
                    return (0, string.Join(", ", details));
                }

                // TRƯỜNG HỢP 1: Pap → Pipe 40mm → Fitting
                if (pipe40 != null)
                {
                    details.Add($"TH1: Pipe40 ({pipe40.Id})");
                    LogHelper.Log($"[ALIGN_CHAIN] Trường hợp 1: Align Pipe 40mm {pipe40.Id}");

                    // Kiểm tra align
                    bool pipeAligned = IsAligned3D(pap, pipe40, tolerance: 3.0 / 304.8);
                    if (!pipeAligned)
                    {
                        bool success = AlignMoveConnectWithPap(doc, pap, pipe40);
                        if (success)
                        {
                            alignedCount++;
                            details.Add("✓ Pipe40 aligned");
                            LogHelper.Log("[ALIGN_CHAIN] ✓ Pipe 40mm đã được align");
                        }
                    }
                    else
                    {
                        details.Add("Pipe40 đã align");
                    }

                    // Align fitting nếu có
                    if (fitting != null)
                    {
                        LogHelper.Log($"[ALIGN_CHAIN] Kiểm tra Fitting {fitting.Id}");
                        bool fittingAligned = IsAligned3D(pipe40, fitting, tolerance: 3.0 / 304.8);
                        if (!fittingAligned)
                        {
                            // Align fitting với pipe 40mm
                            bool success = AlignMoveConnectWithPap(doc, pipe40, fitting);
                            if (success)
                            {
                                alignedCount++;
                                details.Add($"✓ Fitting ({fitting.Id}) aligned");
                                LogHelper.Log("[ALIGN_CHAIN] ✓ Fitting đã được align với Pipe40");
                            }
                        }
                        else
                        {
                            details.Add("Fitting đã align");
                        }
                        
                        // Align sprinkler với fitting (nếu có)
                        if (sprinkler != null)
                        {
                            LogHelper.Log($"[ALIGN_CHAIN] Kiểm tra Sprinkler {sprinkler.Id}");
                            bool sprinklerAligned = IsAligned3D(fitting, sprinkler, tolerance: 3.0 / 304.8);
                            if (!sprinklerAligned)
                            {
                                bool success = AlignMoveConnectWithPap(doc, fitting, sprinkler);
                                if (success)
                                {
                                    alignedCount++;
                                    details.Add($"✓ Sprinkler ({sprinkler.Id}) aligned");
                                    LogHelper.Log("[ALIGN_CHAIN] ✓ Sprinkler đã được align với Fitting");
                                }
                            }
                            else
                            {
                                details.Add($"Sprinkler ({sprinkler.Id}) đã align");
                            }
                        }
                    }
                }
                // TRƯỜNG HỢP 2: Pap → Fitting (trực tiếp)
                else if (fitting != null)
                {
                    details.Add($"TH2: Fitting ({fitting.Id})");
                    LogHelper.Log($"[ALIGN_CHAIN] Trường hợp 2: Align Fitting {fitting.Id} trực tiếp với Pap");

                    bool fittingAligned = IsAligned3D(pap, fitting, tolerance: 3.0 / 304.8);
                    if (!fittingAligned)
                    {
                        bool success = AlignMoveConnectWithPap(doc, pap, fitting);
                        if (success)
                        {
                            alignedCount++;
                            details.Add("✓ Fitting aligned");
                            LogHelper.Log("[ALIGN_CHAIN] ✓ Fitting đã được align");
                        }
                    }
                    else
                    {
                        details.Add("Fitting đã align");
                    }
                    
                    // Align sprinkler với fitting (nếu có)
                    if (sprinkler != null)
                    {
                        LogHelper.Log($"[ALIGN_CHAIN] Kiểm tra Sprinkler {sprinkler.Id}");
                        bool sprinklerAligned = IsAligned3D(fitting, sprinkler, tolerance: 3.0 / 304.8);
                        if (!sprinklerAligned)
                        {
                            bool success = AlignMoveConnectWithPap(doc, fitting, sprinkler);
                            if (success)
                            {
                                alignedCount++;
                                details.Add($"✓ Sprinkler ({sprinkler.Id}) aligned");
                                LogHelper.Log("[ALIGN_CHAIN] ✓ Sprinkler đã được align với Fitting");
                            }
                        }
                        else
                        {
                            details.Add($"Sprinkler ({sprinkler.Id}) đã align");
                        }
                    }
                }

                LogHelper.Log($"[ALIGN_CHAIN] === Hoàn thành: {alignedCount} element aligned ===");
            }
            catch (Exception ex)
            {
                LogHelper.Log($"[ALIGN_CHAIN] Lỗi: {ex.Message}");
                details.Add($"Lỗi: {ex.Message}");
            }

            return (alignedCount, string.Join(", ", details));
        }

        /// <summary>
        /// Lấy vị trí connector của Pap kết nối với ống
        /// </summary>
        private static XYZ GetPapConnectorOrigin(Element pap, Pipe pipe)
        {
            try
            {
                ConnectorManager papCM = GetConnectorManager(pap);
                if (papCM == null) return null;
                
                foreach (Connector papConn in papCM.Connectors)
                {
                    if (papConn.IsConnected)
                    {
                        foreach (Connector connectedConn in papConn.AllRefs)
                        {
                            if (connectedConn.Owner.Id == pipe.Id)
                            {
                                return papConn.Origin;
                            }
                        }
                    }
                }
                
                return null;
            }
            catch
            {
                return null;
            }
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

                // Bước 2: Tìm ống chính (ống được khoan lỗ) để lấy tâm quay
                Pipe mainPipe = FindMainPipe(pap);
                XYZ rotationCenter;
                
                if (mainPipe != null)
                {
                    // Tính tâm quay = giao điểm giữa center line ống chính và vector của Pap
                    rotationCenter = GetRotationCenter(pap, mainPipe);
                }
                else
                {
                    // Không có ống chính (Loại 1: Pap → Sprinkler trực tiếp)
                    // Dùng vị trí connector hướng lên của Pap làm tâm quay
                    Connector upConnector = GetUpwardConnector(pap);
                    if (upConnector != null)
                    {
                        rotationCenter = upConnector.Origin;
                    }
                    else
                    {
                        rotationCenter = GetElementLocation(pap);
                    }
                }
                
                if (rotationCenter == null)
                {
                    result.ErrorMessage = "Không thể tính được tâm quay";
                    return result;
                }

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

                            int dimsDeleted;
                            bool rotated = RotatePapToVertical(doc, pap, rotationCenter, mainPipe, out dimsDeleted);
                            result.DimensionsDeleted = dimsDeleted;
                            
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
        /// Lấy đường kính của pipe (mm)
        /// </summary>
        public static double GetPipeDiameterMM(Pipe pipe)
        {
            if (pipe == null) return 0;
            
            double diameterFeet = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM)?.AsDouble() ?? 0;
            return diameterFeet * 304.8; // Convert feet to mm
        }

        /// <summary>
        /// Tìm element kết nối trực tiếp với một element qua connector hướng xuống
        /// </summary>
        public static Element GetDownwardConnectedElement(Element element)
        {
            Connector downConnector = GetDownwardConnector(element);
            if (downConnector == null || !downConnector.IsConnected) return null;

            foreach (Connector connectedC in downConnector.AllRefs)
            {
                if (connectedC.Owner.Id != element.Id)
                {
                    return connectedC.Owner;
                }
            }
            return null;
        }

        /// <summary>
        /// Kiểm tra element có phải là Reducing/Coupling không
        /// </summary>
        public static bool IsReducingOrCoupling(Element element)
        {
            if (element is FamilyInstance fi)
            {
                // Kiểm tra category là Pipe Fitting
                if (fi.Category?.Id.IntegerValue == (int)BuiltInCategory.OST_PipeFitting)
                {
                    string familyName = fi.Symbol?.Family?.Name?.ToLower() ?? "";
                    string typeName = fi.Name?.ToLower() ?? "";
                    
                    return familyName.Contains("reducing") || 
                           familyName.Contains("coupling") || 
                           familyName.Contains("bushing") ||
                           typeName.Contains("reducing") ||
                           typeName.Contains("coupling") ||
                           typeName.Contains("bushing");
                }
            }
            return false;
        }

        /// <summary>
        /// Kiểm tra element có phải là Sprinkler không
        /// </summary>
        public static bool IsSprinkler(Element element)
        {
            if (element is FamilyInstance fi)
            {
                return fi.Category?.Id.IntegerValue == (int)BuiltInCategory.OST_Sprinklers;
            }
            return false;
        }

        /// <summary>
        /// Căn chỉnh element2 để thẳng hàng với element1 (thẳng đứng theo Z)
        /// Disconnect → Align/Move → Connect
        /// </summary>
        public static bool AlignTwoElements(Document doc, Element element1, Element element2, Connector conn1, Connector conn2)
        {
            try
            {
                // Kiểm tra đã thẳng hàng chưa
                if (IsAlignedVertically(conn1.Origin, conn2.Origin))
                {
                    return true; // Đã thẳng hàng
                }

                // 1. Disconnect
                ConnectionHelper.UnpinElementIfPinned(doc, element2);
                ConnectionHelper.DisconnectTwoElements(doc, element1, element2);

                // 2. Tính vector di chuyển (chỉ X-Y, giữ nguyên Z)
                double deltaX = conn1.Origin.X - conn2.Origin.X;
                double deltaY = conn1.Origin.Y - conn2.Origin.Y;
                XYZ moveVector = new XYZ(deltaX, deltaY, 0);

                // 3. Di chuyển element2 và tất cả element kết nối phía dưới nó
                var chainElements = GetAllConnectedElementsChain(element2);
                
                // Di chuyển element2
                ElementTransformUtils.MoveElement(doc, element2.Id, moveVector);
                
                // Di chuyển các element trong chuỗi (nhưng không di chuyển element1)
                foreach (var chainElement in chainElements)
                {
                    if (chainElement.Id != element1.Id)
                    {
                        ConnectionHelper.UnpinElementIfPinned(doc, chainElement);
                        ElementTransformUtils.MoveElement(doc, chainElement.Id, moveVector);
                    }
                }

                // 4. Connect lại
                // Lấy lại connector sau khi di chuyển
                var conn2Manager = GetConnectorManager(element2);
                if (conn2Manager != null)
                {
                    Connector closestConn = null;
                    double minDist = double.MaxValue;
                    
                    foreach (Connector c in conn2Manager.Connectors)
                    {
                        if (!c.IsConnected)
                        {
                            double dist = c.Origin.DistanceTo(conn1.Origin);
                            if (dist < minDist)
                            {
                                minDist = dist;
                                closestConn = c;
                            }
                        }
                    }

                    if (closestConn != null && minDist < 0.01)
                    {
                        // Move nhỏ để khớp chính xác
                        XYZ finalMove = conn1.Origin - closestConn.Origin;
                        ElementTransformUtils.MoveElement(doc, element2.Id, finalMove);
                        foreach (var chainElement in chainElements)
                        {
                            if (chainElement.Id != element1.Id)
                            {
                                ElementTransformUtils.MoveElement(doc, chainElement.Id, finalMove);
                            }
                        }

                        // Lấy lại connector và connect
                        conn2Manager = GetConnectorManager(element2);
                        foreach (Connector c in conn2Manager.Connectors)
                        {
                            if (!c.IsConnected && c.Origin.DistanceTo(conn1.Origin) < 0.001)
                            {
                                c.ConnectTo(conn1);
                                return true;
                            }
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
        /// XỬ LÝ CHÍNH: Căn chỉnh toàn bộ chuỗi Pap → Pipe/Reducing → Sprinkler
        /// Theo logic:
        /// 1. Tìm ống 65mm kết nối với Pap → Tâm quay
        /// 2. Nếu có Pipe 40mm: kiểm tra vector Pap-Pipe, căn chỉnh nếu cần
        /// 3. Kiểm tra vector với Reducing, căn chỉnh nếu cần
        /// 4. Kiểm tra vector với Sprinkler, căn chỉnh nếu cần
        /// </summary>
        public static AlignmentResult AlignChainFromPap(Document doc, Element pap)
        {
            var result = new AlignmentResult();
            LogHelper.Log("========== AlignChainFromPap START ==========");
            LogHelper.Log($"Pap Id: {pap.Id}, Name: {pap.Name}");

            try
            {
                // Bước 1: Tìm ống 65mm (ống chính) để xác định tâm quay
                Pipe mainPipe = FindPipeByDiameter(pap, 65.0, 5.0); // 65mm ± 5mm
                XYZ rotationCenter = null;
                
                LogHelper.Log($"[Step 1] Tìm ống 65mm: {(mainPipe != null ? $"Found (Id: {mainPipe.Id}, Diameter: {GetPipeDiameterMM(mainPipe):F1}mm)" : "NOT FOUND")}");

                if (mainPipe != null)
                {
                    // Tìm giao điểm center line ống 65 với vector Pap
                    rotationCenter = GetRotationCenter(pap, mainPipe);
                    LogHelper.Log($"[Step 1] Rotation Center: {(rotationCenter != null ? $"({rotationCenter.X:F3}, {rotationCenter.Y:F3}, {rotationCenter.Z:F3})" : "NULL")}");
                }
                else
                {
                    // Không có ống 65mm, dùng connector hướng lên của Pap
                    Connector upConn = GetUpwardConnector(pap);
                    rotationCenter = upConn?.Origin ?? GetElementLocation(pap);
                }

                // Bước 2: Lấy element kết nối phía dưới Pap
                Element nextElement = GetDownwardConnectedElement(pap);
                LogHelper.Log($"[Step 2] Element phía dưới Pap: {(nextElement != null ? $"{nextElement.GetType().Name} (Id: {nextElement.Id}, Name: {nextElement.Name})" : "NULL")}");
                
                if (nextElement == null)
                {
                    LogHelper.Log("[Step 2] ERROR: Không tìm thấy element kết nối phía dưới Pap");
                    result.ErrorMessage = "Không tìm thấy element kết nối phía dưới Pap";
                    return result;
                }

                // Lấy connectors kết nối giữa Pap và nextElement
                var (papConn, nextConn) = GetConnectingConnectors(pap, nextElement);
                LogHelper.Log($"[Step 2] Connectors - Pap: {(papConn != null ? $"({papConn.Origin.X:F3}, {papConn.Origin.Y:F3}, {papConn.Origin.Z:F3})" : "NULL")}");
                LogHelper.Log($"[Step 2] Connectors - Next: {(nextConn != null ? $"({nextConn.Origin.X:F3}, {nextConn.Origin.Y:F3}, {nextConn.Origin.Z:F3})" : "NULL")}");
                
                if (papConn != null && nextConn != null)
                {
                    bool aligned = IsAlignedVertically(papConn.Origin, nextConn.Origin);
                    LogHelper.Log($"[Step 2] IsAlignedVertically: {aligned}");
                }

                // Xác định loại element tiếp theo
                if (nextElement is Pipe pipe40)
                {
                    double diameter = GetPipeDiameterMM(pipe40);
                    LogHelper.Log($"[Step 2] Pipe detected - Diameter: {diameter:F1}mm");
                    
                    if (Math.Abs(diameter - 40.0) < 5.0) // Pipe 40mm
                    {
                        LogHelper.Log("[CASE 1] Pap → Pipe 40mm → Reducing → Sprinkler");
                        
                        // Kiểm tra và căn chỉnh Pap với Pipe 40
                        if (papConn != null && nextConn != null && !IsAlignedVertically(papConn.Origin, nextConn.Origin))
                        {
                            LogHelper.Log("[CASE 1] Căn chỉnh Pap với Pipe 40...");
                            result.ElementsNeedAlignment.Add(pipe40);
                            bool aligned = AlignTwoElements(doc, pap, pipe40, papConn, nextConn);
                            LogHelper.Log($"[CASE 1] Kết quả căn chỉnh Pap-Pipe40: {aligned}");
                            if (aligned)
                            {
                                result.ElementsAligned.Add(pipe40);
                            }
                        }

                        // Tiếp tục với Reducing
                        Element reducing = GetDownwardConnectedElement(pipe40);
                        LogHelper.Log($"[CASE 1] Reducing: {(reducing != null ? $"Id: {reducing.Id}, IsReducing: {IsReducingOrCoupling(reducing)}" : "NULL")}");
                        
                        if (reducing != null && IsReducingOrCoupling(reducing))
                        {
                            var (pipeConn, reducingConn) = GetConnectingConnectors(pipe40, reducing);
                            if (pipeConn != null && reducingConn != null && !IsAlignedVertically(pipeConn.Origin, reducingConn.Origin))
                            {
                                LogHelper.Log("[CASE 1] Căn chỉnh Pipe40 với Reducing...");
                                result.ElementsNeedAlignment.Add(reducing);
                                bool aligned = AlignTwoElements(doc, pipe40, reducing, pipeConn, reducingConn);
                                LogHelper.Log($"[CASE 1] Kết quả căn chỉnh Pipe40-Reducing: {aligned}");
                                if (aligned)
                                {
                                    result.ElementsAligned.Add(reducing);
                                }
                            }

                            // Tiếp tục với Sprinkler
                            Element sprinkler = GetDownwardConnectedElement(reducing);
                            LogHelper.Log($"[CASE 1] Sprinkler: {(sprinkler != null ? $"Id: {sprinkler.Id}, IsSprinkler: {IsSprinkler(sprinkler)}" : "NULL")}");
                            
                            if (sprinkler != null && IsSprinkler(sprinkler))
                            {
                                var (redConn, sprConn) = GetConnectingConnectors(reducing, sprinkler);
                                if (redConn != null && sprConn != null && !IsAlignedVertically(redConn.Origin, sprConn.Origin))
                                {
                                    LogHelper.Log("[CASE 1] Căn chỉnh Reducing với Sprinkler...");
                                    result.ElementsNeedAlignment.Add(sprinkler);
                                    bool aligned = AlignTwoElements(doc, reducing, sprinkler, redConn, sprConn);
                                    LogHelper.Log($"[CASE 1] Kết quả căn chỉnh Reducing-Sprinkler: {aligned}");
                                    if (aligned)
                                    {
                                        result.ElementsAligned.Add(sprinkler);
                                    }
                                }
                            }
                        }
                    }
                }
                else if (IsReducingOrCoupling(nextElement))
                {
                    // TRƯỜNG HỢP 2: Pap → Reducing → Sprinkler (không có Pipe 40)
                    LogHelper.Log("[CASE 2] Pap → Reducing → Sprinkler (không có Pipe 40)");
                    
                    // Kiểm tra và căn chỉnh Pap với Reducing
                    if (papConn != null && nextConn != null && !IsAlignedVertically(papConn.Origin, nextConn.Origin))
                    {
                        LogHelper.Log("[CASE 2] Căn chỉnh Pap với Reducing...");
                        result.ElementsNeedAlignment.Add(nextElement);
                        bool aligned = AlignTwoElements(doc, pap, nextElement, papConn, nextConn);
                        LogHelper.Log($"[CASE 2] Kết quả căn chỉnh Pap-Reducing: {aligned}");
                        if (aligned)
                        {
                            result.ElementsAligned.Add(nextElement);
                        }
                    }

                    // Tiếp tục với Sprinkler
                    Element sprinkler = GetDownwardConnectedElement(nextElement);
                    LogHelper.Log($"[CASE 2] Sprinkler: {(sprinkler != null ? $"Id: {sprinkler.Id}" : "NULL")}");
                    
                    if (sprinkler != null && IsSprinkler(sprinkler))
                    {
                        var (redConn, sprConn) = GetConnectingConnectors(nextElement, sprinkler);
                        if (redConn != null && sprConn != null && !IsAlignedVertically(redConn.Origin, sprConn.Origin))
                        {
                            LogHelper.Log("[CASE 2] Căn chỉnh Reducing với Sprinkler...");
                            result.ElementsNeedAlignment.Add(sprinkler);
                            bool aligned = AlignTwoElements(doc, nextElement, sprinkler, redConn, sprConn);
                            LogHelper.Log($"[CASE 2] Kết quả căn chỉnh Reducing-Sprinkler: {aligned}");
                            if (aligned)
                            {
                                result.ElementsAligned.Add(sprinkler);
                            }
                        }
                    }
                }
                else if (IsSprinkler(nextElement))
                {
                    // TRƯỜNG HỢP 3: Pap → Sprinkler trực tiếp
                    LogHelper.Log("[CASE 3] Pap → Sprinkler trực tiếp");
                    
                    if (papConn != null && nextConn != null && !IsAlignedVertically(papConn.Origin, nextConn.Origin))
                    {
                        LogHelper.Log("[CASE 3] Căn chỉnh Pap với Sprinkler...");
                        result.ElementsNeedAlignment.Add(nextElement);
                        bool aligned = AlignTwoElements(doc, pap, nextElement, papConn, nextConn);
                        LogHelper.Log($"[CASE 3] Kết quả căn chỉnh Pap-Sprinkler: {aligned}");
                        if (aligned)
                        {
                            result.ElementsAligned.Add(nextElement);
                        }
                    }
                }
                else
                {
                    LogHelper.Log($"[WARNING] Element không xác định loại: {nextElement.GetType().Name}");
                }

                // Bước 3: Quay cả cụm quanh tâm quay nếu cần (để vector Pap-Sprinkler thẳng đứng)
                LogHelper.Log("[Step 3] Kiểm tra xoay cụm...");
                LogHelper.Log($"[Step 3] rotationCenter: {(rotationCenter != null ? "OK" : "NULL")}, mainPipe: {(mainPipe != null ? "OK" : "NULL")}");
                
                if (rotationCenter != null && mainPipe != null)
                {
                    Element sprinkler = FindSprinklerInChain(pap, doc);
                    LogHelper.Log($"[Step 3] Sprinkler trong chuỗi: {(sprinkler != null ? $"Id: {sprinkler.Id}" : "NULL")}");
                    
                    if (sprinkler != null)
                    {
                        XYZ papPos = GetElementLocation(pap);
                        XYZ sprPos = GetElementLocation(sprinkler);
                        LogHelper.Log($"[Step 3] Pap Position: ({papPos?.X:F3}, {papPos?.Y:F3}, {papPos?.Z:F3})");
                        LogHelper.Log($"[Step 3] Sprinkler Position: ({sprPos?.X:F3}, {sprPos?.Y:F3}, {sprPos?.Z:F3})");

                        if (papPos != null && sprPos != null && !IsParallelToZAxis(papPos, sprPos))
                        {
                            double angle = GetAngleWithZAxis(papPos, sprPos);
                            double angleDegrees = angle * 180.0 / Math.PI;
                            LogHelper.Log($"[Step 3] Góc với trục Z: {angleDegrees:F2}°");

                            if (angleDegrees > 0.1)
                            {
                                LogHelper.Log($"[Step 3] Cần xoay {angleDegrees:F2}° để thẳng đứng");
                                result.RotationApplied = true;
                                result.RotationAngle = angleDegrees;

                                int dimsDeleted;
                                RotatePapToVertical(doc, pap, rotationCenter, mainPipe, out dimsDeleted);
                                result.DimensionsDeleted += dimsDeleted;
                            }
                        }
                    }
                }

                result.Success = true;
                if (result.ElementsNeedAlignment.Count == 0)
                {
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
        /// Tìm Pipe có đường kính cụ thể kết nối với element
        /// </summary>
        public static Pipe FindPipeByDiameter(Element element, double diameterMM, double toleranceMM = 5.0)
        {
            var connectedElements = GetConnectedElements(element);
            
            foreach (var connected in connectedElements)
            {
                if (connected is Pipe pipe)
                {
                    double pipeDiameter = GetPipeDiameterMM(pipe);
                    if (Math.Abs(pipeDiameter - diameterMM) <= toleranceMM)
                    {
                        return pipe;
                    }
                }
            }
            
            return null;
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
        public int DimensionsDeleted { get; set; } = 0; // Số dimensions đã xóa
        public List<Element> ElementsNeedAlignment { get; set; } = new List<Element>();
        public List<Element> ElementsAligned { get; set; } = new List<Element>();
        public string ErrorMessage { get; set; } = "";
    }
}