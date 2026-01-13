using System;
using System.Diagnostics;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;

namespace Quoc_MEP
{
    /// <summary>
    /// Logic đơn giản: Chỉ kiểm tra Pap với ống 65mm
    /// Bỏ hết logic phức tạp về pipe chain và sprinkler
    /// </summary>
    public static partial class SprinklerAlignmentHelper
    {
        /// <summary>
        /// Căn chỉnh Pap đơn giản - chỉ cần Pap và ống 65mm
        /// QUAN TRỌNG: Xử lý TUẦN TỰ từng Pap:
        /// 1. Xoay Pap này
        /// 2. Tìm chain của Pap này
        /// 3. Align chain của Pap này
        /// 4. Xong → Trả về kết quả
        /// KHÔNG tách rời: tìm hết trước → xoay sau
        /// </summary>
        public static AlignmentResult AlignPapSimple(Document doc, Element pap)
        {
            var result = new AlignmentResult();
            
            try
            {
                Debug.WriteLine($"[SIMPLE] ╔═══════════════════════════════════════════╗");
                Debug.WriteLine($"[SIMPLE] ║  XỬ LÝ RIÊNG LẺ PAP {pap.Id}");
                Debug.WriteLine($"[SIMPLE] ╚═══════════════════════════════════════════╝");
                Debug.WriteLine($"[SIMPLE] Bước 1/4: Tìm ống 65mm kết nối với Pap");
                
                // Bước 1: Tìm ống 65mm kết nối với Pap
                Pipe pipe65 = FindPipe65Connected(pap);
                if (pipe65 == null)
                {
                    result.ErrorMessage = "Không tìm thấy ống 65mm kết nối với Pap";
                    Debug.WriteLine($"[SIMPLE] {result.ErrorMessage}");
                    return result;
                }
                
                Debug.WriteLine($"[SIMPLE] ✓ Tìm thấy ống 65: {pipe65.Id}");
                Debug.WriteLine($"[SIMPLE] Bước 2/4: Xác định điểm giao (tâm quay)");
                
                // Bước 2: Lấy vị trí giao điểm giữa Pap và ống 65 (dùng làm tâm quay)
                XYZ rotationCenter = GetPapPipeIntersection(pap, pipe65);
                if (rotationCenter == null)
                {
                    result.ErrorMessage = "Không xác định được giao điểm giữa Pap và ống 65";
                    Debug.WriteLine($"[SIMPLE] {result.ErrorMessage}");
                    return result;
                }
                
                Debug.WriteLine($"[SIMPLE] ✓ Tâm quay: ({rotationCenter.X:F3}, {rotationCenter.Y:F3}, {rotationCenter.Z:F3})");
                Debug.WriteLine($"[SIMPLE] Bước 3/4: Kiểm tra và xoay Pap (nếu cần)");
                
                // Bước 3: Lấy hướng của Pap
                XYZ papDirection = GetPapDirection(pap);
                if (papDirection == null || papDirection.IsZeroLength())
                {
                    result.ErrorMessage = "Không xác định được hướng Pap";
                    Debug.WriteLine($"[SIMPLE] {result.ErrorMessage}");
                    return result;
                }
                
                Debug.WriteLine($"[SIMPLE] Hướng Pap: ({papDirection.X:F3}, {papDirection.Y:F3}, {papDirection.Z:F3})");
                
                // Bước 4: Tính góc lệch CHÍNH XÁC của Pap này
                // Sử dụng điểm giao (rotationCenter) để tính góc chính xác từ tâm quay
                XYZ verticalDirection = -XYZ.BasisZ; // Hướng xuống
                
                // Kiểm tra góc giữa Pap direction và trục thẳng đứng hướng xuống
                // Không dùng Math.Abs để giữ thông tin hướng
                double dotProduct = papDirection.DotProduct(verticalDirection);
                double angleDegrees = Math.Acos(Math.Max(-1.0, Math.Min(1.0, Math.Abs(dotProduct)))) * 180.0 / Math.PI;
                
                Debug.WriteLine($"[SIMPLE] *** Pap {pap.Id} - Góc lệch CHÍNH XÁC: {angleDegrees:F4}° (dotProduct: {dotProduct:F4}) ***");
                Debug.WriteLine($"[SIMPLE] Điểm giao (tâm quay): ({rotationCenter.X:F3}, {rotationCenter.Y:F3}, {rotationCenter.Z:F3})");
                
                // Nếu đã thẳng đứng và hướng xuống (góc < 0.5 độ và dotProduct > 0.99), không cần quay
                if (angleDegrees < 0.5 && dotProduct > 0.99)
                {
                    result.Success = true;
                    result.AlreadyAligned = true;
                    result.ErrorMessage = $"Pap đã gần thẳng đứng (lệch {angleDegrees:F2}°)";
                    Debug.WriteLine($"[SIMPLE] Pap {pap.Id} đã thẳng đứng");
                    return result;
                }
                
                Debug.WriteLine($"[SIMPLE] Pap {pap.Id} CẦN XOAY {angleDegrees:F2}° để thẳng đứng");
                
                Debug.WriteLine($"[SIMPLE] >>> Đang xoay Pap {pap.Id}...");
                int dimsDeleted;
                bool rotated = RotatePapToVertical(doc, pap, rotationCenter, pipe65, out dimsDeleted);
                
                if (rotated)
                {
                    result.Success = true;
                    result.RotationApplied = true;
                    result.RotationAngle = angleDegrees;
                    result.DimensionsDeleted = dimsDeleted;
                    result.ElementsAligned.Add(pap);
                    Debug.WriteLine($"[SIMPLE] ✓ Đã quay Pap {pap.Id} thành công, xóa {dimsDeleted} dimensions");
                    
                    // Bước 6: SAU KHI XOAY XONG → Tìm và align chain NGAY CHO PAP NÀY
                    Debug.WriteLine($"[SIMPLE] Bước 4/4: Tìm và align chain sau khi xoay xong Pap {pap.Id}");
                    // Bước 6: Align toàn bộ chuỗi (Pipe 40mm + Fitting) với Pap
                    Debug.WriteLine($"[SIMPLE] === Align chuỗi kết nối với Pap {pap.Id} ===");
                    var (alignedCount, chainDetails) = SprinklerAlignmentHelper.AlignChainWithPap(doc, pap);
                    
                    string alignMsg = "";
                    if (alignedCount > 0)
                    {
                        alignMsg = $", đã align {alignedCount} element ({chainDetails})";
                        Debug.WriteLine($"[SIMPLE] ✓ Đã align {alignedCount} element trong chuỗi");
                    }
                    else if (!string.IsNullOrEmpty(chainDetails))
                    {
                        alignMsg = $", {chainDetails}";
                    }
                    
                    result.ErrorMessage = $"Đã xoay {angleDegrees:F2}°, xóa {dimsDeleted} dimensions{alignMsg}";
                }
                else
                {
                    result.ErrorMessage = "Không thể quay Pap";
                    Debug.WriteLine($"[SIMPLE] {result.ErrorMessage}");
                }
                
                return result;
            }
            catch (Exception ex)
            {
                result.ErrorMessage = $"Lỗi: {ex.Message}";
                Debug.WriteLine($"[SIMPLE] Exception: {ex.Message}\n{ex.StackTrace}");
                return result;
            }
        }
        
        /// <summary>
        /// Tìm ống 65mm kết nối với Pap
        /// </summary>
        private static Pipe FindPipe65Connected(Element pap)
        {
            try
            {
                ConnectorManager connectorManager = GetConnectorManager(pap);
                if (connectorManager == null) return null;
                
                foreach (Connector connector in connectorManager.Connectors)
                {
                    if (connector.IsConnected)
                    {
                        foreach (Connector connectedConnector in connector.AllRefs)
                        {
                            Element connectedElement = connectedConnector.Owner;
                            if (connectedElement is Pipe pipe && pipe.Id != pap.Id)
                            {
                                // Lấy đường kính ống (inches)
                                Parameter diamParam = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM);
                                if (diamParam != null)
                                {
                                    double diameterFeet = diamParam.AsDouble();
                                    double diameterMM = diameterFeet * 304.8; // Chuyển feet sang mm
                                    
                                    Debug.WriteLine($"[SIMPLE] Tìm thấy ống: {pipe.Id}, Diameter: {diameterMM:F0}mm");
                                    
                                    // Kiểm tra xem có phải ống 65mm không (cho phép sai số ±5mm)
                                    if (Math.Abs(diameterMM - 65) < 5)
                                    {
                                        return pipe;
                                    }
                                }
                            }
                        }
                    }
                }
                
                // Nếu không tìm thấy ống 65mm chính xác, trả về ống lớn nhất
                Pipe largestPipe = null;
                double largestDiameter = 0;
                
                foreach (Connector connector in connectorManager.Connectors)
                {
                    if (connector.IsConnected)
                    {
                        foreach (Connector connectedConnector in connector.AllRefs)
                        {
                            Element connectedElement = connectedConnector.Owner;
                            if (connectedElement is Pipe pipe && pipe.Id != pap.Id)
                            {
                                Parameter diamParam = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM);
                                if (diamParam != null)
                                {
                                    double diameterFeet = diamParam.AsDouble();
                                    if (diameterFeet > largestDiameter)
                                    {
                                        largestDiameter = diameterFeet;
                                        largestPipe = pipe;
                                    }
                                }
                            }
                        }
                    }
                }
                
                if (largestPipe != null)
                {
                    Debug.WriteLine($"[SIMPLE] Không tìm thấy ống 65mm, dùng ống lớn nhất: {largestPipe.Id}, Diameter: {largestDiameter * 304.8:F0}mm");
                }
                
                return largestPipe;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[SIMPLE] FindPipe65Connected Exception: {ex.Message}");
                return null;
            }
        }
        
        /// <summary>
        /// Lấy điểm trên center line của ống 65 gần connector Pap nhất (dùng làm tâm quay)
        /// Đây là điểm chiếu vuông góc của connector Pap lên center line ống 65
        /// </summary>
        private static XYZ GetPapPipeIntersection(Element pap, Pipe pipe)
        {
            try
            {
                // Lấy center line của ống 65
                LocationCurve locCurve = pipe.Location as LocationCurve;
                if (locCurve == null)
                {
                    Debug.WriteLine("[SIMPLE] Pipe không có LocationCurve");
                    return null;
                }
                
                Line pipeCenterLine = locCurve.Curve as Line;
                if (pipeCenterLine == null)
                {
                    // Nếu không phải Line, lấy 2 điểm đầu cuối
                    XYZ p1 = locCurve.Curve.GetEndPoint(0);
                    XYZ p2 = locCurve.Curve.GetEndPoint(1);
                    pipeCenterLine = Line.CreateBound(p1, p2);
                }
                
                // Lấy vị trí connector của Pap
                XYZ papConnectorOrigin = null;
                ConnectorManager papCM = GetConnectorManager(pap);
                if (papCM != null)
                {
                    foreach (Connector papConn in papCM.Connectors)
                    {
                        if (papConn.IsConnected)
                        {
                            foreach (Connector connectedConn in papConn.AllRefs)
                            {
                                if (connectedConn.Owner.Id == pipe.Id)
                                {
                                    papConnectorOrigin = papConn.Origin;
                                    break;
                                }
                            }
                        }
                        if (papConnectorOrigin != null) break;
                    }
                }
                
                if (papConnectorOrigin == null)
                {
                    Debug.WriteLine("[SIMPLE] Không tìm thấy connector của Pap kết nối với ống");
                    return GetElementLocation(pap);
                }
                
                // Tính CHÍNH XÁC điểm giao giữa trục Pap và center line ống 65
                // Phương pháp: chiếu connector Pap lên center line
                IntersectionResult projResult = pipeCenterLine.Project(papConnectorOrigin);
                if (projResult != null)
                {
                    XYZ projectedPoint = projResult.XYZPoint;
                    double distanceMM = projResult.Distance * 304.8;
                    
                    Debug.WriteLine($"[SIMPLE] === ĐIỂM GIAO CHÍNH XÁC ===");
                    Debug.WriteLine($"[SIMPLE] Vị trí connector Pap: ({papConnectorOrigin.X:F4}, {papConnectorOrigin.Y:F4}, {papConnectorOrigin.Z:F4})");
                    Debug.WriteLine($"[SIMPLE] Điểm chiếu lên center line: ({projectedPoint.X:F4}, {projectedPoint.Y:F4}, {projectedPoint.Z:F4})");
                    Debug.WriteLine($"[SIMPLE] Khoảng cách vuông góc: {distanceMM:F2}mm");
                    Debug.WriteLine($"[SIMPLE] ==============================");
                    
                    return projectedPoint;
                }
                
                // Fallback: dùng vị trí connector của Pap
                Debug.WriteLine("[SIMPLE] CẢNH BÁO: Không chiếu được lên center line, dùng vị trí connector");
                return papConnectorOrigin;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[SIMPLE] GetPapPipeIntersection Exception: {ex.Message}");
                return null;
            }
        }
    }
}
