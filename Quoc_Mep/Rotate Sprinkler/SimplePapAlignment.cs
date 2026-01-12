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
        /// </summary>
        public static AlignmentResult AlignPapSimple(Document doc, Element pap)
        {
            var result = new AlignmentResult();
            
            try
            {
                Debug.WriteLine($"[SIMPLE] Bắt đầu xử lý Pap {pap.Id}");
                
                // Bước 1: Tìm ống 65mm kết nối với Pap
                Pipe pipe65 = FindPipe65Connected(pap);
                if (pipe65 == null)
                {
                    result.ErrorMessage = "Không tìm thấy ống 65mm kết nối với Pap";
                    Debug.WriteLine($"[SIMPLE] {result.ErrorMessage}");
                    return result;
                }
                
                Debug.WriteLine($"[SIMPLE] Tìm thấy ống 65: {pipe65.Id}");
                
                // Bước 2: Lấy vị trí giao điểm giữa Pap và ống 65 (dùng làm tâm quay)
                XYZ rotationCenter = GetPapPipeIntersection(pap, pipe65);
                if (rotationCenter == null)
                {
                    result.ErrorMessage = "Không xác định được giao điểm giữa Pap và ống 65";
                    Debug.WriteLine($"[SIMPLE] {result.ErrorMessage}");
                    return result;
                }
                
                Debug.WriteLine($"[SIMPLE] Tâm quay: ({rotationCenter.X:F3}, {rotationCenter.Y:F3}, {rotationCenter.Z:F3})");
                
                // Bước 3: Lấy hướng của Pap
                XYZ papDirection = GetPapDirection(pap);
                if (papDirection == null || papDirection.IsZeroLength())
                {
                    result.ErrorMessage = "Không xác định được hướng Pap";
                    Debug.WriteLine($"[SIMPLE] {result.ErrorMessage}");
                    return result;
                }
                
                Debug.WriteLine($"[SIMPLE] Hướng Pap: ({papDirection.X:F3}, {papDirection.Y:F3}, {papDirection.Z:F3})");
                
                // Bước 4: Kiểm tra góc với trục Z (thẳng đứng)
                XYZ zAxis = XYZ.BasisZ;
                double dotProduct = Math.Abs(papDirection.DotProduct(zAxis));
                double angleDegrees = Math.Acos(Math.Max(-1.0, Math.Min(1.0, dotProduct))) * 180.0 / Math.PI;
                
                Debug.WriteLine($"[SIMPLE] Góc lệch với trục Z: {angleDegrees:F2}°");
                
                // Nếu đã thẳng đứng (góc < 0.1 độ), không cần quay
                if (angleDegrees < 0.1)
                {
                    result.Success = true;
                    result.AlreadyAligned = true;
                    Debug.WriteLine("[SIMPLE] Pap đã thẳng đứng");
                    return result;
                }
                
                // Bước 5: Quay Pap để thẳng đứng
                int dimsDeleted;
                bool rotated = RotatePapToVertical(doc, pap, rotationCenter, pipe65, out dimsDeleted);
                
                if (rotated)
                {
                    result.Success = true;
                    result.RotationApplied = true;
                    result.RotationAngle = angleDegrees;
                    result.DimensionsDeleted = dimsDeleted;
                    result.ElementsAligned.Add(pap);
                    Debug.WriteLine($"[SIMPLE] Đã quay Pap thành công, xóa {dimsDeleted} dimensions");
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
        /// Lấy giao điểm giữa Pap và ống 65 (dùng làm tâm quay)
        /// </summary>
        private static XYZ GetPapPipeIntersection(Element pap, Pipe pipe)
        {
            try
            {
                // Lấy connector của Pap kết nối với ống
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
                                // Trả về vị trí connector của Pap (giao điểm với ống)
                                return papConn.Origin;
                            }
                        }
                    }
                }
                
                // Fallback: dùng vị trí của Pap
                return GetElementLocation(pap);
            }
            catch
            {
                return null;
            }
        }
    }
}
