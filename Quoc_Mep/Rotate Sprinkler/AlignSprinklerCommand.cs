using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace Quoc_MEP
{
    /// <summary>
    /// Command để căn chỉnh Sprinkler và Pipe thẳng hàng với Pap
    /// Pap là đối tượng cố định làm chuẩn
    /// Kiểm tra theo trục Z (vertical) - X-Y phải khớp nhau
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class AlignSprinklerCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                // Bước 1: Chọn nhiều Pap (Pipe Accessory Point)
                IList<Reference> papRefs = uidoc.Selection.PickObjects(
                    ObjectType.Element,
                    new PapSelectionFilter(),
                    "Chọn các cụm Pap-Pipe-Sprinkler cần căn chỉnh (có thể chọn nhiều):");

                if (papRefs == null || papRefs.Count == 0)
                {
                    message = "Không chọn được Pap nào";
                    return Result.Cancelled;
                }

                List<Element> paps = new List<Element>();
                foreach (Reference papRef in papRefs)
                {
                    Element pap = doc.GetElement(papRef);
                    if (pap != null)
                    {
                        paps.Add(pap);
                    }
                }

                // Bước 2: Kiểm tra và căn chỉnh từng cụm
                using (Transaction trans = new Transaction(doc, "Căn chỉnh Sprinkler thẳng hàng với Pap"))
                {
                    trans.Start();

                    try
                    {
                        int totalSuccess = 0;
                        int totalFailed = 0;
                        int totalAlreadyAligned = 0;
                        int totalRotated = 0;
                        int totalDimensionsDeleted = 0;
                        List<string> errors = new List<string>();
                        List<string> rotationDetails = new List<string>();

                        // XỬ LÝ TUẦN TỰ TỪNG PAP MỘT (không xử lý hàng loạt)
                        // Flow cho MỖI Pap: Xoay Pap → Tìm chain → Align chain → Xong Pap này → Sang Pap khác
                        foreach (Element pap in paps)
                        {
                            Debug.WriteLine($"\n========== BẮT ĐẦU XỬ LÝ PAP {pap.Id} ==========");
                            
                            // Xử lý RIÊNG LẺ Pap này: Xoay + Tìm + Align trong 1 lần gọi
                            // Không tách rời: tìm tất cả trước rồi mới align sau
                            AlignmentResult result = SprinklerAlignmentHelper.AlignPapSimple(doc, pap);

                            if (result.Success)
                            {
                                totalSuccess++;
                                totalDimensionsDeleted += result.DimensionsDeleted;
                                if (result.AlreadyAligned && !result.RotationApplied)
                                {
                                    totalAlreadyAligned++;
                                }
                                if (result.RotationApplied)
                                {
                                    totalRotated++;
                                    rotationDetails.Add($"Pap {pap.Id}: {result.RotationAngle:F2}°");
                                }
                                
                                // Log chi tiết từ result.ErrorMessage (có thông tin debug)
                                if (!string.IsNullOrEmpty(result.ErrorMessage))
                                {
                                    rotationDetails.Add($"Pap {pap.Id}: {result.ErrorMessage}");
                                }
                                
                                Debug.WriteLine($"========== HOÀN THÀNH PAP {pap.Id} ==========\n");
                            }
                            else
                            {
                                totalFailed++;
                                if (!string.IsNullOrEmpty(result.ErrorMessage))
                                {
                                    errors.Add($"Pap {pap.Id}: {result.ErrorMessage}");
                                }
                                Debug.WriteLine($"========== THẤT BẠI PAP {pap.Id} ==========\n");
                            }
                        }

                        if (totalSuccess > 0)
                        {
                            trans.Commit();

                            string msg = $"Đã xử lý {paps.Count} Pap:\n";
                            msg += $"✓ Thành công: {totalSuccess}\n";
                            if (totalAlreadyAligned > 0)
                            {
                                msg += $"  - Đã thẳng đứng sẵn: {totalAlreadyAligned}\n";
                            }
                            if (totalRotated > 0)
                            {
                                msg += $"  - Đã quay để căn chỉnh: {totalRotated}\n";
                            }
                            if (rotationDetails.Count > 0 && rotationDetails.Count <= 5)
                            {
                                msg += "\nChi tiết:\n" + string.Join("\n", rotationDetails);
                            }
                            if (totalDimensionsDeleted > 0)
                            {
                                msg += $"\n⚠ Đã xóa {totalDimensionsDeleted} dimensions để tránh lỗi\n";
                            }
                            if (totalFailed > 0)
                            {
                                msg += $"✗ Thất bại: {totalFailed}";
                                if (errors.Count > 0)
                                {
                                    msg += "\n\nChi tiết lỗi:\n" + string.Join("\n", errors.Take(5));
                                    if (errors.Count > 5)
                                    {
                                        msg += $"\n... và {errors.Count - 5} lỗi khác";
                                    }
                                }
                            }
                            TaskDialog.Show("Kết quả", msg);

                            return Result.Succeeded;
                        }
                        else
                        {
                            trans.RollBack();
                            message = "Không thể căn chỉnh được cụm nào.\n";
                            if (errors.Count > 0)
                            {
                                message += string.Join("\n", errors.Take(3));
                            }
                            return Result.Failed;
                        }
                    }
                    catch (Exception ex)
                    {
                        trans.RollBack();
                        message = $"Lỗi khi căn chỉnh: {ex.Message}";
                        return Result.Failed;
                    }
                }
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return Result.Cancelled;
            }
            catch (Exception ex)
            {
                message = $"Lỗi không mong muốn: {ex.Message}";
                return Result.Failed;
            }
        }
    }

    /// <summary>
    /// Selection filter cho Pap (Pipe Fitting/Accessory)
    /// </summary>
    public class PapSelectionFilter : ISelectionFilter
    {
        public bool AllowElement(Element elem)
        {
            if (elem == null) return false;

            // Chấp nhận Pipe Fitting
            if (elem.Category != null)
            {
                var categoryId = elem.Category.Id.IntegerValue;

                // Pipe Fitting và Pipe Accessory
                if (categoryId == (int)BuiltInCategory.OST_PipeFitting ||
                    categoryId == (int)BuiltInCategory.OST_PipeAccessory)
                {
                    return true;
                }
            }

            // Chấp nhận FamilyInstance có MEPModel
            if (elem is FamilyInstance fi && fi.MEPModel != null)
            {
                var category = elem.Category;
                if (category != null)
                {
                    var categoryId = category.Id.IntegerValue;
                    if (categoryId == (int)BuiltInCategory.OST_PipeFitting ||
                        categoryId == (int)BuiltInCategory.OST_PipeAccessory)
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        public bool AllowReference(Reference reference, XYZ position)
        {
            return false;
        }
    }
}