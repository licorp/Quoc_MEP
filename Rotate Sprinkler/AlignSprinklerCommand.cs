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
                // Bước 1: Chọn Pap (Pipe Accessory Point) làm chuẩn
                Reference papRef = uidoc.Selection.PickObject(
                    ObjectType.Element,
                    new PapSelectionFilter(),
                    "Chọn Pap (Pipe Fitting/Accessory) làm chuẩn:");

                if (papRef == null)
                {
                    message = "Không chọn được Pap";
                    return Result.Cancelled;
                }

                Element pap = doc.GetElement(papRef);

                // Bước 2: Kiểm tra và căn chỉnh
                using (Transaction trans = new Transaction(doc, "Căn chỉnh Sprinkler thẳng hàng với Pap"))
                {
                    trans.Start();

                    try
                    {
                        // Thực hiện căn chỉnh XY + Quay song song Z
                        AlignmentResult result = SprinklerAlignmentHelper.AlignAndRotateToZ(doc, pap);

                        if (result.Success)
                        {
                            trans.Commit();

                            if (result.AlreadyAligned && !result.RotationApplied)
                            {
                                TaskDialog.Show("Kết quả", "Các element đã thẳng hàng với Pap và song song trục Z, không cần căn chỉnh.");
                            }
                            else
                            {
                                string msg = $"Đã căn chỉnh thành công {result.ElementsAligned.Count} element.";
                                if (result.RotationApplied)
                                {
                                    msg += $"\nĐã quay cụm element {result.RotationAngle:F2}° để song song với trục Z.";
                                }
                                TaskDialog.Show("Kết quả", msg);
                            }

                            return Result.Succeeded;
                        }
                        else
                        {
                            trans.RollBack();
                            message = !string.IsNullOrEmpty(result.ErrorMessage) 
                                ? result.ErrorMessage 
                                : "Không thể căn chỉnh. Vui lòng kiểm tra lại các kết nối.";
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
