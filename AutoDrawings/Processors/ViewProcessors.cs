using InvAddIn.Drawings;
using Inventor;
using System.Collections.Generic;
using System.Linq;

namespace InvAddIn.Processors {
    internal static class SheetProcessors {
        internal static void CreateViews(SheetParameters parameters, Sheet sheet) {
            var baseViewParameters = parameters.ViewsParameters[0];
            var baseview = sheet.DrawingViews.AddBaseView(parameters.Model, baseViewParameters.Point, 1.0, baseViewParameters.Orientation, baseViewParameters.Style);
            baseview.ScaleString = baseViewParameters.ScaleString;

            if (parameters.ViewsParameters.Count == 1) return;

            for (int i = 1; i < parameters.ViewsParameters.Count; i++) {
                var view_p = parameters.ViewsParameters[i];
                //strange bug with Point2d Position parameter. Works fine if TransitionGeometry.CreatePoint2d inplace
                //but E_INVALIDARG if make it in another place
                sheet.DrawingViews.AddProjectedView(baseview, view_p.Point.Copy(), view_p.Style);
            }
        }


        internal static void ChangeTitleBlock(SheetParameters parameters, Sheet sheet) {
            sheet.TitleBlock.Delete();
            sheet.AddTitleBlock(parameters.TitleBlock);
        }


        internal static void AddPartsList(SheetParameters parameters, Sheet sheet) {
            sheet.Activate();
            if (parameters.Model.DocumentType != DocumentTypeEnum.kAssemblyDocumentObject) return;

            var x = 2.0;
            var y = parameters.SheetFormat.Name == "А4" ? 4.5 : 0.5;
            var oPlacementPoint = CAddIn.App.TransientGeometry.CreatePoint2d(x, y);
            PartsList oPartsList;
            try {
                oPartsList = sheet.PartsLists.Add(
                    ViewOrModel: parameters.Model,
                    PlacementPoint: oPlacementPoint,
                    Level: PartsListLevelEnum.kStructuredAllLevels,
                    NumberingScheme: null,
                    WrapLeft: false);
            } catch {
                oPartsList = sheet.PartsLists.Add(
                    ViewOrModel: parameters.Model,
                    PlacementPoint: oPlacementPoint,
                    Level: PartsListLevelEnum.kPartsOnly,
                    NumberingScheme: null,
                    WrapLeft: false);
            }
            var vSize = oPartsList.RangeBox.MaxPoint.Y - oPartsList.RangeBox.MinPoint.Y;
            oPartsList.Position = CAddIn.App.TransientGeometry.CreatePoint2d(x, y + vSize);
        }


        internal static void AddDimensions(SheetParameters parameters, Sheet sheet) {
            void addDimensionToView(DrawingView dv, BaselineDimensionSets d) {
                var curves = dv.DrawingCurves;
                var intentCollection = CAddIn.App.TransientObjects.CreateObjectCollection();
                var tmpCollection = new List<DrawingCurve>();

                foreach (DrawingCurve curve in curves) {
                    switch (curve.ProjectedCurveType) {
                        case Curve2dTypeEnum.kUnknownCurve2d:
                            break;
                        case Curve2dTypeEnum.kLineCurve2d:
                        case Curve2dTypeEnum.kLineSegmentCurve2d:
                            tmpCollection.Add(curve);
                            break;
                        case Curve2dTypeEnum.kCircleCurve2d:
                            tmpCollection.Add(curve);
                            break;
                        case Curve2dTypeEnum.kCircularArcCurve2d:
                        case Curve2dTypeEnum.kEllipseFullCurve2d:
                        case Curve2dTypeEnum.kEllipticalArcCurve2d:
                            //tmpCollection.Add(curve);
                            break;
                        case Curve2dTypeEnum.kBSplineCurve2d:
                            break;
                        case Curve2dTypeEnum.kPolylineCurve2d:
                            break;
                        default:
                            break;
                    }
                }

                var sorted = tmpCollection
                    .OrderBy(p => p.StartPoint != null ? p.StartPoint.X : p.CenterPoint.X)
                    .ThenByDescending(p => p.StartPoint != null ? p.StartPoint.Y : p.CenterPoint.Y);

                foreach (var item in sorted) {
                    intentCollection.Add(sheet.CreateGeometryIntent(item));
                }

                var hPoint = CAddIn.App.TransientGeometry.CreatePoint2d(dv.Center.X, dv.Top + 1);
                var vPoint = CAddIn.App.TransientGeometry.CreatePoint2d(dv.Left - 1, dv.Center.Y);

                d.Add(intentCollection, hPoint, DimensionTypeEnum.kHorizontalDimensionType);
                d.Add(intentCollection, vPoint, DimensionTypeEnum.kVerticalDimensionType);
            }

            if (parameters.Model.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject) return;
            var dimensions = sheet.DrawingDimensions.BaselineDimensionSets;
            foreach (DrawingView view in sheet.DrawingViews)
                addDimensionToView(view, dimensions);
        }


        internal static void AddCenterlines(SheetParameters parameters, Sheet sheet) {
            var doc = parameters.DrawDoc;
            var centerlineSettings = doc.DrawingSettings.AutomatedCenterlineSettings;
            centerlineSettings.ApplyToHoles = true;
            centerlineSettings.ApplyToCylinders = false;

            foreach (DrawingView view in sheet.DrawingViews) {
                view.SetAutomatedCenterlineSettings(centerlineSettings);
            }

        }
    }
}
