using Inventor;
using System.Collections.Generic;

namespace InvAddIn.Drawings {
    internal class SheetParameters {
        public string PartNumber { get; }
        public string Description { get; }
        public SheetFormat SheetFormat { get; set; }
        public List<ViewParameters> ViewsParameters { get; private set; } = new List<ViewParameters>();
        public _Document Model { get; private set; }
        public DrawingDocument DrawDoc { get; private set; }
        public TitleBlockDefinition TitleBlock { get; set; }
        private TitleBlockDefinitions TitleBlocks { get; set; }


        public SheetParameters(ModelInfo modelParameters, DrawingDocument drawDoc) {
            Model = modelParameters.Definition.Document as _Document;
            DrawDoc = drawDoc;
            TitleBlocks = DrawDoc.TitleBlockDefinitions;
            TitleBlock = null;
            PartNumber = modelParameters.PartNumber;
            Description = modelParameters.Description;
        }

        public ViewParameters AddViewParameters() {
            var parameters = new ViewParameters();
            ViewsParameters.Add(parameters);
            return parameters;
        }

        public void AddBaseView(Point2d point) {
            var parameters = new ViewParameters(point: point);
            ViewsParameters.Add(parameters);
        }

        public void AddProjectedView(Point2d point) {
            var parameters = new ViewParameters(point: point);
            ViewsParameters.Add(parameters);
        }

        public void AddTitleBlock(string sheetFormat) {
            if (sheetFormat.Contains("А4")) TitleBlock = TitleBlocks[1];
            else if (sheetFormat.Contains("А3")) TitleBlock = TitleBlocks[2];
        }

        public static SheetParameters CreateA42Views(ModelInfo cdef, DrawingDocument drawDoc) {
            var parameters = new SheetParameters(cdef, drawDoc);
            parameters.SheetFormat = drawDoc.SheetFormats["А4"];
            parameters.AddBaseView(point: CAddIn.App.TransientGeometry.CreatePoint2d(11.5, 22.5));
            parameters.AddProjectedView(point: CAddIn.App.TransientGeometry.CreatePoint2d(11.5, 15.5));
            return parameters;
        }

        public static SheetParameters CreateA43Views(ModelInfo cdef, DrawingDocument drawDoc) {
            var parameters = new SheetParameters(cdef, drawDoc);
            parameters.SheetFormat = drawDoc.SheetFormats["А4"];
            parameters.AddBaseView(point: CAddIn.App.TransientGeometry.CreatePoint2d(8.5, 22.5));
            parameters.AddProjectedView(point: CAddIn.App.TransientGeometry.CreatePoint2d(13.5, 22.5));
            parameters.AddProjectedView(point: CAddIn.App.TransientGeometry.CreatePoint2d(8.5, 15.5));
            return parameters;
        }

        public static SheetParameters CreateA32Views(ModelInfo cdef, DrawingDocument drawDoc) {
            var parameters = new SheetParameters(cdef, drawDoc);
            parameters.SheetFormat = drawDoc.SheetFormats["А3"];
            parameters.AddBaseView(point: CAddIn.App.TransientGeometry.CreatePoint2d(11.5, 22.5));
            parameters.AddProjectedView(point: CAddIn.App.TransientGeometry.CreatePoint2d(11.5, 15.5));
            return parameters;
        }

        public static SheetParameters CreateA33Views(ModelInfo cdef, DrawingDocument drawDoc) {
            var parameters = new SheetParameters(cdef, drawDoc);
            parameters.SheetFormat = drawDoc.SheetFormats["А3"];
            parameters.AddBaseView(point: CAddIn.App.TransientGeometry.CreatePoint2d(8.5, 22.5));
            parameters.AddProjectedView(point: CAddIn.App.TransientGeometry.CreatePoint2d(18.5, 22.5));
            parameters.AddProjectedView(point: CAddIn.App.TransientGeometry.CreatePoint2d(8.5, 15.5));
            return parameters;
        }

        public static SheetParameters CreateViewsAuto(ModelInfo cdef, DrawingDocument drawDoc) {
            var doc = cdef.Definition.Document as Inventor.Document;
            if (doc.ReferencedDocuments.Count > 2) {
                var parameters = SheetParameters.CreateA33Views(cdef, drawDoc);
                parameters.AddTitleBlock("А3");
                return parameters;
            } else if (doc.ReferencedDocuments.Count == 2) {
                var parameters = SheetParameters.CreateA42Views(cdef, drawDoc);
                parameters.AddTitleBlock("А3");
                return parameters;
            } else {
                var parameters = SheetParameters.CreateA42Views(cdef, drawDoc);
                parameters.AddTitleBlock("А4");
                return parameters;
            }
        }
    }
}
