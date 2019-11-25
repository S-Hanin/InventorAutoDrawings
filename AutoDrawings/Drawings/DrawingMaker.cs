using Inventor;
using System;
using System.Collections.Generic;
using System.Linq;

namespace InvAddIn.Drawings {
    internal class DrawingMaker {
        public List<SheetParameters> SheetsParameters { get; private set; }
        internal readonly List<ModelInfo> Parts = new List<ModelInfo>();
        public DrawingDocument DrawDoc { get; set; }
        internal readonly List<Action<SheetParameters, Sheet>> Processors = new List<Action<SheetParameters, Sheet>>();


        public ComponentDefinition AddModel(ComponentDefinition cdef) {
            Parts.Add(new ModelInfo(cdef));
            return cdef;
        }


        public void MakeDrawing() {
            DrawDoc = AddDrawingDocument();
            SheetsParameters = CreateSheetsParameters(DrawDoc);

            foreach (var item in SheetsParameters) {
                var sheet = CreateSheet(item, DrawDoc);
                foreach (var processor in Processors) {
                    processor(item, sheet);
                }
            }
        }


        private List<SheetParameters> CreateSheetsParameters(DrawingDocument drawDoc) {
            return Parts.Select(item => SheetParameters.CreateViewsAuto(item, drawDoc)).ToList();
        }


        private DrawingDocument AddDrawingDocument() {
            var template = CAddIn.App.FileManager.GetTemplateFile(DocumentTypeEnum.kDrawingDocumentObject);
            return CAddIn.App.Documents.Add(DocumentTypeEnum.kDrawingDocumentObject, template) as Inventor.DrawingDocument;
        }


        private Sheet CreateSheet(SheetParameters parameters, DrawingDocument drawDoc) {
            return DrawDoc.Sheets.AddUsingSheetFormat(parameters.SheetFormat);
        }
    }
}
