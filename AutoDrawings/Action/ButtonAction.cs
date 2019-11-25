using InvAddIn.Drawings;
using InvAddIn.Extensions;
using InvAddIn.Processors;
using InvAddIn.UI;
using Inventor;

namespace InvAddIn.Action {
    internal class ButtonAction {
        private AssemblyDocument oAsmDoc;
        //private DrawingDocument oDrawDoc;
        private DrawingMaker Writer { get; set; }

        public ButtonAction() {
            oAsmDoc = CAddIn.App.ActiveDocument as AssemblyDocument;
            Writer = new DrawingMaker();

            Writer.Processors.Add(SheetProcessors.CreateViews);
            Writer.Processors.Add(SheetProcessors.ChangeTitleBlock);
            Writer.Processors.Add(SheetProcessors.AddPartsList);
            Writer.Processors.Add(SheetProcessors.AddDimensions);
            Writer.Processors.Add(SheetProcessors.AddCenterlines);
        }

        public void Run() {
            var oBOM = oAsmDoc.ComponentDefinition.BOM;
            oBOM.StructuredViewEnabled = true;
            oBOM.StructuredViewFirstLevelOnly = false;

            var oBOMView = oBOM.BOMViews[1];

            //Create drawing sheet for top level assembly document
            this.Action(oAsmDoc.ComponentDefinition as ComponentDefinition, 0);

            oBOMView.BOMViewTraverse(this.Action);
            var form = new FormMain(Writer);
            form.ShowDialog();
        }

        private void Action(ComponentDefinition cdef, int level) {
            if (!cdef.IsProjectPart()) return;
            Writer.AddModel(cdef);
        }
    }
}