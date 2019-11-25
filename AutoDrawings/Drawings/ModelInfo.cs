using Inventor;

namespace InvAddIn.Drawings {
    internal class ModelInfo {
        public ComponentDefinition Definition { get; private set; }
        private Document doc;

        public string SheetFormat { get; set; } = "A4";
        public string PartNumber => doc.PropertySets["Design Tracking Properties"]["Part Number"].Value as string;
        public string Description => doc.PropertySets["Design Tracking Properties"]["Description"].Value as string;

        public ModelInfo(ComponentDefinition definition) {
            Definition = definition;
            doc = definition.Document as Document;
        }
    }
}
