﻿using InvAddIn.UI;

namespace InvAddIn {
    public class CAddIn {
        public static Inventor.Application App { get; set; }

        public CAddIn(Inventor.Application app) {
            App = app;
        }

        public void CreateUI() {
            var ribbon = new ParametrizedRibbon();
            var button = new Button();
            ribbon.AddButton(button);
        }
    }
}
