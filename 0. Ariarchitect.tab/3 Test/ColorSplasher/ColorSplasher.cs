using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using WinForms = System.Windows.Forms;
using Drawing = System.Drawing;

class RFToolsAutomation // DO NOT CHANGE THE CLASS NAME
{
    public void Execute(string message, UIApplication MyUIApplication) // DO NOT CHANGE THE NAME OF THIS FUNCTION
    {
        UIDocument uidoc = MyUIApplication.ActiveUIDocument;
        Document doc = uidoc.Document;

        WinForms.Form form = new WinForms.Form();
        form.Text = "Color Splasher - Categories and Parameters";
        form.Width = 600;
        form.Height = 700;

        WinForms.Label label1 = new WinForms.Label();
        label1.Text = "Categories:";
        label1.Top = 10;
        label1.Left = 10;
        form.Controls.Add(label1);

        WinForms.ListBox lbCategories = new WinForms.ListBox();
        lbCategories.Top = 30;
        lbCategories.Left = 10;
        lbCategories.Width = 250;
        lbCategories.Height = 300;
        lbCategories.SelectionMode = WinForms.SelectionMode.One;
        form.Controls.Add(lbCategories);

        WinForms.Label label2 = new WinForms.Label();
        label2.Text = "Parameters:";
        label2.Top = 10;
        label2.Left = 300;
        form.Controls.Add(label2);

        WinForms.ListBox lbParameters = new WinForms.ListBox();
        lbParameters.Top = 30;
        lbParameters.Left = 300;
        lbParameters.Width = 250;
        lbParameters.Height = 300;
        lbParameters.SelectionMode = WinForms.SelectionMode.One;
        form.Controls.Add(lbParameters);

        WinForms.Label label3 = new WinForms.Label();
        label3.Text = "Values:";
        label3.Top = 340;
        label3.Left = 10;
        label3.Width = 100;
        form.Controls.Add(label3);

        WinForms.ListView lvValues = new WinForms.ListView();
        lvValues.Top = 360;
        lvValues.Left = 10;
        lvValues.Width = 540;
        lvValues.Height = 250;
        lvValues.View = WinForms.View.Details;
        lvValues.FullRowSelect = true;
        lvValues.GridLines = true;
        lvValues.Columns.Add("Value", 300);
        lvValues.Columns.Add("Color", 200);
        form.Controls.Add(lvValues);

        WinForms.Button btnApply = new WinForms.Button();
        btnApply.Text = "Apply colors";
        btnApply.Top = 620;
        btnApply.Left = 10;
        form.Controls.Add(btnApply);

        WinForms.Button btnSave = new WinForms.Button();
        btnSave.Text = "Save scheme";
        btnSave.Top = 620;
        btnSave.Left = 120;
        form.Controls.Add(btnSave);

        WinForms.Button btnLoad = new WinForms.Button();
        btnLoad.Text = "Load scheme";
        btnLoad.Top = 620;
        btnLoad.Left = 230;
        form.Controls.Add(btnLoad);

        WinForms.Button btnClear = new WinForms.Button();
        btnClear.Text = "Clear overrides";
        btnClear.Top = 620;
        btnClear.Left = 340;
        form.Controls.Add(btnClear);

        lvValues.MouseClick += (s, e) =>
        {
            if (lvValues.SelectedItems.Count > 0)
            {
                WinForms.ColorDialog colorDialog = new WinForms.ColorDialog();
                if (colorDialog.ShowDialog() == WinForms.DialogResult.OK)
                {
                    var item = lvValues.SelectedItems[0];
                    Drawing.Color newColor = colorDialog.Color;
                    item.SubItems[1].Text = newColor.ToArgb().ToString("X");
                    item.BackColor = newColor;
                }
            }
        };

        List<string> excludedCategories = new List<string>()
        {
            "<Room separation>", "Cameras", "Curtain wall grids", "Elevations", "Grids", "Model Groups",
            "Property Line segments", "Section Boxes", "Shaft openings", "Structural beam systems", "Views",
            "Structural opening cut", "Structural trusses", "<Space separation>", "Duct systems", "Lines",
            "Piping systems", "Matchline", "Center line", "Curtain Roof Grids", "Rectangular Straight wall opening"
        };

        List<BuiltInCategoryInfo> categoryList = new List<BuiltInCategoryInfo>();

        var collector = new FilteredElementCollector(doc, doc.ActiveView.Id)
            .WhereElementIsNotElementType()
            .WhereElementIsViewIndependent()
            .ToElements()
            .Where(e => e.Category != null && !excludedCategories.Contains(e.Category.Name))
            .ToList();

        foreach (var e in collector)
        {
            string catName = e.Category.Name;
            if (!categoryList.Any(c => c.Name == catName))
            {
                List<string> paramNames = new List<string>();

                foreach (Parameter p in e.Parameters)
                {
                    if (p.Definition != null && p.Definition.Name != "Category")
                    {
                        if (!paramNames.Contains(p.Definition.Name))
                            paramNames.Add(p.Definition.Name);
                    }
                }

                Element typeElement = doc.GetElement(e.GetTypeId());
                if (typeElement != null)
                {
                    foreach (Parameter p in typeElement.Parameters)
                    {
                        if (p.Definition != null && p.Definition.Name != "Category")
                        {
                            if (!paramNames.Contains(p.Definition.Name))
                                paramNames.Add(p.Definition.Name);
                        }
                    }
                }

                categoryList.Add(new BuiltInCategoryInfo
                {
                    Name = catName,
                    Parameters = paramNames.OrderBy(n => n).ToList()
                });
            }
        }

        foreach (var c in categoryList)
            lbCategories.Items.Add(c);

        lbCategories.SelectedIndexChanged += (s, e) =>
        {
            lbParameters.Items.Clear();
            if (lbCategories.SelectedItem != null)
            {
                var selected = (BuiltInCategoryInfo)lbCategories.SelectedItem;
                foreach (var p in selected.Parameters)
                {
                    lbParameters.Items.Add(p);
                }
            }
        };

        lbParameters.SelectedIndexChanged += (s, e) =>
        {
            lvValues.Items.Clear();
            if (lbParameters.SelectedItem != null && lbCategories.SelectedItem != null)
            {
                string paramName = lbParameters.SelectedItem.ToString();
                var selected = (BuiltInCategoryInfo)lbCategories.SelectedItem;
                HashSet<string> values = new HashSet<string>();
                Random rnd = new Random();

                foreach (var el in collector)
                {
                    if (el.Category != null && el.Category.Name == selected.Name)
                    {
                        string value = GetParameterValue(el, paramName, doc);
                        if (!string.IsNullOrEmpty(value))
                            values.Add(value);
                    }
                }

                foreach (string val in values)
                {
                    WinForms.ListViewItem item = new WinForms.ListViewItem(val);
                    Drawing.Color color = Drawing.Color.FromArgb(rnd.Next(256), rnd.Next(256), rnd.Next(256));
                    item.SubItems.Add(color.ToArgb().ToString("X"));
                    item.BackColor = color;
                    lvValues.Items.Add(item);
                }
            }
        };

        btnApply.Click += (s, e) =>
        {
            if (lbCategories.SelectedItem == null || lbParameters.SelectedItem == null)
                return;

            string paramName = lbParameters.SelectedItem.ToString();
            var selectedCat = (BuiltInCategoryInfo)lbCategories.SelectedItem;

            using (Transaction tx = new Transaction(doc, "Apply Color Overrides"))
            {
                tx.Start();

                foreach (var el in collector)
                {
                    if (el.Category == null || el.Category.Name != selectedCat.Name)
                        continue;

                    string val = GetParameterValue(el, paramName, doc);

                    foreach (WinForms.ListViewItem item in lvValues.Items)
                    {
                        string targetVal = item.Text;
                        if (val == targetVal)
                        {
                            Drawing.Color sysColor = item.BackColor;
                            Autodesk.Revit.DB.Color revitColor = new Autodesk.Revit.DB.Color(sysColor.R, sysColor.G, sysColor.B);

                            OverrideGraphicSettings ogs = new OverrideGraphicSettings();
                            ogs.SetProjectionLineColor(revitColor);
                            ogs.SetSurfaceForegroundPatternColor(revitColor);

                            // Use solid fill if possible
                            FillPatternElement solidFill = new FilteredElementCollector(doc)
                                .OfClass(typeof(FillPatternElement))
                                .Cast<FillPatternElement>()
                                .FirstOrDefault(f => f.GetFillPattern().IsSolidFill);

                            if (solidFill != null)
                            {
                                ogs.SetSurfaceForegroundPatternId(solidFill.Id);
                            }

                            doc.ActiveView.SetElementOverrides(el.Id, ogs);
                            break;
                        }
                    }
                }

                tx.Commit();
            }
        };

        btnSave.Click += (s, e) =>
        {
            if (lvValues.Items.Count == 0)
                return;

            WinForms.SaveFileDialog saveDialog = new WinForms.SaveFileDialog();
            saveDialog.Filter = "Color Scheme (*.csch)|*.csch";
            saveDialog.Title = "Save Color Scheme";

            if (saveDialog.ShowDialog() == WinForms.DialogResult.OK)
            {
                using (StreamWriter writer = new StreamWriter(saveDialog.FileName))
                {
                    foreach (WinForms.ListViewItem item in lvValues.Items)
                    {
                        Drawing.Color c = item.BackColor;
                        writer.WriteLine(item.Text + ":" + c.ToArgb().ToString("X"));
                    }
                }
            }
        };

        btnLoad.Click += (s, e) =>
        {
            if (lvValues.Items.Count == 0)
                return;

            WinForms.OpenFileDialog openDialog = new WinForms.OpenFileDialog();
            openDialog.Filter = "Color Scheme (*.csch)|*.csch";
            openDialog.Title = "Load Color Scheme";

            if (openDialog.ShowDialog() == WinForms.DialogResult.OK)
            {
                Dictionary<string, Drawing.Color> colorMap = new Dictionary<string, Drawing.Color>();
                using (StreamReader reader = new StreamReader(openDialog.FileName))
                {
                    string line;
                    while ((line = reader.ReadLine()) != null)
                    {
                        string[] parts = line.Split(':');
                        if (parts.Length == 2)
                        {
                            int argb;
                            if (Int32.TryParse(parts[1], System.Globalization.NumberStyles.HexNumber, null, out argb))
                            {
                                colorMap[parts[0]] = Drawing.Color.FromArgb(argb);
                            }
                        }
                    }
                }

                foreach (WinForms.ListViewItem item in lvValues.Items)
                {
                    if (colorMap.ContainsKey(item.Text))
                    {
                        Drawing.Color c = colorMap[item.Text];
                        item.BackColor = c;
                        item.SubItems[1].Text = c.ToArgb().ToString("X");
                    }
                }
            }
        };

        btnClear.Click += (s, e) =>
        {
            using (Transaction tx = new Transaction(doc, "Clear Element Overrides"))
            {
                tx.Start();
                foreach (var el in collector)
                {
                    doc.ActiveView.SetElementOverrides(el.Id, new OverrideGraphicSettings());
                }
                tx.Commit();
            }
        };

        WinForms.Application.Run(form);
    }

    private string GetParameterValue(Element e, string paramName, Document doc)
    {
        foreach (Parameter p in e.Parameters)
        {
            if (p.Definition != null && p.Definition.Name == paramName)
                return ReadParameter(p, doc);
        }

        Element typeElem = doc.GetElement(e.GetTypeId());
        if (typeElem != null)
        {
            foreach (Parameter p in typeElem.Parameters)
            {
                if (p.Definition != null && p.Definition.Name == paramName)
                    return ReadParameter(p, doc);
            }
        }

        return "";
    }

    private string ReadParameter(Parameter p, Document doc)
    {
        switch (p.StorageType)
        {
            case StorageType.String:
                return p.AsString();
            case StorageType.Double:
            case StorageType.Integer:
                return p.AsValueString();
            case StorageType.ElementId:
                ElementId id = p.AsElementId();
                if (id.IntegerValue >= 0)
                {
                    Element e = doc.GetElement(id);
                    if (e != null) return e.Name;
                }
                return id.IntegerValue.ToString();
            default:
                return "";
        }
    }

    class BuiltInCategoryInfo
    {
        public string Name { get; set; }
        public List<string> Parameters { get; set; }

        public override string ToString()
        {
            return Name;
        }
    }
}
