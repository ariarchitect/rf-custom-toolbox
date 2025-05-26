//20250525 Oleg Ariarskiy
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Linq;
using System.Drawing;
using System.Windows.Forms;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

class RFToolsAutomation // DO NOT CHANGE THE CLASS NAME
{
    Dictionary<string, TreeNode> allNodes = new Dictionary<string, TreeNode>();

    public void Execute(string message, UIApplication MyUIApplication) // DO NOT CHANGE THE NAME OF THIS FUNCTION
    {
        UIDocument uidoc = MyUIApplication.ActiveUIDocument;
        Document doc = uidoc.Document;

        List<BuiltInCategory> realCategories = Enum.GetValues(typeof(BuiltInCategory))
            .Cast<BuiltInCategory>()
            .Where(bic =>
            {
                try
                {
                    Category cat = Category.GetCategory(doc, bic);
                    return cat != null && cat.CategoryType == CategoryType.Model &&
                           !cat.IsTagCategory &&
                           new FilteredElementCollector(doc).OfCategory(bic).WhereElementIsNotElementType().Any();
                }
                catch { return false; }
            })
            .OrderBy(bic => Category.GetCategory(doc, bic).Name)
            .ToList();

        var form = new System.Windows.Forms.Form();
        form.Text = "Select Categories, Parameters, Values";
        form.Width = 800;
        form.Height = 800;

        var tree = new System.Windows.Forms.TreeView();
        tree.Dock = DockStyle.Fill;

        var imageList = new ImageList();
        imageList.ImageSize = new Size(32, 32);
        imageList.Images.Add("unchecked", DrawCheckbox(System.Drawing.Color.White, Pens.Black, false, false, 32));
        imageList.Images.Add("partial", DrawCheckbox(System.Drawing.Color.White, Pens.Black, false, true, 32));
        imageList.Images.Add("checked", DrawCheckbox(System.Drawing.Color.White, Pens.Black, true, false, 32));
        tree.ImageList = imageList;

        tree.MouseDown += (sender, e) =>
        {
            TreeNode node = tree.GetNodeAt(e.X, e.Y);
            if (node == null) return;
            if (e.X > node.Bounds.Left - 32 && e.X < node.Bounds.Right)
            {
                int currentState = node.Tag is int ? (int)node.Tag : 0;
                int nextState = currentState == 2 ? 0 : 2;

                ApplyStateToNode(node, nextState);
                UpdateChildrenRecursive(node, nextState);
                UpdateParentState(node.Parent);
            }
        };

        foreach (BuiltInCategory bic in realCategories)
        {
            Category cat = Category.GetCategory(doc, bic);
            var catNode = new TreeNode(cat.Name);
            ApplyStateToNode(catNode, 0);

            var elements = new FilteredElementCollector(doc).OfCategory(bic).WhereElementIsNotElementType().ToList();
            var paramNames = new HashSet<string>();
            var allParams = new List<Parameter>();

            foreach (var e in elements)
                foreach (Parameter p in e.Parameters)
                    if (p.Definition != null && paramNames.Add(p.Definition.Name))
                        allParams.Add(p);

            foreach (var p in allParams.OrderBy(p => p.Definition.Name))
            {
                var paramNode = new TreeNode(p.Definition.Name);
                ApplyStateToNode(paramNode, 0);

                var values = new HashSet<string>();
                foreach (var el in elements)
                {
                    Parameter param = el.LookupParameter(p.Definition.Name);
                    if (param != null && param.HasValue)
                    {
                        string val = param.AsValueString() ?? param.AsString();
                        if (!string.IsNullOrEmpty(val))
                            values.Add(val);
                    }
                }

                foreach (var val in values.OrderBy(v => v))
                {
                    var valNode = new TreeNode(val);
                    ApplyStateToNode(valNode, 0);
                    paramNode.Nodes.Add(valNode);
                    allNodes[cat.Name + "||" + p.Definition.Name + "||" + val] = valNode;
                }

                catNode.Nodes.Add(paramNode);
                allNodes[cat.Name + "||" + p.Definition.Name] = paramNode;
            }

            tree.Nodes.Add(catNode);
            allNodes[cat.Name] = catNode;
        }

        var okBtn = new Button { Text = "OK", Dock = DockStyle.Bottom, Height = 40, DialogResult = DialogResult.OK };

        var exportBtn = new Button { Text = "Export", Dock = DockStyle.Bottom, Height = 30 };
        exportBtn.Click += (s, e) =>
        {
            var save = new SaveFileDialog { Filter = "Text files|*.txt" };
            if (save.ShowDialog() == DialogResult.OK)
            {
                SaveSelectionsToTxt(save.FileName, tree.Nodes);
                MessageBox.Show("Exported.");
            }
        };

        var importBtn = new Button { Text = "Import", Dock = DockStyle.Bottom, Height = 30 };
        importBtn.Click += (s, e) =>
        {
            var open = new OpenFileDialog { Filter = "Text files|*.txt" };
            if (open.ShowDialog() == DialogResult.OK)
            {
                LoadSelectionsFromTxt(open.FileName);
                MessageBox.Show("Imported.");
            }
        };

        form.Controls.Add(tree);
        form.Controls.Add(okBtn);
        form.Controls.Add(exportBtn);
        form.Controls.Add(importBtn);

        if (form.ShowDialog() == DialogResult.OK)
        {
            // Выделить подходящие элементы в Revit
            var selectedKeys = new HashSet<string>();
            foreach (var kv in allNodes)
            {
                if ((int)(kv.Value.Tag ?? 0) == 2 && kv.Key.Split(new string[] { "||" }, StringSplitOptions.None).Length == 3)
                    selectedKeys.Add(kv.Key);
            }

            var idsToSelect = new List<ElementId>();
            foreach (Element e in new FilteredElementCollector(doc).WhereElementIsNotElementType())
            {
                Category cat = e.Category;
                if (cat == null) continue;

                foreach (Parameter p in e.Parameters)
                {
                    if (p == null || p.Definition == null || !p.HasValue) continue;

                    string paramName = p.Definition.Name;
                    string val = p.AsValueString() ?? p.AsString();
                    if (!string.IsNullOrEmpty(val))
                    {
                        string key = cat.Name + "||" + paramName + "||" + val;
                        if (selectedKeys.Contains(key))
                        {
                            idsToSelect.Add(e.Id);
                            break;
                        }
                    }
                }
            }

            if (idsToSelect.Count > 0)
            {
                uidoc.Selection.SetElementIds(idsToSelect);
                TaskDialog.Show("Selection", idsToSelect.Count + " elements selected.");
            }
            else
            {
                TaskDialog.Show("Selection", "No matching elements found.");
            }
        }
    }

    void SaveSelectionsToTxt(string path, TreeNodeCollection nodes)
    {
        using (StreamWriter sw = new StreamWriter(path))
        {
            foreach (TreeNode cat in nodes)
            {
                if ((int)(cat.Tag ?? 0) == 0) continue;
                foreach (TreeNode param in cat.Nodes)
                {
                    if ((int)(param.Tag ?? 0) == 0) continue;
                    foreach (TreeNode val in param.Nodes)
                    {
                        if ((int)(val.Tag ?? 0) == 2)
                        {
                            sw.WriteLine(cat.Text + "||" + param.Text + "||" + val.Text);
                        }
                    }
                }
            }
        }
    }

    void LoadSelectionsFromTxt(string path)
    {
        string[] lines = File.ReadAllLines(path);
        foreach (string line in lines)
        {
            string[] parts = line.Split(new string[] {"||"}, StringSplitOptions.None);
            if (parts.Length != 3) continue;
            string key = parts[0] + "||" + parts[1] + "||" + parts[2];
            if (allNodes.ContainsKey(key))
                ApplyStateToNode(allNodes[key], 2);
        }
        foreach (var node in allNodes.Values)
            UpdateParentState(node.Parent);
    }

    Bitmap DrawCheckbox(System.Drawing.Color fillColor, Pen border, bool tick, bool partial, int size)
    {
        var bmp = new Bitmap(size, size);
        using (Graphics g = Graphics.FromImage(bmp))
        {
            g.Clear(System.Drawing.Color.Transparent);
            g.FillRectangle(new SolidBrush(fillColor), 4, 4, size - 8, size - 8);
            g.DrawRectangle(border, 4, 4, size - 8, size - 8);
            if (tick)
            {
                g.DrawLine(Pens.Black, size / 4, size / 2, size / 2, size - size / 4);
                g.DrawLine(Pens.Black, size / 2, size - size / 4, size - size / 5, size / 4);
            }
            else if (partial)
            {
                int boxSize = size / 4;
                g.FillRectangle(Brushes.Black, (size - boxSize) / 2, (size - boxSize / 2) / 2, boxSize, boxSize / 2);
            }
        }
        return bmp;
    }

    void ApplyStateToNode(TreeNode node, int state)
    {
        string key = state == 0 ? "unchecked" : state == 1 ? "partial" : "checked";
        node.ImageKey = key;
        node.SelectedImageKey = key;
        node.Tag = state;
    }

    void UpdateChildrenRecursive(TreeNode node, int state)
    {
        foreach (TreeNode child in node.Nodes)
        {
            ApplyStateToNode(child, state);
            UpdateChildrenRecursive(child, state);
        }
    }

    void UpdateParentState(TreeNode parent)
    {
        if (parent == null) return;

        var states = parent.Nodes.Cast<TreeNode>().Select(n => (int)(n.Tag ?? 0)).ToList();
        int resultState = states.All(s => s == 2) ? 2 :
                          states.All(s => s == 0) ? 0 : 1;

        ApplyStateToNode(parent, resultState);
        UpdateParentState(parent.Parent);
    }
}
