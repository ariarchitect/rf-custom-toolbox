using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

class RFToolsAutomation // DO NOT CHANGE THE CLASS NAME
{
    public void Execute(string message, UIApplication MyUIApplication) // DO NOT CHANGE THE NAME OF THIS FUNCTION
    {
        UIDocument uidoc = MyUIApplication.ActiveUIDocument;
        Document doc = uidoc.Document;

        // Get all linked models in the project
        List<RevitLinkInstance> linkInstances = new FilteredElementCollector(doc)
            .OfClass(typeof(RevitLinkInstance))
            .Cast<RevitLinkInstance>()
            .ToList();

        if (linkInstances.Count == 0)
        {
            TaskDialog.Show("BBox LinkdID", "No linked models found.");
            return;
        }

        // Create the form
        System.Windows.Forms.Form form = new System.Windows.Forms.Form();
        form.Text = "BBox LinkdID";
        form.Width = 300;
        form.Height = 180;
        form.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
        form.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
        form.MinimizeBox = false;
        form.MaximizeBox = false;

        System.Windows.Forms.ComboBox comboBox = new System.Windows.Forms.ComboBox();
        comboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
        comboBox.Width = 200;
        foreach (var link in linkInstances)
        {
            comboBox.Items.Add(link.Name);
        }
        comboBox.SelectedIndex = 0;
        comboBox.Top = 20;
        comboBox.Left = 40;

        System.Windows.Forms.TextBox inputBox = new System.Windows.Forms.TextBox();
        inputBox.Width = 200;
        inputBox.Top = 60;
        inputBox.Left = 40;

        System.Windows.Forms.Button okButton = new System.Windows.Forms.Button();
        okButton.Text = "OK";
        okButton.Width = 80;
        okButton.Top = 100;
        okButton.Left = 100;
        okButton.DialogResult = System.Windows.Forms.DialogResult.OK;
        form.AcceptButton = okButton;

        form.Controls.Add(comboBox);
        form.Controls.Add(inputBox);
        form.Controls.Add(okButton);

        BoundingBoxXYZ transformedBox = null;
        RevitLinkInstance selectedLink = null;

        if (form.ShowDialog() == System.Windows.Forms.DialogResult.OK)
        {
            string selectedLinkName = comboBox.SelectedItem.ToString();
            int userNumber = 0;
            if (!int.TryParse(inputBox.Text, out userNumber))
            {
                TaskDialog.Show("Input Error", "Please enter a valid integer.");
                return;
            }

            // Find the selected linked model
            foreach (var link in linkInstances)
            {
                if (link.Name == selectedLinkName)
                {
                    selectedLink = link;
                    break;
                }
            }
            if (selectedLink == null)
            {
                TaskDialog.Show("Error", "Could not find the selected linked model.");
                return;
            }

            Document linkedDoc = selectedLink.GetLinkDocument();
            if (linkedDoc == null)
            {
                TaskDialog.Show("Error", "The linked model is not loaded.");
                return;
            }

            ElementId linkedId = new ElementId((long)userNumber);
            Element linkedElement = linkedDoc.GetElement(linkedId);

            if (linkedElement == null)
            {
                TaskDialog.Show("Result", string.Format("Element with ID {0} was not found in {1}.", userNumber, selectedLinkName));
                return;
            }

            string typeName = linkedElement.GetType().Name;

            // Get the bounding box from the linked element
            BoundingBoxXYZ bbox = linkedElement.get_BoundingBox(null);

            if (bbox != null)
            {
                Transform transform = selectedLink.GetTotalTransform();
                XYZ min = transform.OfPoint(bbox.Min);
                XYZ max = transform.OfPoint(bbox.Max);

                transformedBox = new BoundingBoxXYZ();
                transformedBox.Min = min;
                transformedBox.Max = max;

                string info = string.Format(
                    "Linked model: {0}\nElement ID: {1}\nElement type: {2}\n\nBoundingBox:\nX: {3:F2} - {4:F2}\nY: {5:F2} - {6:F2}\nZ: {7:F2} - {8:F2}",
                    selectedLinkName,
                    userNumber,
                    typeName,
                    min.X, max.X,
                    min.Y, max.Y,
                    min.Z, max.Z
                );

                TaskDialog.Show("Element Info", info);
            }
            else
            {
                TaskDialog.Show("Element Info", string.Format("Element found: {0}, but it has no BoundingBox.", typeName));
            }
        }

        // Apply Section Box and zoom using ZoomAndCenterRectangle
        if (transformedBox != null)
        {
            View3D view3D = doc.ActiveView as View3D;
            if (view3D == null || view3D.IsTemplate)
            {
                TaskDialog.Show("Error", "Active view is not a 3D view or is a template.");
                return;
            }

            using (Transaction t = new Transaction(doc, "Set Section Box"))
            {
                t.Start();
                view3D.IsSectionBoxActive = true;
                view3D.SetSectionBox(transformedBox);
                t.Commit();
            }

			// Use ZoomAndCenterRectangle to fit the (expanded) section box in view
			IList<UIView> uiviews = uidoc.GetOpenUIViews();
			UIView uiview = null;
			foreach (UIView uv in uiviews)
			{
				if (uv.ViewId == view3D.Id)
				{
					uiview = uv;
					break;
				}
			}

			if (uiview != null)
			{
				XYZ pt1 = transformedBox.Min;
				XYZ pt2 = transformedBox.Max;

				// Expand bounding box around center
				double scale = 2.0; // make this configurable
				XYZ center = (pt1 + pt2) * 0.5;
				XYZ halfSize = (pt2 - center) * scale;

				XYZ newPt1 = center - halfSize;
				XYZ newPt2 = center + halfSize;

				uiview.ZoomAndCenterRectangle(newPt2, newPt1); // top-right, bottom-left
			}

        }
    }
}
