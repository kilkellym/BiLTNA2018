/*
 * Created by SharpDevelop.
 * User: micha
 * Date: 8/11/2018
 * Time: 6:41 AM
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI.Selection;
using System.Collections.Generic;
using System.Linq;

namespace MyFirstModule
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.DB.Macros.AddInId("806AE212-5106-486E-8236-A1082205BA1C")]
	public partial class ThisDocument
	{
		private void Module_Startup(object sender, EventArgs e)
		{

		}

		private void Module_Shutdown(object sender, EventArgs e)
		{

		}

		#region Revit Macros generated code
		private void InternalStartup()
		{
			this.Startup += new System.EventHandler(Module_Startup);
			this.Shutdown += new System.EventHandler(Module_Shutdown);
		}
		#endregion
		
		public void LinesToWalls() {
			// define current document
			Document curDoc = this.Application.ActiveUIDocument.Document;
			
			// define counter
			int counter = 0;
			
			// get all lines
			FilteredElementCollector lineColl = new FilteredElementCollector(curDoc);
			lineColl.OfClass(typeof(CurveElement));
			
			// get only model curves - if needed
			//CurveElementFilter cef = new CurveElementFilter(CurveElementType.ModelCurve);
			//lineColl.WherePasses(cef);
			
			// get all levels
			FilteredElementCollector lColl = new FilteredElementCollector(curDoc);
			lColl.OfCategory(BuiltInCategory.OST_Levels);
			lColl.WhereElementIsNotElementType();
			
			// start transaction
			Transaction t = new Transaction(curDoc, "Lines to walls");
			t.Start();
			
			// loop through model lines
			foreach (CurveElement curCurve in lineColl) {
				// create wall
				Wall w = Wall.Create(curDoc, curCurve.GeometryCurve, lColl.FirstElementId(), false);
			
				// increment counter
				counter++;
			}
			
			// commit changes
			t.Commit();
			t.Dispose();
			
			// alert user
			TaskDialog.Show("BiLT NA 2018", "Created " + counter.ToString() + " new walls.");
			
		}
		
		public void SheetsFromViews() {
			// define current document
			Document curDoc = this.Application.ActiveUIDocument.Document;
			
			// define sheet num and counter
			int sheetNum = 101;
			int counter = 0;
			
			// get all plan views
			FilteredElementCollector planColl = new FilteredElementCollector(curDoc);
			planColl.OfClass(typeof(ViewPlan));
			
			// get all titleblocks
			FilteredElementCollector tblockColl = new FilteredElementCollector(curDoc);
			tblockColl.OfCategory(BuiltInCategory.OST_TitleBlocks);
			
			// start transaction
			Transaction t = new Transaction(curDoc, "Sheets from Plans");
			t.Start();
			
			// loop through plan views
			foreach (View v in planColl) {
				if(v.IsTemplate == false) {
					if(v.ViewType == ViewType.FloorPlan) {
						ViewSheet newVS = ViewSheet.Create(curDoc, tblockColl.FirstElementId());
				
						// update sheet number and sheet name
						newVS.SheetNumber = "A" + sheetNum.ToString();
						newVS.Name = v.Name;
						
						// create viewport and add view to sheet
						Viewport vp = Viewport.Create(curDoc, newVS.Id, v.Id, new XYZ(0,0,0));
						
						// increment 
						counter++;
						sheetNum++;
					}
				}
			}
			
			// commit changes
			t.Commit();
			t.Dispose();
			
			// alert user
			TaskDialog.Show("Complete", "Created " + counter.ToString() + " new sheets.");
		}
		
		public void DeleteUnusedViews() {
			// define current document
			Document curDoc = this.Application.ActiveUIDocument.Document;
			
			// get all views
			FilteredElementCollector viewColl = new FilteredElementCollector(curDoc);
			viewColl.OfCategory(BuiltInCategory.OST_Views);
			
			// get all sheets
			FilteredElementCollector sheetColl = new FilteredElementCollector(curDoc);
			sheetColl.OfCategory(BuiltInCategory.OST_Sheets);
			
			// create list of sheets to delete
			List<View> viewsToDelete = new List<View>();
			
			// loop through view and check each one
			foreach (View curView in viewColl) {
				// check if view is a template
				if(curView.IsTemplate == false) {
					// check if view can be added to sheet
					if(Viewport.CanAddViewToSheet(curDoc, sheetColl.FirstElementId(), curView.Id) == true) {
					   	// check if view has prefix
					   	if(curView.Name.Contains("working") == false) {
					   		// add view to delete list
					   		viewsToDelete.Add(curView);
					   	}
					}
				}
			}
			
			// create transaction
			Transaction t = new Transaction(curDoc, "Delete unused views");
			t.Start();
			
			// delete views in list
			foreach (View viewToDelete in viewsToDelete) {
				// delete view
				curDoc.Delete(viewToDelete.Id);
			}
			
			// commit changes
			t.Commit();
			t.Dispose();
			
			// alert user
			TaskDialog.Show("Complete", "Deleted " + viewsToDelete.Count().ToString() + " unused views.");
		}
		
		
	}
}