using System;
using System.Linq;
using System.Collections.Generic;

using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.ApplicationServices;

namespace Dimension
{
    [Transaction(TransactionMode.Manual)]
    public class Command:IExternalCommand
    {
        private Document _doc;
        private UIDocument _uidoc;
        private UIApplication _uiapp;
        private const string _template = @"L:\4_Revit\Imperial Templates\Generic Model.rft";
        private XYZ[] _vert = {   new XYZ(0.0,    0.0,    0.0),
                                  new XYZ(10.0,   0.0,    0.0),
                                  new XYZ(10.0,   10.0,   0.0),
                                  new XYZ(0.0,    10.0,   0.0)
                              };
        private ReferenceArray _width;
        private ReferenceArray _height;

        private ReferenceArray _leftCon;
        private ReferenceArray _botCon;
        private ReferenceArray _rightCon;
        private ReferenceArray _topCon;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                _uiapp = commandData.Application;
                _uidoc = _uiapp.ActiveUIDocument;
                _doc = _uidoc.Document;

                BuildFamily(_doc);

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }


        private void BuildFamily(Document familyDoc)
        {
            try
            {
                Transaction tR = new Transaction(familyDoc, "Set References");
                if (tR.Start() == TransactionStatus.Started)
                {
                    SetReferences(familyDoc);
                    tR.Commit();
                }
                Extrusion extrusion = null;
                Transaction tE = new Transaction(familyDoc, "Create Extrusion");
                if (tE.Start() == TransactionStatus.Started)
                {
                    extrusion = CreateExtrusion(familyDoc);
                    tE.Commit();
                }
                if (null != extrusion)
                {
                    Transaction tC = new Transaction(familyDoc, "Set Constraints");
                    if (tC.Start() == TransactionStatus.Started)
                    {
                        SetConstraints(familyDoc,extrusion);
                        tC.Commit();
                    }
                }
            }

            catch (Exception ex)
            {
                TaskDialog.Show("Build Error", ex.Message);
            }
        }

        private void SetReferences(Document familyDoc)
        {
            try
            {
                FilteredElementCollector col = new FilteredElementCollector(familyDoc);
                col.OfClass(typeof(ReferencePlane)).WhereElementIsNotElementType().ToElements();

                foreach (ReferencePlane rplane in col)
                {
                    if (rplane.Name == @"Center (Left/Right)")
                    {
                        _rightCon = new ReferenceArray();
                        _leftCon = new ReferenceArray();
                        _width = new ReferenceArray();
                        _leftCon.Append(rplane.Reference);
                        _width.Append(rplane.Reference);
                        ReferencePlane newplane = familyDoc.FamilyCreate.NewReferencePlane(_vert[1], _vert[2], new XYZ(0, 0, 1), familyDoc.ActiveView);
                        _rightCon.Append(newplane.Reference);
                        _width.Append(newplane.Reference);
                    }
                    else if (rplane.Name == @"Center (Front/Back)")
                    {
                        _topCon = new ReferenceArray();
                        _botCon = new ReferenceArray();
                        _height = new ReferenceArray();
                        _botCon.Append(rplane.Reference);
                        _height.Append(rplane.Reference);
                        ReferencePlane newplane = familyDoc.FamilyCreate.NewReferencePlane(_vert[2], _vert[3], new XYZ(0, 0, 1), familyDoc.ActiveView);
                        _topCon.Append(newplane.Reference);
                        _height.Append(newplane.Reference);
                    }
                }
            }

            catch (Exception ex)
            {
                TaskDialog.Show("Reference Error", ex.Message);
            }
        }

        private Extrusion CreateExtrusion(Document familyDoc)
        {
            try
            {
                Plane plane = _uiapp.Application.Create.NewPlane(new XYZ(0.0, 0.0, 1.0), 
                                                                 new XYZ(0.0, 0.0, 0.0)
                                                                 );

                SketchPlane s_plane = familyDoc.FamilyCreate.NewSketchPlane(plane);
                ReferenceArray ra = new ReferenceArray();
                CurveArray profile = new CurveArray();
                CurveArrArray caa = new CurveArrArray();

                int i = 0;
                while (i < 3)
                {
                    profile.Append(familyDoc.Application.Create.NewLineBound(_vert[i],_vert[i + 1]));
                    ++i;
                }
                profile.Append(familyDoc.Application.Create.NewLineBound(_vert[i], _vert[i - i]));
                caa.Append(profile);
                Extrusion extrusion = familyDoc.FamilyCreate.NewExtrusion(true, caa, s_plane, 10.0);

                Line line = familyDoc.Application.Create.NewLine(_vert[0], _vert[1], true);
                ConstructParam(familyDoc, _width, line, "Width");
                line = familyDoc.Application.Create.NewLine(_vert[1], _vert[2], true);
                ConstructParam(familyDoc, _height, line, "Height");

                return extrusion;
            }

            catch (Exception ex)
            {
                TaskDialog.Show("Extrusion Error", ex.Message);

                return null;
            }
        }
        
        private void SetConstraints(Document familyDoc, Extrusion extrusion)
        {
            CurveArrArray curvesArr = new CurveArrArray();
            curvesArr = extrusion.Sketch.Profile;
            
            foreach (CurveArray ca in curvesArr)
            {
                CurveArrayIterator itor = ca.ForwardIterator();
                itor.Reset();
                itor.MoveNext();
                Line l = itor.Current as Line;
                _botCon.Append(l.Reference);
                itor.MoveNext();
                l = itor.Current as Line;
                _rightCon.Append(l.Reference);
                itor.MoveNext();
                l = itor.Current as Line;
                _topCon.Append(l.Reference);
                itor.MoveNext();
                l = itor.Current as Line;
                _leftCon.Append(l.Reference);
            }
            ReferenceArrayArray conArray = new ReferenceArrayArray();
            conArray.Append(_botCon);
            conArray.Append(_rightCon);
            conArray.Append(_topCon);
            conArray.Append(_leftCon);

            Line line = familyDoc.Application.Create.NewLine(_vert[1], _vert[2], true);
            ConstructConstraint(familyDoc, _rightCon, line);
            line = familyDoc.Application.Create.NewLine(_vert[2], _vert[3], true);
            ConstructConstraint(familyDoc, _topCon, line);
            line = familyDoc.Application.Create.NewLine(_vert[3], _vert[0], true);
            ConstructConstraint(familyDoc, _leftCon, line);
            line = familyDoc.Application.Create.NewLine(_vert[0], _vert[1], true);
            ConstructConstraint(familyDoc, _botCon, line);
        }

        private void ConstructConstraint(Document familyDoc, ReferenceArray ra, Line line)
        {
            Autodesk.Revit.DB.Dimension con = familyDoc.FamilyCreate.NewDimension(familyDoc.ActiveView, line, ra);
            //con.IsLocked = true;
        }

        private void ConstructParam(Document familyDoc, ReferenceArray ra, Line line, string label)
        {
            Autodesk.Revit.DB.Dimension dim = familyDoc.FamilyCreate.NewDimension(familyDoc.ActiveView, line, ra);
            FamilyParameter param = familyDoc.FamilyManager.AddParameter(label, BuiltInParameterGroup.PG_CONSTRAINTS, ParameterType.Length, false);
            dim.Label = param;
        }

        //private ModelCurve MakeLine(Document familyDoc, XYZ ptA, XYZ ptB)
        //{
        //    try
        //    {
        //        Line line = familyDoc.Application.Create.NewLine(ptA, ptB, true);
        //        XYZ norm = ptA.CrossProduct(ptB);
        //        if (norm.IsZeroLength()) norm = XYZ.BasisZ;
        //        Plane plane = familyDoc.Application.Create.NewPlane(norm, ptB);
                
        //        SketchPlane skplane = familyDoc.FamilyCreate.NewSketchPlane(plane);

        //        return familyDoc.FamilyCreate.NewModelCurve(line, skplane);
        //    }

        //    catch (Exception ex)
        //    {
        //        TaskDialog.Show("Line Error", ex.Message);

        //        return null;
        //    }
        //}

    }
}
