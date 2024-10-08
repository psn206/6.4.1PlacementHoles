
using Autodesk.Revit.Attributes;

using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.DB.Plumbing;

namespace PlacementHoles
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {

            Document arDoc = commandData.Application.ActiveUIDocument.Document;        
            Document ovDoc = arDoc.Application.Documents.OfType<Document>().Where(x => x.Title.Contains("ОВ")).FirstOrDefault();
            if (ovDoc == null)
            {
                TaskDialog.Show("Ошибка", "Не найден файл ОВ");
                return Result.Succeeded;
            }

            FamilySymbol familySymbol = new FilteredElementCollector(arDoc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_GenericModel)
                .OfType<FamilySymbol>()
                .Where(x => x.FamilyName.Equals("Отверстие"))
                .FirstOrDefault();

            if (familySymbol == null)
            {
                TaskDialog.Show("Ошибка", "Семейство отверстие не найдено!");
                return Result.Succeeded;
            }

            List<Duct> ducts = new FilteredElementCollector(ovDoc)
                .OfClass(typeof(Duct))
                .OfType<Duct>()
                .ToList();

            List<Pipe> pipes = new FilteredElementCollector(ovDoc)
               .OfClass(typeof(Pipe))
               .OfType<Pipe>()
               .ToList();
            List<MEPCurve> communications = new List<MEPCurve>();
            communications.AddRange(pipes);
            communications.AddRange(ducts);          

            View3D view3D = new FilteredElementCollector(arDoc)
              .OfClass(typeof(View3D))
              .OfType<View3D>()
              .Where(x => !x.IsTemplate)
              .FirstOrDefault();
            if (view3D == null)
            {
                TaskDialog.Show("Ошибка", "Отсутствует 3D вид!");
                return Result.Succeeded;
            }

            ReferenceIntersector referenceIntersector = new ReferenceIntersector(new ElementClassFilter(typeof(Wall)), FindReferenceTarget.Element, view3D);
            Transaction ts = new Transaction(arDoc, "Вставка отверстия");
            ts.Start();            
           
            foreach ( MEPCurve c in communications)
            {
                Line curve = (c.Location as LocationCurve).Curve as Line;
                XYZ point = curve.GetEndPoint(0);
                XYZ derection = curve.Direction;

                List<ReferenceWithContext> intersections = referenceIntersector.Find(point, derection)
                .Where(x => x.Proximity <= curve.Length)
                .Distinct(new ReferenceWithContextElementEqualityComparer())
                .ToList();

                foreach (ReferenceWithContext refer in intersections)
                {
                    double proximity = refer.Proximity;
                    Reference reference = refer.GetReference();
                    Wall wall = arDoc.GetElement(reference.ElementId) as Wall;
                    Level level = arDoc.GetElement(wall.LevelId) as Level;
                    XYZ pointHole = point + (derection * proximity);
                    FamilyInstance hole = arDoc.Create.NewFamilyInstance(pointHole, familySymbol, wall, level, StructuralType.NonStructural);
                    Parameter width = hole.LookupParameter("Ширина");
                    Parameter heigth = hole.LookupParameter("Высота");                  
                    width.Set(c.Diameter);
                    heigth.Set(c.Diameter);                
                }
            }    
   
            ts.Commit();
            return Result.Succeeded;
        }
    }






    public class ReferenceWithContextElementEqualityComparer : IEqualityComparer<ReferenceWithContext>
    {
        public bool Equals(ReferenceWithContext x, ReferenceWithContext y)
        {
            if (ReferenceEquals(x, y)) return true;
            if (ReferenceEquals(null, x)) return false;
            if (ReferenceEquals(null, y)) return false;

            var xReference = x.GetReference();

            var yReference = y.GetReference();

            return xReference.LinkedElementId == yReference.LinkedElementId
                       && xReference.ElementId == yReference.ElementId;
        }

        public int GetHashCode(ReferenceWithContext obj)
        {
            var reference = obj.GetReference();

            unchecked
            {
                return (reference.LinkedElementId.GetHashCode() * 397) ^ reference.ElementId.GetHashCode();
            }
        }
    }
}
