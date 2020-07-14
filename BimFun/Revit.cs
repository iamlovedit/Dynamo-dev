using Autodesk.DesignScript.Geometry;
using Autodesk.DesignScript.Runtime;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Revit.Elements;
using Revit.Elements.Views;
using Revit.GeometryConversion;
using RevitServices.Persistence;
using RevitServices.Transactions;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB.Mechanical;
using System.Diagnostics;

namespace BimFun
{
    /// <summary>
    /// For Revit
    /// </summary>
    public static class BimFun_Revit
    {
        /// <summary>
        /// Current document
        /// </summary>
        private static readonly Document doc;
        static BimFun_Revit()
        {
            doc = Utils.GetDocument();
        }
        /// <summary>
        /// 复制族类型
        /// </summary>
        /// <param name="type">用于复制的族类型</param>
        /// <param name="newNames">族类型名称</param>
        /// <returns>输出新的族类型</returns>
        public static List<Revit.Elements.FamilyType> DuplicateFamilyType(Revit.Elements.FamilyType type, List<string> newNames)
        {
            TransactionManager.Instance.EnsureInTransaction(doc);
            var newTypes = new List<Revit.Elements.FamilyType>();

            if (!newNames.Any())
            {
                throw new ArgumentNullException(nameof(newNames));
            }
            if (type == null)
            {
                throw new ArgumentException(nameof(type));
            }

            foreach (var name in newNames)
            {
                var newType = ((ElementType)type.InternalElement).Duplicate(name) as FamilySymbol;
                newTypes.Add(ElementWrapper.Wrap(newType, true));
            }
            return newTypes;
        }
        /// <summary>
        /// 删除图元
        /// </summary>
        /// <param name="elements">要删除的图元列表</param>
        /// <returns>删除结果</returns>
        [MultiReturn(new[] { "results", "ids" })]
        public static Dictionary<bool, ElementId> DeleteElements(IEnumerable<Revit.Elements.Element> elements)
        {
            TransactionManager.Instance.EnsureInTransaction(doc);
            Dictionary<bool, ElementId> results = new Dictionary<bool, ElementId>();
            foreach (var elem in elements)
            {
                ElementId id = new ElementId(elem.Id);
                try
                {
                    doc.Delete(id);
                    results.Add(true, id);
                }
                catch (Exception)
                {
                    results.Add(false, id);
                    continue;
                }
            }
            return results;
        }
        /// <summary>
        /// 将明细表导出至Excel
        /// </summary>
        /// <param name="schedule"></param>
        /// <param name="directory"></param>
        /// <returns></returns>
        public static bool ExportScheduleToExcel(ScheduleView schedule, string directory)
        {
            return Utils.Write2Excel((ViewSchedule)schedule.InternalElement, directory);
        }
        /// <summary>
        /// 切换图元连接顺序
        /// </summary>
        /// <param name="elements"></param>
        /// <param name="others"></param>
        public static void SwitchElementsOrder(IEnumerable<Revit.Elements.Element> elements, IEnumerable<Revit.Elements.Element> others)
        {
            if (!elements.Any()) throw new ArgumentNullException(nameof(elements));
            if (!others.Any()) throw new ArgumentNullException(nameof(others));
            TransactionManager.Instance.EnsureInTransaction(doc);
            try
            {
                foreach (var elem in elements)
                {
                    foreach (var elem1 in others)
                    {
                        if (JoinGeometryUtils.AreElementsJoined(doc, elem.InternalElement, elem1.InternalElement))
                        {
                            JoinGeometryUtils.SwitchJoinOrder(doc, elem.InternalElement, elem1.InternalElement);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }
        }
        /// <summary>
        /// 获取revit中指定类型的族
        /// </summary>
        /// <param name="category">指定类别</param>
        /// <returns></returns>
        public static List<Revit.Elements.FamilyType> GetFamilyTypes(Revit.Elements.Category category)
        {
            return new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).OfCategoryId(new ElementId(category.Id)).Cast<FamilySymbol>()
                  .ToList().OrderBy(x => x.Name, new FileComparer()).Select(x => ElementWrapper.Wrap(x, false)).ToList();

        }

        /// <summary>
        /// 通过xyz值放置族实例
        /// </summary>
        /// <param name="familyType">用于放置的族</param>
        /// <param name="x">X坐标值</param>
        /// <param name="y">Y坐标值</param>
        /// <param name="z">Z坐标值</param>
        /// <returns>放置好的族实例</returns>
        public static Revit.Elements.FamilyInstance FamilyinstanceByXYZ(Revit.Elements.FamilyType familyType, double x, double y, double z)
        {
            TransactionManager.Instance.EnsureInTransaction(doc);
            XYZ point = new XYZ(x, y, z);
            FamilySymbol familySymbol = (FamilySymbol)familyType.InternalElement;
            if (!familySymbol.IsActive) familySymbol.Activate();
            Autodesk.Revit.DB.FamilyInstance familyInstance = doc.Create.NewFamilyInstance(point, familySymbol, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
            return familyInstance.ToDSType(true) as Revit.Elements.FamilyInstance;
        }
        /// <summary>
        /// 通过xyz和标高放置族实例
        /// </summary>
        /// <param name="familyType">族类型</param>
        /// <param name="x">x值</param>
        /// <param name="y">y值</param>
        /// <param name="z">z值</param>
        /// <param name="level">标高</param>
        /// <returns></returns>
        public static Revit.Elements.FamilyInstance FamilyinstanceByXYZ(Revit.Elements.FamilyType familyType, double x, double y, double z, Revit.Elements.Level level)
        {
            TransactionManager.Instance.EnsureInTransaction(doc);
            XYZ point = new XYZ(x, y, z);
            FamilySymbol familySymbol = (FamilySymbol)familyType.InternalElement;
            if (!familySymbol.IsActive) familySymbol.Activate();
            Autodesk.Revit.DB.FamilyInstance familyInstance = doc.Create.NewFamilyInstance(point, familySymbol, (Autodesk.Revit.DB.Level)level.InternalElement,
                Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
            return familyInstance.ToDSType(true) as Revit.Elements.FamilyInstance;
        }
        /// <summary>
        /// 通过标高和直线创建风管
        /// </summary>
        /// <param name="systemType">系统类型</param>
        /// <param name="ductType">风管类型</param>
        /// <param name="level">标高</param>
        /// <param name="line">直线</param>
        /// <returns></returns>
        public static Revit.Elements.Element DuctByLineAndLevel(Revit.Elements.FamilyType systemType, Revit.Elements.FamilyType ductType, Revit.Elements.Level level, Autodesk.DesignScript.Geometry.Line line)
        {
            TransactionManager.Instance.EnsureInTransaction(doc);
            XYZ startPoint = line.StartPoint.ToRevitType();
            XYZ endPoint = line.EndPoint.ToRevitType();
            Duct duct = Duct.Create(doc, new ElementId(systemType.Id), new ElementId(ductType.Id), new ElementId(level.Id), startPoint, endPoint);
            return ElementWrapper.ToDSType(duct, true);
        }
        /// <summary>
        /// 将Geometry导入Revit文档中
        /// </summary>
        /// <param name="solids">Solids</param>
        /// <param name="familyName">族名称</param>
        /// <param name="templatePath">族样板路径</param>
        /// <param name="categoryName">族类别名称</param>
        /// <returns></returns>
        public static Revit.Elements.FamilyInstance NewInstance(IEnumerable<Autodesk.DesignScript.Geometry.Solid> solids, string familyName, string templatePath, string categoryName)
        {
            string[] m_availableViews = { "ThreeD", "FloorPlan", "EngineeringPlan", "CeilingPlan", "Elevation", "Section" };
            try
            {
                XYZ m_point = XYZ.Zero;
                SATImportOptions m_importOptions = new SATImportOptions
                {
                    Placement = ImportPlacement.Origin,
                    Unit = ImportUnit.Foot,
                };
                DisplayUnitType m_units = doc.GetUnits().GetFormatOptions(UnitType.UT_Length).DisplayUnits;
                double m_factor = UnitUtils.ConvertFromInternalUnits(1, m_units);
                Geometry m_refGeo = solids.FirstOrDefault();
                IEnumerable<Geometry> m_results = m_factor != 1 ? solids.Select(g => g.Scale(m_factor)) : solids;

                Vector m_vector = Vector.ByTwoPoints(m_refGeo.BoundingBox.MinPoint, m_refGeo.BoundingBox.MaxPoint);
                IEnumerable<Geometry> m_tranlatedGeos = m_results.Select(g => g.Translate(m_vector));
                Autodesk.Revit.ApplicationServices.Application m_app = doc.Application;
                Document m_familyDoc = m_app.NewFamilyDocument(templatePath);
                //先将geometry导出位sat 然后再递归找出solid
                string m_tempSat = Path.Combine(Path.GetTempPath(), $"{familyName}.sat");

                string m_sat = Geometry.ExportToSAT(m_tranlatedGeos, m_tempSat);
                Autodesk.Revit.DB.View m_targetView = new FilteredElementCollector(m_familyDoc).OfClass(typeof(Autodesk.Revit.DB.View)).OfType<Autodesk.Revit.DB.View>()
                      .Where(v => IsAcceptable(v)).FirstOrDefault();
                TransactionManager m_trans = TransactionManager.Instance;
                m_trans.EnsureInTransaction(m_familyDoc);
                ElementId m_satId = m_familyDoc.Import(m_sat, m_importOptions, m_targetView);
                Autodesk.Revit.DB.Element m_satElement = m_familyDoc.GetElement(m_satId);
                Autodesk.Revit.DB.Solid m_solid = GetSolid(m_satElement.get_Geometry(new Options() { DetailLevel = ViewDetailLevel.Fine, ComputeReferences = true }));
                m_familyDoc.Delete(m_satId);

                Autodesk.Revit.DB.Category m_familyCategory = m_familyDoc.Settings.Categories.get_Item(categoryName);
                m_familyDoc.OwnerFamily.FamilyCategory = m_familyCategory;

                FreeFormElement m_form = FreeFormElement.Create(m_familyDoc, m_solid);
                m_trans.ForceCloseTransaction();
                string m_tempRfa = Path.Combine(Path.GetTempPath(), $"{familyName}.rfa");
                m_familyDoc.SaveAs(m_tempRfa, new SaveAsOptions() { OverwriteExistingFile = true });

                Autodesk.Revit.DB.Family m_loadedFamily = m_familyDoc.LoadFamily(doc, new FamilyLoadOptions());
                m_familyDoc.Close(false);
                if (File.Exists(m_tempRfa))
                {
                    File.Delete(m_tempRfa);
                }
                FamilySymbol m_symbol = m_loadedFamily.GetFamilySymbolIds().Select(doc.GetElement).OfType<FamilySymbol>().FirstOrDefault();
                m_trans.EnsureInTransaction(doc);
                if (!m_symbol.IsActive)
                {
                    m_symbol.Activate();
                }
                Autodesk.Revit.DB.FamilyInstance m_instance = doc.Create.NewFamilyInstance(m_point, m_symbol, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                ElementTransformUtils.MoveElement(doc, m_instance.Id, m_vector.ToRevitType());
                m_trans.TransactionTaskDone();
                return m_instance.ToDSType(false) as Revit.Elements.FamilyInstance;
            }
            catch (Exception e)
            {
                Trace.WriteLine(e.Message);
                return null;
            }


            bool IsAcceptable(Autodesk.Revit.DB.View view)
            {
                if (view.IsTemplate) return false;
                foreach (var name in m_availableViews)
                {
                    if (view.ViewType.ToString() == name)
                    {
                        return true;
                    }
                }
                return false;
            }
        }
        private static Autodesk.Revit.DB.Solid GetSolid(GeometryElement geometryElement)
        {
            Autodesk.Revit.DB.Solid m_result = null;
            foreach (var m_geoObj in geometryElement)
            {
                if (m_geoObj is Autodesk.Revit.DB.Solid m_solid && m_solid.Volume > 0)
                {
                    m_result = m_solid;
                }
                else if (m_geoObj is GeometryInstance geometryInstance)
                {
                    m_result = GetSolid(geometryInstance.GetInstanceGeometry());
                }
            }
            return m_result;
        }
    }
}
