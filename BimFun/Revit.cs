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
        public static Document doc = Utils.GetDocument();

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
                  .ToList().OrderBy(x => x.Name, new FileComparer()).ToList().ConvertAll(x => ElementWrapper.Wrap(x as FamilySymbol, false));

        }
    }
}
