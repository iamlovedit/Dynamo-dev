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
    }
}
