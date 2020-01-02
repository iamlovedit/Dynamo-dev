using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using RevitServices.Persistence;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace BimFun
{
    public static class Utils
    {
        /// <summary>
        /// Get the current document of Revit
        /// </summary>
        /// <returns></returns>
        public static Document GetDocument() => DocumentManager.Instance.CurrentDBDocument;
        /// <summary>
        /// Get the current uidocument of Revit
        /// </summary>
        /// <returns></returns>
        public static UIDocument GetUIDocument() => DocumentManager.Instance.CurrentUIDocument;
        /// <summary>
        /// Get the current uiapplication of Revit
        /// </summary>
        /// <returns></returns>
        public static UIApplication GetUIApplication() => DocumentManager.Instance.CurrentUIApplication;
    }
}
