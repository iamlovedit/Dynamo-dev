using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using RevitServices.Persistence;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using Autodesk.DesignScript.Runtime;
using System.IO;

namespace BimFun
{
    /// <summary>
    /// utils
    /// </summary>
    [IsVisibleInDynamoLibrary(false)]
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

        /// <summary>
        /// export excel by filepath
        /// </summary>
        /// <param name="schedule"></param>
        /// <param name="exportPath"></param>
        /// <returns></returns>
        public static bool Write2Excel(ViewSchedule schedule, string exportPath)
        {
            if (schedule == null) throw new ArgumentNullException(nameof(schedule));
            TableData tableData = schedule.GetTableData();
            TableSectionData body = tableData.GetSectionData(SectionType.Body);

            int colums = body.NumberOfColumns;
            int rows = body.NumberOfRows;
            HSSFWorkbook workbook = new HSSFWorkbook();
            ISheet sheet = workbook.CreateSheet(schedule.Name);
            try
            {
                int index = 0;
                for (int i = 0; i < rows; i++)
                {
                    if (i == 1)
                    {
                        i++;
                        continue;
                    }
                    IRow row = sheet.CreateRow(i - index);
                    for (int j = 0; j < colums; j++)
                    {
                        double width = body.GetColumnWidthInPixels(j);
                        sheet.SetColumnWidth(j, (int)width * 50);
                        ICell cell = row.CreateCell(j);
                        string value = schedule.GetCellText(SectionType.Body, i, j);
                        cell.SetCellValue(value);
                    }
                    string fileName = schedule.Name + ".xls";
                    if (!Directory.Exists(exportPath))
                    {
                        Directory.CreateDirectory(exportPath);
                    }
                    string filePath = Path.Combine(exportPath, fileName);

                    using (FileStream fs = File.Create(filePath))
                    {
                        workbook.Write(fs);
                        fs.Close();
                    }
                }
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
    }
    [IsVisibleInDynamoLibrary(false)]
    public class FamilyLoadOptions : IFamilyLoadOptions
    {
        public bool OnFamilyFound(bool familyInUse, out bool overwriteParameterValues)
        {
            overwriteParameterValues = true;
            return true;
        }
        public bool OnSharedFamilyFound(Family sharedFamily, bool familyInUse, out FamilySource source, out bool overwriteParameterValues)
        {
            source = FamilySource.Project;
            overwriteParameterValues = true;
            return true;
        }
    }
}

