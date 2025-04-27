using System;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;


namespace Sw_toolkit
{
    //另存为
    class SW_Export
    {
        public static void Export_PDF(SldWorks swApp)//导出PDF
        {
            // 获取当前活动的文档
            ModelDoc2 swModel = swApp.ActiveDoc;
            ModelDocExtension swModExt = (ModelDocExtension)swModel.Extension;
            int errors = 0;
            int warnings = 0;

            string filename = @"C:\foodprocessor.pdf";

            ExportPdfData swExportPDFData = (ExportPdfData)swApp.GetExportFileData((int)swExportDataFileType_e.swExportPdfData);
            bool boolstatus = swModExt.SaveAs(filename, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_Silent, swExportPDFData, ref errors, ref warnings);
        }
    }
}
