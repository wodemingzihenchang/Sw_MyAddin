using SolidWorks.Interop.sldworks;

namespace Sw_MyAddin.SLDPRT
{
    class Combined_Entity
    {
        public static void Function(ISldWorks swApp)
        {
            ModelDoc2 swModel = swApp.ActiveDoc;
            string folder = System.IO.Path.GetDirectoryName(swModel.GetPathName());
            string name = System.IO.Path.GetFileName(swModel.GetPathName());
            //创建文件夹
            System.IO.Directory.CreateDirectory(folder + "\\合并");
            bool boolstatus = swModel.Extension.SaveDeFeaturedFile(folder + "\\合并\\合并-" + name);

        }

    }
}
