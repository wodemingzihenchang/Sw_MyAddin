
using SolidWorks.Interop.sldworks;
using SolidWorks.Visualize.Interfaces;
using System.Windows.Forms;

namespace My_SWAddin.SLDPRT
{
    class VisualizeAddin
    {
        /// <summary>
        /// 在SW使用SWV进行简单渲染
        /// </summary>
        /// <param name="SwApp"></param>
        public static void DoSimpleRender(ISldWorks SwApp)
        {
            //string VisualizeAddinProgID = "";

            // Create a FolderBrowserDialog
            var folderDialog = new FolderBrowserDialog();

            // Show the dialog and get the result
            var result = folderDialog.ShowDialog();

            // If user selects a folder, update the OutputFolder property with the selected folder path
            if (result != DialogResult.OK)
            {
                return;
            }

            IVisualizeAddin vizAddin = (IVisualizeAddin)SwApp.GetAddInObject("SolidWorks.Visualize.Implementation.VisualizeAddin");
            IVisualizeAddinManager vizAddingMgr = vizAddin?.GetAddinManager();
            if (vizAddingMgr == null)
            {
                return;
            }

            vizAddingMgr.RenderOptions.OutputFolder = folderDialog.SelectedPath;
            vizAddingMgr.Render();
        }
    }
}
