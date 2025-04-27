using System;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace Sw_toolkit
{
    class FreezeBar
    {
        public static SldWorks swApp = (SldWorks)SwModelDoc.ConnectToSolidWorks();
        /// <summary>
        /// 根据选择的零件路径，进行冻结全部特征
        /// </summary>
        /// <param name="OpenPart"></param>
        public static void FreezeBar_All(string filepath)
        {
            //开启冻结栏
            swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swUserEnableFreezeBar, true);
            ModelDoc2 swDoc = (ModelDoc2)swApp.OpenDoc(filepath, 1); //激活路径下的文件
            if (swDoc.GetType() != 1) { return; }//判断为零件
            else
            {
                swDoc.ForceRebuild3(false);
                swDoc.FeatureManager.EditFreeze((int)swMoveFreezeBarTo_e.swMoveFreezeBarToEnd, "", true);
                swDoc.Save(); swApp.CloseDoc(swDoc.GetPathName());
            }
        }
    }
}
