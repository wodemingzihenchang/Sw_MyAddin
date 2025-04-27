using SolidWorks.Interop.sldworks;

namespace Sw_MyAddin
{
    class SW_Event
    {
        //用于事件对象共享。
        private static SldWorks swApp;
        private static PartDoc pDoc;
        private static AssemblyDoc aDoc;
        private static DrawingDoc dDoc;
        public SW_Event(SldWorks Addin_swApp) { swApp = Addin_swApp; }



        #region Event Methods 事件方法

        //保存事件 
        private int SaveEvent()
        {
            //swApp.SendMsgToUser("触发保存事件");
            return 0;
        }
        private int pDoc_Save_FileSaveNotify(string FileName)
        {
            SaveEvent();
            pDoc.FileSaveNotify -= pDoc_Save_FileSaveNotify;
            return 0;
        }
        private int aDoc_Save_FileSaveNotify(string FileName)
        {
            SaveEvent();
            aDoc.FileSaveNotify -= aDoc_Save_FileSaveNotify;
            return 0;
        }
        private int dDoc_Save_FileSaveNotify(string FileName)
        {
            SaveEvent();
            dDoc.FileSaveNotify -= dDoc_Save_FileSaveNotify;
            return 0;
        }

        //选择事件
        private static int pDoc_UserSelectionPostNotify()
        {
            int functionReturnValue = 0;
            swApp.SendMsgToUser2("一个实体被选中了在零件文档中。", 2, 2);
            return functionReturnValue;
        }

        #endregion






        #region Event Handlers 事件监控

        //保存事件
        public int SaveDoc()
        {
            swApp.ActiveDocChangeNotify += swapp_ActiveDocChangeNotify;

            ModelDoc mDoc = (ModelDoc)swApp.ActiveDoc;
            if (mDoc == null) { pDoc = null; aDoc = null; return 0; }
            if (mDoc.GetType() == 1) { pDoc = (PartDoc)mDoc; aDoc = null; dDoc = null; }
            else if (mDoc.GetType() == 2) { aDoc = (AssemblyDoc)mDoc; pDoc = null; dDoc = null; }
            else if (mDoc.GetType() == 3) { dDoc = (DrawingDoc)mDoc; pDoc = null; aDoc = null; }

            if ((pDoc != null)) { pDoc.FileSaveNotify += pDoc_Save_FileSaveNotify; }
            if ((aDoc != null)) { aDoc.FileSaveNotify += aDoc_Save_FileSaveNotify; }
            if ((dDoc != null)) { dDoc.FileSaveNotify += dDoc_Save_FileSaveNotify; }
            return 0;
        }

        //切换文件时运行
        private int swapp_ActiveDocChangeNotify() { SaveDoc(); swApp.ActiveDocChangeNotify -= swapp_ActiveDocChangeNotify; return 0; }

        // 选择事件
        public static void Selection_Event() { }

        #endregion


    }
}
