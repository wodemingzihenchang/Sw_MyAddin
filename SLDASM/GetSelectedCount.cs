using Microsoft.VisualBasic;
using SolidWorks.Interop.sldworks;
using System.Windows.Forms;


namespace Sw_MyAddin.SLDASM
{
    class GetSelectedCount
    {
        /// <summary>
        /// 统计所选零件的数量
        /// </summary>
        public static void Function(ISldWorks swApp)
        {
            ModelDoc2 swDoc = swApp.ActiveDoc;
            if (swDoc != null)
            {
                //选择管理器对象
                SelectionMgr swSelMgr = swDoc.SelectionManager;
                //获取SOLIDWORKS主框架。
                Frame swFrame = swApp.Frame();
                //在状态栏左侧的主状态栏区域中显示文本字符串。
                //GetSelectedObjectCount2(-1)方法为获取被选择对象的数量,-1表示所有对象
                swFrame.SetStatusBarText("选择了 " + swSelMgr.GetSelectedObjectCount2(-1) + " 个零部件");
            }
            else { MessageBox.Show("请打开装配体"); }
        }
        public static void GetSelected_unique(SldWorks swApp) //统计零件（唯一）
        {
            ModelDoc2 swDoc = swApp.ActiveDoc;
            AssemblyDoc swAssy = (AssemblyDoc)swDoc;
            if (swAssy != null)
            {
                SelectionMgr swSelMgr = swAssy.GetSelectedMember(); //选择管理器对象
                Collection swCompsColl = new Collection(); //用于存放选中组件的集合

                //GetSelectedObjectCount2(-1)方法为获取被选择对象的数量,-1表示所有对象
                for (int i = 0; i < swSelMgr.GetSelectedObjectCount2(-1); i++)
                {//按照给定的序号，获取选定零部件
                    Component2 swComp = swSelMgr.GetSelectedObjectsComponent2(i);

                    if (swComp != null)
                    {//仅获取唯一组件，不重复统计 'get only unique components
                    }
                }
                //获取SOLIDWORKS主框架。
                Frame swFrame = swApp.Frame();
                //在状态栏左侧的主状态栏区域中显示文本字符串。
                swFrame.SetStatusBarText("Selected " + swCompsColl.Count + " component(s)");
            }
            else { MessageBox.Show("请打开装配体,Please open assembly"); }
        }
        public static bool Contains(Collection coll, object item) //判断是否为同一个组件
        {
            for (int i = 0; i < coll.Count; i++)
            {//遍历集合中的组件，判断是否为同一个组件
                if (coll[i] == item) { return false; }
            }
            return true;
        }
    }
}
