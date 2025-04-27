using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;

namespace Sw_MyAddin.SLDASM
{
    class Ordering
    {
        /// <summary>
        /// 对装配体里的零部件进行名字排序，注意这里的排序会破坏原有的文件夹归纳。
        /// </summary>
        public static void Function(ISldWorks swApp)
        {
            //检查是否有效的装配体文档打开
            ModelDoc2 swDoc = swApp.ActiveDoc;
            if (swDoc == null || swDoc.GetType() != 2) { swApp.SendMsgToUser("请打开装配体后再试。"); return; }
            //获得所有顶级零部件
            AssemblyDoc swAsm = (AssemblyDoc)swDoc;
            object[] components = swAsm.GetComponents(true);
            string[] sArray = new string[components.Length];
            int[] iArray = new int[components.Length];
            //获得零部件组合
            Component component, component2;
            for (int i = 0; i < components.Length; i++)
            {
                component = (Component)components[i];
                sArray[i] = component.Name; Console.WriteLine(sArray[i]);
                iArray[i] = i;
            }
            //冒泡排序
            string sTemp; int iTemp;
            for (int x = 0; x < components.Length; x++)
            {
                for (int y = x + 1; y < components.Length; y++)
                {
                    if (String.Compare(sArray[x], sArray[y]) == 1)
                    {
                        sTemp = sArray[x];
                        iTemp = iArray[x];
                        sArray[x] = sArray[y];
                        iArray[x] = iArray[y];
                        sArray[y] = sTemp;
                        iArray[y] = iTemp;
                    }
                }
            }
            //反序
            //for (int i = components.Length - 1; i >= 1; i--)
            //{
            //    component = (Component)components[iArray[i - 1]];
            //    component2 = (Component)components[iArray[i]];
            //    //Console.WriteLine(iArray[i-1]);
            //    Console.WriteLine(component.Name2 + " - " + component2.Name2);
            //    swAsm.ReorderComponents(component, component2, 1);//反序
            //    //swAsm.ReorderComponents(component, component2, 2);//正序
            //}
            //正序
            for (int i = 0; i < components.Length - 1; i++)
            {
                component = (Component)components[iArray[i]];
                component2 = (Component)components[iArray[i + 1]];
                //Console.WriteLine(iArray[i-1]);
                Console.WriteLine(component.Name2 + " - " + component2.Name2);
                swAsm.ReorderComponents(component2, component, 1);
            }

        }
        public static void GB_Components(SldWorks swApp)//排序归类GB标准件
        {
            //检查是否有效的装配体文档打开
            ModelDoc2 swDoc = swApp.ActiveDoc;
            if (swDoc == null || swDoc.GetType() != 2) { swApp.SendMsgToUser("请打开装配体后再试。"); return; }
            AssemblyDoc swAsm = (AssemblyDoc)swDoc;

            ///移到文件夹
            Feature feature = swDoc.FeatureManager.InsertFeatureTreeFolder2((int)swFeatureTreeFolderType_e.swFeatureTreeFolder_EmptyBefore);
            feature = swAsm.FeatureByName("标准件");
            if (feature == null) { feature.Name = "标准件"; }

            //获得零部件组合
            object[] components = swAsm.GetComponents(true);
            Component component;
            for (int i = 0; i < components.Length; i++)
            {
                component = (Component)components[i];
                if (component.Name.Contains("GB") || component.Name.Contains("gb"))
                {
                    swAsm.ReorderComponents(components[i], feature, (int)swReorderComponentsWhere_e.swReorderComponents_LastInFolder);
                }
            }
        }
        public static void Ordering_SelectComponents()
        {

        }     //排序所选零件
    }
}