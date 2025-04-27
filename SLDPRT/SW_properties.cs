using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Windows.Forms;

namespace Sw_toolkit
{
    class SW_properties
    {
        static object vPropNamesObject = null; //属性名
        static object vPropTypes = null;       //属性类型
        static object vPropValues = null;      //属性值

        public static void GetPropvalue(SldWorks swApp)//获取自定义属性string filenames, int Doctype, int times
        {
            //获得自定义属性对象
            ModelDoc2 swModel = swApp.ActiveDoc;
            CustomPropertyManager cusPropMgr = swModel.Extension.CustomPropertyManager[""];
            //获取自定义属性内容，ref用以返回属性名，类型，值
            cusPropMgr.GetAll2(ref vPropNamesObject, ref vPropTypes, ref vPropValues, swCustomInfoGetResult_e.swCustomInfoGetResult_NotPresent);
            object[] vPropNames = (object[])vPropNamesObject;
            string[] propValues = (string[])vPropValues;

            //写入内容
            if (vPropNames == null) { return; }
            for (int i = 0; i < vPropNames.Length; i++)
            {
                Console.WriteLine(vPropNames[i]);
                Console.WriteLine(propValues[i]);
            }
        }

        public static void SetPropertiesToExcel(SldWorks swApp, object[] vPropNames, string[] propValues)//设置自定义属性
        {
            //获得自定义属性对象
            ModelDoc2 swModel = swApp.ActiveDoc;
            CustomPropertyManager cusPropMgr = swModel.Extension.CustomPropertyManager[""];

            for (int i = 0; i < vPropNames.Length; i++)
            {
                if (vPropNames[i] != null)
                {
                    string PropertyName = vPropNames[i].ToString();  // 获得属性名
                    string PropertyValue = propValues[i].ToString(); // 获得属性值
                    cusPropMgr.Add3(PropertyName, (int)swCustomInfoType_e.swCustomInfoText, PropertyValue, (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);//输入属性内容
                }
            }
            MessageBox.Show("完成");
        }
        public static void GetPartData(SldWorks swApp)//
        {
            // 获取当前活动的文档
            ModelDoc2 swModel = swApp.ActiveDoc;

            //请先打开零件: ..\TemplateModel\clamp1.sldprt

            if (swApp != null)
            {
                swModel = (ModelDoc2)swApp.ActiveDoc; //当前零件

                //获取通用属性值
                string project = swModel.GetCustomInfoValue("", "Project");

                swModel.DeleteCustomInfo2("", "Qty"); //删除指定项
                swModel.AddCustomInfo3("", "Qty", 30, "1"); //增加通用属性值

                var ConfigNames = (string[])swModel.GetConfigurationNames(); //所有配置名称

                Configuration swConfig = null;

                foreach (var configName in ConfigNames)//遍历所有配置
                {
                    swConfig = (Configuration)swModel.GetConfigurationByName(configName);

                    var manger = swModel.Extension.CustomPropertyManager[configName];
                    //删除当前配置中的属性
                    manger.Delete2("Code");
                    //增加一个属性到些配置
                    manger.Add3("Code", (int)swCustomInfoType_e.swCustomInfoText, "A-" + configName, (int)swCustomPropertyAddOption_e.swCustomPropertyReplaceValue);
                    //获取此配置中的Code属性
                    string tempCode = manger.Get("Code");
                    //获取此配置中的Description属性

                    var tempDesc = manger.Get("Description");
                    Console.WriteLine("  Name of configuration  ---> " + configName + " Desc.=" + tempCode);
                }

            }
        }
    }
}
