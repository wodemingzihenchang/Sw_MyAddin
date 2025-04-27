using System;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Xml;

namespace Sw_toolkit
{
    class MaterialCustomprop
    {
        public static SldWorks swApp = (SldWorks)SwModelDoc.ConnectToSolidWorks();

        /// <summary>
        /// 输入材料自定义属性名，返回属性值
        /// </summary>
        /// <param name="propname"></param>
        /// <returns>材料自定义属性</returns>
        public static string GetMaterialDatabases(ModelDoc2 swDoc ,string propname)
        {
            //获得当前材料
            string material_db = swDoc.MaterialIdName.Split('|')[0];
            string material_name = swDoc.MaterialIdName.Split('|')[1];
            string material_xmlpath = "";

            //获得当前材料库//Console.WriteLine("Material schema pathname = " + swApp.GetMaterialSchemaPathName());
            object[] vMatDBarr = (object[])swApp.GetMaterialDatabases();
            foreach (object item in vMatDBarr)
            {
                //if (item.ToString().Contains(material_db)) { material_xmlpath = item.ToString(); }不能区分大小写
                if (item.ToString().IndexOf(material_db, StringComparison.OrdinalIgnoreCase) >= 0) { material_xmlpath = item.ToString(); }
            }

            //将XML文件加载进来
            XmlDocument doc = new XmlDocument(); doc.Load(material_xmlpath);
            //获取根元素+子元素列表
            XmlElement element_root = doc.DocumentElement;
            XmlNodeList node_lists = element_root.GetElementsByTagName("material");
            foreach (XmlElement element in node_lists)
            {
                //元素名+元素值
                //Console.WriteLine(element.Name + element.GetAttribute("name") + element.Value);
                //判断元素属性名是否符合修改对象
                if (material_name == element.GetAttribute("name"))
                {
                    //所有节点属性元素
                    XmlNodeList element_allprop = element.SelectNodes("custom/prop");
                    if (element_allprop != null)
                    {
                        foreach (XmlNode item in element_allprop)
                        {//对所有属性节点进行判断
                            XmlAttributeCollection prop_attribute = item.Attributes;
                            if (prop_attribute[0].Value==propname)
                            {   //用属性名判断使用哪个属性
                                string s0 = prop_attribute[0].Value; //属性名
                                string s1 = prop_attribute[1].Value; //说明
                                string s2 = prop_attribute[2].Value; //数值
                                string s3 = prop_attribute[3].Value; //单位

                                return s2;
                            }
                           
                        }
                    }
                    else { Console.WriteLine(material_name + "没属性"); }
                }
            }
            return "";
        }

    }
}
