using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.IO;

namespace Sw_MyAddin.SLDDRW
{
    class Sw_RenameDrawing
    {
        public static void Rename(ISldWorks swApp)
        {
            // 获取零件文件的路径和名称
            ModelDoc2 swModel = swApp.ActiveDoc;
            string partPath = swModel.GetPathName();
            string partDirectory = Path.GetDirectoryName(partPath);
            string partName = Path.GetFileNameWithoutExtension(partPath);
            string partExtension = Path.GetExtension(partPath);

            // 查找同名工程图文件
            string drawingPath = Path.Combine(partDirectory, partName + ".SLDDRW");

            if (System.IO.File.Exists(drawingPath))
            {
                // 提示用户输入新名称          
                string newName = Microsoft.VisualBasic.Interaction.InputBox("请输入新的零件名称（不带扩展名）", partName, partName);

                if (!string.IsNullOrEmpty(newName))
                {
                    string newPartPath = Path.Combine(partDirectory, newName + partExtension);// 新零件文件路径
                    string newDrawingPath = Path.Combine(partDirectory, newName + ".SLDDRW");// 新工程图文件路径
                    try
                    {
                        // 打开工程图
                        swApp.OpenDoc(drawingPath, 3);
                        swApp.ActivateDoc(partPath);

                        // 重命名零部件
                        swModel.Extension.SelectByID2(Path.GetFileName(partPath), "COMPONENT", 0, 0, 0, false, 0, null, 0);
                        int i = swModel.Extension.RenameDocument(newName); Console.WriteLine(i);
                        swModel.Save();
                        swApp.CloseDoc(newPartPath);

                        //工程图
                        swApp.ActivateDoc(drawingPath);
                        ModelDoc2 swModel_Drw = swApp.ActiveDoc;
                        swModel_Drw.Save();
                        swApp.CloseDoc(drawingPath);

                        // 重命名零件文件&工程图文件
                        //System.IO.File.Move(partPath, newPartPath);
                        System.IO.File.Move(drawingPath, newDrawingPath);
                    }
                    catch (Exception ex) { Console.WriteLine($"重命名失败：{ex.Message}"); }
                }
                else { Console.WriteLine("未输入新名称。"); }
            }
            else { Console.WriteLine("未找到与零件同名的工程图文件。"); }
        }
    }
}
