using System;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace Sw_toolkit
{
    class SW_packAndGo
    {
        private static SldWorks swApp = (SldWorks)SwModelDoc.ConnectToSolidWorks();
        private static ModelDoc2 swDoc;
        private void Debug_Click(object sender, EventArgs e)
        {
            swDoc = swApp.ActiveDoc;
            ModelDocExtension extension = swDoc.Extension;
            PackAndGo packAndGo = extension.GetPackAndGo();

            //保存到目标文件夹的根目录。
            packAndGo.FlattenToSingleFolder = true;
            //包括内容
            packAndGo.IncludeDrawings = true;
            packAndGo.IncludeSimulationResults = true;
            packAndGo.IncludeSuppressed = true;
            packAndGo.IncludeToolboxComponents = true;
            //设置前后缀
            packAndGo.AddPrefix = "前缀";
            packAndGo.AddSuffix = "后缀";
            //设置路径
            packAndGo.SetSaveToName(true, @"D:\Mywork\功能测试\01-SW API\#Sw_旧版本\新建文件夹\新建文件夹"); Console.WriteLine(packAndGo.GetSaveToName());
            //

            int[] vs = extension.SavePackAndGo(packAndGo);//0打包成功；1用户输入不正确；2文件存在；3保存空文件；4保存错误；
            foreach (int item in vs)
            {
                Console.WriteLine(item); ;
            }
        }
        public void Mainss()
        {
            ModelDoc2 swModelDoc = default(ModelDoc2);
            ModelDocExtension swModelDocExt = default(ModelDocExtension);
            PackAndGo swPackAndGo = default(PackAndGo);
            string openFile = null;
            bool status = false;
            int warnings = 0;
            int errors = 0;
            int i = 0;
            int namesCount = 0;
            string myPath = null;
            int[] statuses = null;

            // Open assembly
            openFile = "C:\\Users\\Public\\Documents\\SOLIDWORKS\\SOLIDWORKS 2018\\samples\\tutorial\\advdrawings\\handle.sldasm";
            swModelDoc = (ModelDoc2)swApp.OpenDoc6(openFile, (int)swDocumentTypes_e.swDocASSEMBLY, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref errors, ref warnings);
            swModelDocExt = (ModelDocExtension)swModelDoc.Extension;

            // Get Pack and Go object
            Console.WriteLine("Pack and Go");
            swPackAndGo = (PackAndGo)swModelDocExt.GetPackAndGo();

            // Get number of documents in assembly
            namesCount = swPackAndGo.GetDocumentNamesCount();
            Console.WriteLine("  Number of model documents: " + namesCount);


            // Include any drawings, SOLIDWORKS Simulation results, and SOLIDWORKS Toolbox components
            swPackAndGo.IncludeDrawings = true;
            Console.WriteLine(" Include drawings: " + swPackAndGo.IncludeDrawings);
            swPackAndGo.IncludeSimulationResults = true;
            Console.WriteLine(" Include SOLIDWORKS Simulation results: " + swPackAndGo.IncludeSimulationResults);
            swPackAndGo.IncludeToolboxComponents = true;
            Console.WriteLine(" Include SOLIDWORKS Toolbox components: " + swPackAndGo.IncludeToolboxComponents);

            // Get current paths and filenames of the assembly's documents
            object fileNames;
            object[] pgFileNames = new object[namesCount - 1];
            status = swPackAndGo.GetDocumentNames(out fileNames);
            pgFileNames = (object[])fileNames;

            Console.WriteLine("");
            Console.WriteLine("  Current path and filenames: ");
            if ((pgFileNames != null))
            {
                for (i = 0; i <= pgFileNames.GetUpperBound(0); i++)
                {
                    Console.WriteLine("    The path and filename is: " + pgFileNames[i]);
                }
            }

            // Get current save-to paths and filenames of the assembly's documents
            object pgFileStatus;
            status = swPackAndGo.GetDocumentSaveToNames(out fileNames, out pgFileStatus);
            pgFileNames = (object[])fileNames;
            Console.WriteLine("");
            Console.WriteLine("  Current default save-to filenames: ");
            if ((pgFileNames != null))
            {
                for (i = 0; i <= pgFileNames.GetUpperBound(0); i++)
                {
                    Console.WriteLine("   The path and filename is: " + pgFileNames[i]);
                }
            }

            // Set folder where to save the files
            myPath = "C:\\temp\\";
            status = swPackAndGo.SetSaveToName(true, myPath);

            // Flatten the Pack and Go folder structure; save all files to the root directory
            swPackAndGo.FlattenToSingleFolder = true;

            // Add a prefix and suffix to the filenames
            swPackAndGo.AddPrefix = "SW_";
            swPackAndGo.AddSuffix = "_PackAndGo";

            // Verify document paths and filenames after adding prefix and suffix
            object getFileNames;
            object getDocumentStatus;
            string[] pgGetFileNames = new string[namesCount - 1];

            status = swPackAndGo.GetDocumentSaveToNames(out getFileNames, out getDocumentStatus);
            pgGetFileNames = (string[])getFileNames;
            Console.WriteLine("");
            Console.WriteLine("  My Pack and Go path and filenames after adding prefix and suffix: ");
            for (i = 0; i <= namesCount - 1; i++)
            {
                Console.WriteLine("    My path and filename is: " + pgGetFileNames[i]);
            }

            // Pack and Go
            statuses = (int[])swModelDocExt.SavePackAndGo(swPackAndGo);

        }

    }
}
