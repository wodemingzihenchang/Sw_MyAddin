using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorks.Interop.swpublished;
using SolidWorksTools;
using SolidWorksTools.File;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;


namespace Sw_MyAddin
{
    /// <summary>
    /// Summary description for Sw_MyAddin.
    /// </summary>
    //[Guid("dc757526-2088-4d14-b9ca-195b24787a85"), ComVisible(true)]
    [Guid("F8117953-FF24-40B7-BD35-DA0637C5842C"), ComVisible(true)]
    [SwAddin(Description = "Sw_MyAddin description", Title = "Sw_MyAddin", LoadAtStartup = true)]
    public class SwAddin : ISwAddin
    {
        #region Local Variables ȫ�ֱ���

        ISldWorks swApp = null;
        ICommandManager iCmdMgr = null;
        int addinID = 0;
        BitmapHandler iBmp;
        int registerID;

        public const int mainCmdGroupID = 5;
        public const int mainItemID1 = 0;
        public const int mainItemID2 = 1;
        public const int mainItemID3 = 2;
        public const int flyoutGroupID = 91;

        string[] mainIcons = new string[1];
        string[] icons = new string[1];

        #region Event Handler Variables

        Hashtable openDocs = new Hashtable();
        SolidWorks.Interop.sldworks.SldWorks SwEventPtr = null;
        #endregion

        #region Property Manager Variables
        public UserPMPage ppage = null;
        #endregion


        // Public Properties
        public ISldWorks SwApp { get { return swApp; } }
        public ICommandManager CmdMgr
        {
            get { return iCmdMgr; }
        }

        public Hashtable OpenDocs
        {
            get { return openDocs; }
        }

        #endregion

        #region SolidWorks Registration ע����

        [ComRegisterFunctionAttribute]
        public static void RegisterFunction(Type t)
        {
            #region Get Custom Attribute: SwAddinAttribute
            SwAddinAttribute SWattr = null;
            Type type = typeof(SwAddin);

            foreach (System.Attribute attr in type.GetCustomAttributes(false))
            {
                if (attr is SwAddinAttribute)
                {
                    SWattr = attr as SwAddinAttribute;
                    break;
                }
            }

            #endregion

            try
            {
                Microsoft.Win32.RegistryKey hklm = Microsoft.Win32.Registry.LocalMachine;
                Microsoft.Win32.RegistryKey hkcu = Microsoft.Win32.Registry.CurrentUser;

                string keyname = "SOFTWARE\\SolidWorks\\Addins\\{" + t.GUID.ToString() + "}";
                Microsoft.Win32.RegistryKey addinkey = hklm.CreateSubKey(keyname);
                addinkey.SetValue(null, 0);

                addinkey.SetValue("Description", SWattr.Description);
                addinkey.SetValue("Title", SWattr.Title);

                keyname = "Software\\SolidWorks\\AddInsStartup\\{" + t.GUID.ToString() + "}";
                addinkey = hkcu.CreateSubKey(keyname);
                addinkey.SetValue(null, Convert.ToInt32(SWattr.LoadAtStartup), Microsoft.Win32.RegistryValueKind.DWord);
            }
            catch (System.NullReferenceException nl)
            {
                Console.WriteLine("There was a problem registering this dll: SWattr is null. \n\"" + nl.Message + "\"");
                System.Windows.Forms.MessageBox.Show("There was a problem registering this dll: SWattr is null.\n\"" + nl.Message + "\"");
            }

            catch (System.Exception e)
            {
                Console.WriteLine(e.Message);

                System.Windows.Forms.MessageBox.Show("There was a problem registering the function: \n\"" + e.Message + "\"");
            }
        }

        [ComUnregisterFunctionAttribute]
        public static void UnregisterFunction(Type t)
        {
            try
            {
                Microsoft.Win32.RegistryKey hklm = Microsoft.Win32.Registry.LocalMachine;
                Microsoft.Win32.RegistryKey hkcu = Microsoft.Win32.Registry.CurrentUser;

                string keyname = "SOFTWARE\\SolidWorks\\Addins\\{" + t.GUID.ToString() + "}";
                hklm.DeleteSubKey(keyname);

                keyname = "Software\\SolidWorks\\AddInsStartup\\{" + t.GUID.ToString() + "}";
                hkcu.DeleteSubKey(keyname);
            }
            catch (System.NullReferenceException nl)
            {
                Console.WriteLine("There was a problem unregistering this dll: " + nl.Message);
                System.Windows.Forms.MessageBox.Show("There was a problem unregistering this dll: \n\"" + nl.Message + "\"");
            }
            catch (System.Exception e)
            {
                Console.WriteLine("There was a problem unregistering this dll: " + e.Message);
                System.Windows.Forms.MessageBox.Show("There was a problem unregistering this dll: \n\"" + e.Message + "\"");
            }
        }

        #endregion

        #region ISwAddin Implementation 
        public SwAddin() { }
        public bool ConnectToSW(object ThisSW, int cookie)
        {
            swApp = (SldWorks)ThisSW;
            addinID = cookie;

            //Setup callbacks
            swApp.SetAddinCallbackInfo(0, this, addinID);

            #region ���ù�����
            iCmdMgr = swApp.GetCommandManager(cookie);
            AddCommandMgr();
            #endregion

            #region �����¼�
            //SwEventPtr = (SolidWorks.Interop.sldworks.SldWorks)swApp;
            //openDocs = new Hashtable();
            //AttachEventHandlers();
            SW_Event _Event = new SW_Event((SldWorks)swApp);
            _Event.SaveDoc();
            #endregion

            #region ������������ҳ
            AddPMP();
            #endregion

            return true;
        }

        public bool DisconnectFromSW()
        {
            RemoveCommandMgr();
            RemovePMP();
            //DetachEventHandlers();

            System.Runtime.InteropServices.Marshal.ReleaseComObject(iCmdMgr);
            iCmdMgr = null;
            System.Runtime.InteropServices.Marshal.ReleaseComObject(swApp);
            swApp = null;
            //The addin _must_ call GC.Collect() here in order to retrieve all managed code pointers 
            GC.Collect();
            GC.WaitForPendingFinalizers();

            GC.Collect();
            GC.WaitForPendingFinalizers();

            return true;
        }
        #endregion

        #region UI Methods ��������
        public void AddCommandMgr()
        {
            int cmdGroupErr = 0; bool ignorePrevious = false;
            string Title = "Sw_MyAddin", ToolTip = "Sw_MyAddin";

            if (iBmp == null) { iBmp = new BitmapHandler(); }
            Assembly thisAssembly = System.Reflection.Assembly.GetAssembly(this.GetType());
            //���ID��Ϣ���浽ע����� get the ID information stored in the registry
            object registryIDs;
            bool getDataResult = iCmdMgr.GetGroupDataFromRegistry(mainCmdGroupID, out registryIDs);
            int[] knownIDs = new int[2] { mainItemID1, mainItemID2 };
            //if the IDs don't match, reset the commandGroup
            if (getDataResult) { if (!CompareIDs((int[])registryIDs, knownIDs)) { ignorePrevious = true; } }
            ICommandGroup cmdGroup = iCmdMgr.CreateCommandGroup2(mainCmdGroupID, Title, ToolTip, "", -1, ignorePrevious, ref cmdGroupErr);
            // ��λͼ��ӵ���Ŀ�У�����������ΪǶ����Դ�������ṩλͼ��ֱ��·��
            // ͼ��λ����Ҫ�ŵ�temp�ļ�����
            icons[0] = iBmp.CreateFileFromResourceBitmap("Sw_MyAddin.icon.toolbar64x.png", thisAssembly);
            mainIcons[0] = iBmp.CreateFileFromResourceBitmap("Sw_MyAddin.icon.mainicon_64.png", thisAssembly);
            // ͼ��λ����Ҫ�ŵ�SLDWORK������
            //icons[0] = System.Environment.CurrentDirectory + @"\icon\toolbar64x.png";
            //mainIcons[0] = System.Environment.CurrentDirectory + @"\icon\toolbar64x.png";

            //�½�������
            //cmdGroup = iCmdMgr.CreateCommandGroup2(mainCmdGroupID, Title, ToolTip, "", -1, ignorePrevious, ref cmdGroupErr);
            cmdGroup.IconList = icons;          //��������ͼ��·��
            cmdGroup.MainIconList = mainIcons;  //���ò��ͼ��·��

            #region �������
            List<int> cmd_list = new List<int>();
            //��ť����
            //  cmdIndex0 = cmdGroup.AddCommandItem2("��������", ������λ��0, "�����ͣ˵��", "״̬��˵��", ͼƬλ��1, "������", "", ����ID, ��������);
            int cmdIndex0 = cmdGroup.AddCommandItem2("���ܲ��԰�ť", -1, "���ܲ��԰�ť", "���ܲ��԰�ť", 0, "Debug_Function", "EnableForPartsAndAssemsInDesktop", 0, (int)swCommandItemType_e.swMenuItem); cmd_list.Add(cmdIndex0);
            int cmdIndex1 = cmdGroup.AddCommandItem2("��ʵ���  ��", -1, "��ʵ���  ��", "��ʵ���  ��", 1, "Function1", "", 0, (int)swCommandItemType_e.swMenuItem); cmd_list.Add(cmdIndex1);
            int cmdIndex2 = cmdGroup.AddCommandItem2("�߽���  ͼ", -1, "�߽���  ͼ", "�߽���  ͼ", 2, "Function2", "", 0, (int)swCommandItemType_e.swMenuItem); cmd_list.Add(cmdIndex2);
            int cmdIndex3 = cmdGroup.AddCommandItem2("������ɫ�ֲ�", -1, "������ɫ�ֲ�", "������ɫ�ֲ�", 3, "Function3", "", 0, (int)swCommandItemType_e.swMenuItem); cmd_list.Add(cmdIndex3);
            int cmdIndex4 = cmdGroup.AddCommandItem2("�㲿����  ��", -1, "�㲿����  ��", "�㲿����  ��", 4, "Function4", "", 0, (int)swCommandItemType_e.swMenuItem); cmd_list.Add(cmdIndex4);
            int cmdIndex5 = cmdGroup.AddCommandItem2("����ɵ�����", -1, "���������ý��㲿������ɶ��������", "����ɵ�����", 5, "Function5", "", 0, (int)swCommandItemType_e.swMenuItem); cmd_list.Add(cmdIndex5);
            int cmdIndex6 = cmdGroup.AddCommandItem2("д������", -1, "��װ��������д��ͼ�ŷ�������Ժͷ���ʽ", "д������", 6, "Function6", "", 0, (int)swCommandItemType_e.swMenuItem); cmd_list.Add(cmdIndex6);
            int cmdIndex7 = cmdGroup.AddCommandItem2("��ͼֽ������", -1, "����������͹���ͼ�����ֹ���", "������", 6, "Function7", "", 0, (int)swCommandItemType_e.swMenuItem); cmd_list.Add(cmdIndex7);

            //int cmdIndex1 = cmdGroup.AddCommandItem2("���Կ�", -1, "���Կ�", "���Կ�", 1, "ShowPMP", "EnablePMP", 0, (int)swCommandItemType_e.swMenuItem);
            int cmdIndex20 = cmdGroup.AddCommandItem2("����", -1, "����", "����", 0, "Sw_MyAddin_Help", "", 0, (int)swCommandItemType_e.swMenuItem);
            int cmdIndex21 = cmdGroup.AddCommandItem2("����", -1, "����", "����", 0, "Sw_MyAddin_Setting", "", 0, (int)swCommandItemType_e.swMenuItem);

            iCmdMgr.GetCommandTab(1, "1");//��ȡָ���ĵ����͵�ָ�����������ѡ�
            cmdGroup.HasToolbar = true; //�Ƿ���й�����
            cmdGroup.HasMenu = false;//�Ƿ���в˵�
            cmdGroup.Activate();//����������

            #region //...��ʱ�۵�
            //������
            FlyoutGroup flyGroup = iCmdMgr.CreateFlyoutGroup2(flyoutGroupID, "Dynamic Flyout", "Flyout Tooltip", "Flyout Hint",
            cmdGroup.MainIconList, cmdGroup.IconList, "FlyoutCallback", "FlyoutEnable");
            flyGroup.AddCommandItem("FlyoutCommand 1", "test", 0, "FlyoutCommandItem1", "FlyoutEnableCommandItem1");
            flyGroup.FlyoutType = (int)swCommandFlyoutStyle_e.swCommandFlyoutStyle_Simple;

            bool bResult;
            //������ʹ�ó���
            int[] docTypes = new int[]{(int)swDocumentTypes_e.swDocASSEMBLY,
                                       (int)swDocumentTypes_e.swDocDRAWING,
                                       (int)swDocumentTypes_e.swDocPART};
            foreach (int type in docTypes)
            {
                CommandTab cmdTab;
                cmdTab = iCmdMgr.GetCommandTab(type, Title);

                if (cmdTab != null & !getDataResult | ignorePrevious)//���ѡ����ڣ������Ǻ�����ע�����Ϣ���������������ID���������´���ѡ�������id����ƥ�䣬����ѡ���Ϊ�հ�if tab exists, but we have ignored the registry info (or changed command group ID), re-create the tab.  Otherwise the ids won't matchup and the tab will be blank
                {
                    bool res = iCmdMgr.RemoveCommandTab(cmdTab);
                    cmdTab = null;
                }

                //���cmdTabΪ�գ��������ȼ���(����������֮��)��Ȼ��������ӵ�ѡ��� if cmdTab is null, must be first load (possibly after reset), add the commands to the tabs
                if (cmdTab == null)
                {
                    cmdTab = iCmdMgr.AddCommandTab(type, Title);
                    CommandTabBox cmdBox = cmdTab.AddCommandTabBox();
                    int i = 21;
                    int[] cmdIDs = new int[i + 1];
                    int[] TextType = new int[i + 1];
                    //������1
                    int[] vs = other.Setting.cmd_nums;
                    for (int f = 0; f < vs.Length; f++)
                    {
                        cmdIDs[vs[f]] = cmdGroup.get_CommandID(cmd_list[vs[f]]);
                        TextType[vs[f]] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow;
                    }
                    cmdIDs[i - 1] = cmdGroup.get_CommandID(cmdIndex20);
                    TextType[i - 1] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow;
                    cmdIDs[i] = cmdGroup.get_CommandID(cmdIndex21);
                    TextType[i] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow;
                    //���ɹ�����
                    bResult = cmdBox.AddCommands(cmdIDs, TextType);
                    //����������������������ͼ��λ�ã�
                    cmdTab.AddSeparator(cmdBox, cmdIDs[1]);
                    cmdTab.AddSeparator(cmdBox, cmdIDs[i - 1]);

                    ////������2
                    //CommandTabBox cmdBox1 = cmdTab.AddCommandTabBox();
                    //cmdIDs = new int[1];
                    //TextType = new int[1];                    
                    //cmdIDs[0] = flyGroup.CmdID; //cmdIDs[0] = cmdGroup.ToolbarId;
                    //TextType[0] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow | (int)swCommandTabButtonFlyoutStyle_e.swCommandTabButton_ActionFlyout;
                    //bResult = cmdBox1.AddCommands(cmdIDs, TextType);
                    //cmdTab.AddSeparator(cmdBox1, cmdIDs[0]);
                }
            }
            #region ����������ͼ��
            //�ڲ����ж����������е���˵��д���������ͼ�� Create a third-party icon in the context-sensitive menus of faces in parts
            //Ҫ�鿴�˲˵����Ҽ������������κ��� To see this menu, right click on any face in the part
            //Frame swFrame;

            //swFrame = swApp.Frame();
            //bResult = swFrame.AddMenuPopupIcon3((int)swDocumentTypes_e.swDocPART, (int)swSelectType_e.swSelFACES, "third-party context-sensitive CSharp", addinID, "PopupCallbackFunction", "PopupEnable", "", cmdGroup.MainIconList);

            ////������ע���ݲ˵� create and register the shortcut menu
            //registerID = swApp.RegisterThirdPartyPopupMenu();

            ////�ڿ�ݲ˵��Ķ�����Ӳ˵��жϷ� add a menu break at the top of the shortcut menu
            //bResult = swApp.AddItemToThirdPartyPopupMenu2(registerID, (int)swDocumentTypes_e.swDocPART, "Menu Break", addinID, "", "", "", "", "", (int)swMenuItemType_e.swMenuItemType_Break);

            ////���ݲ˵������������Ŀ add a couple of items to the shortcut menu
            //bResult = swApp.AddItemToThirdPartyPopupMenu2(registerID, (int)swDocumentTypes_e.swDocPART, "Test1", addinID, "TestCallback", "EnableTest", "", "Test1", mainIcons[0], (int)swMenuItemType_e.swMenuItemType_Default);
            //bResult = swApp.AddItemToThirdPartyPopupMenu2(registerID, (int)swDocumentTypes_e.swDocPART, "Test2", addinID, "TestCallback", "EnableTest", "", "Test2", mainIcons[0], (int)swMenuItemType_e.swMenuItemType_Default);

            ////���ݲ˵���ӷָ��� add a separator bar to the shortcut menu
            //bResult = swApp.AddItemToThirdPartyPopupMenu2(registerID, (int)swDocumentTypes_e.swDocPART, "separator", addinID, "", "", "", "", "", (int)swMenuItemType_e.swMenuItemType_Separator);

            ////���ݲ˵������һ�� add another item to the shortcut menu
            //bResult = swApp.AddItemToThirdPartyPopupMenu2(registerID, (int)swDocumentTypes_e.swDocPART, "Test3", addinID, "TestCallback", "EnableTest", "", "Test3", mainIcons[0], (int)swMenuItemType_e.swMenuItemType_Default);

            ////�ڿ�ݲ˵��Ĳ˵��������ͼ�� add an icon to a menu bar of the shortcut menu
            //bResult = swApp.AddItemToThirdPartyPopupMenu2(registerID, (int)swDocumentTypes_e.swDocPART, "", addinID, "TestCallback", "EnableTest", "", "NoOp", mainIcons[0], (int)swMenuItemType_e.swMenuItemType_Default);
            #endregion
            #endregion
        }

        public void RemoveCommandMgr()
        {
            iBmp.Dispose();

            iCmdMgr.RemoveCommandGroup(mainCmdGroupID);
            iCmdMgr.RemoveFlyoutGroup(flyoutGroupID);
        }
        public bool CompareIDs(int[] storedIDs, int[] addinIDs)
        {
            List<int> storedList = new List<int>(storedIDs);
            List<int> addinList = new List<int>(addinIDs);

            addinList.Sort();
            storedList.Sort();

            if (addinList.Count != storedList.Count)
            {
                return false;
            }
            else
            {

                for (int i = 0; i < addinList.Count; i++)
                {
                    if (addinList[i] != storedList[i])
                    {
                        return false;
                    }
                }
            }
            return true;
        }
        public Boolean AddPMP()
        {
            ppage = new UserPMPage(this);
            return true;
        }
        public Boolean RemovePMP()
        {
            ppage = null;
            return true;
        }
        #endregion

        #endregion

        #region UI Callbacks �����
        public void Debug_Function() { Console.WriteLine(System.Environment.CurrentDirectory); SwApp.SendMsgToUser("Hello World"); }
        public void Function1() { SLDPRT.Combined_Entity.Function(swApp); }//����ϲ���ʵ��
        public void Function2() { Bounding.Get_Bounding(swApp); }//�߽��3D��ͼ
        public void Function3() { SLDDRW.SW_Layer.Layered(swApp); }//ͼ����ɫ�ֲ�
        public void Function4() { SLDASM.Ordering.Function(swApp); }//װ�����㲿������
        public void Function5() { }// �������õ�װ���弰���㲿������ɶ�Ӧ�ĵ������ļ������ڲ�ͬ�ļ�����
        public void Function6() { SLDASM.Equation.Function(swApp); }//�������д�����Ժͷ���ʽ
        public void Function7() { SLDDRW.Sw_RenameDrawing.Rename(swApp); }// 
        public void Function8() { }//  
        public void Function9() { }
        public void Function10() { }
        public void PopupCallbackFunction() { swApp.ShowThirdPartyPopupMenu(registerID, 500, 500); }
        public void ShowPMP() { if (ppage != null) ppage.Show(); }
        public void FlyoutCallback()
        {
            FlyoutGroup flyGroup = iCmdMgr.GetFlyoutGroup(flyoutGroupID);
            flyGroup.RemoveAllCommandItems();

            flyGroup.AddCommandItem(System.DateTime.Now.ToLongTimeString(), "test", 0, "FlyoutCommandItem1", "FlyoutEnableCommandItem1");

        }
        public int FlyoutEnable()
        {
            return 1;
        }
        public void FlyoutCommandItem1()
        {
            swApp.SendMsgToUser("Flyout command 1");
        }
        public int FlyoutEnableCommandItem1()
        {
            return 1;
        }

        public void Sw_MyAddin_Help() { System.Diagnostics.Process.Start("https://www.bilibili.com/video/BV1uCw9ecEjp"); }
        public void Sw_MyAddin_Setting() { other.Setting setting = new other.Setting(); setting.Show(); }
        #endregion
    }
}
