using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorks.Interop.swpublished;
using System;
using System.Runtime.InteropServices;
using System.Diagnostics;

namespace Sw_toolkit
{
    [ComVisibleAttribute(true)]
    public class clsPropMgr : PropertyManagerPage2Handler9
    {
        //属性管理器页面所需的控件对象Control objects required for the PropertyManager page 
        PropertyManagerPage2 pm_Page;
        PropertyManagerPageGroup pm_Group;
        PropertyManagerPageSelectionbox pm_Selection;
        PropertyManagerPageSelectionbox pm_Selection2;
        PropertyManagerPageLabel pm_Label;
        PropertyManagerPageCombobox pm_Combo;
        PropertyManagerPageListbox pm_List;
        PropertyManagerPageNumberbox pm_Number;
        PropertyManagerPageOption pm_Radio;
        PropertyManagerPageSlider pm_Slider;
        PropertyManagerPageTab pm_Tab;
        PropertyManagerPageButton pm_Button;
        PropertyManagerPageButton pm_Button2;
        PropertyManagerPageBitmapButton pm_BMPButton;
        PropertyManagerPageBitmapButton pm_BMPButton2;
        PropertyManagerPageBitmap pm_Bitmap;
        PropertyManagerPageActiveX pm_ActiveX;

        //页面中的每个控件都需要一个惟一的IDEach control in the page needs a unique ID 
        const int GroupID = 1;
        const int LabelID = 2;
        const int SelectionID = 3;
        const int ComboID = 4;
        const int ListID = 5;
        const int Selection2ID = 6;
        const int NumberID = 7;
        const int RadioID = 8;
        const int SliderID = 9;
        const int TabID = 10;
        const int ButtonID = 11;
        const int ButtonID2 = 17;
        const int BMPButtonID = 12;
        const int BMPButtonID2 = 13;
        const int BitmapID = 14;
        const int ActiveXID = 15;

        public void Show() { pm_Page.Show2(0); }


        //在创建新实例时运行以下代码The following runs when a new instance//of the class is created
        public clsPropMgr(SldWorks swApp)
        {
            string caption = "";
            int options = (int)swAddControlOptions_e.swControlOptions_Visible + (int)swAddControlOptions_e.swControlOptions_Enabled;
            string tip = "提示：";

            int longerrors = 0;
            string[] listItems = new string[4];

            ////创建页面Create the PropertyManager page,属性管理器按钮类型
            options
                = (int)swPropertyManagerButtonTypes_e.swPropertyManager_OkayButton
                + (int)swPropertyManagerButtonTypes_e.swPropertyManager_CancelButton
                + (int)swPropertyManagerPageOptions_e.swPropertyManagerOptions_LockedPage
                + (int)swPropertyManagerPageOptions_e.swPropertyManagerOptions_PushpinButton;
            pm_Page = (PropertyManagerPage2)swApp.CreatePropertyManagerPage("属性页名称", (int)options, this, ref longerrors);

            //确保页面被正确创建Make sure that the page was created properly
            if (longerrors == (int)swPropertyManagerPageStatus_e.swPropertyManagerPage_Okay)
            {
                //菜单栏Add a tab
                pm_Tab = pm_Page.AddTab(TabID, "菜单名", "", 0);

                //控件组Add a group box to the tab
                pm_Group = (PropertyManagerPageGroup)pm_Tab.AddGroupBox(GroupID, "Controls", options);

                //选择框Add selection boxes,AddControl2(ID,类型,名称,对齐,选项,提示)
                pm_Selection = (PropertyManagerPageSelectionbox)pm_Group.AddControl2(SelectionID, (int)swPropertyManagerPageControlType_e.swControlType_Selectionbox, caption, (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_Indent, (int)options, "Select an edge, face, vertex, solid body, or a component");
                pm_Selection2 = (PropertyManagerPageSelectionbox)pm_Group.AddControl2(Selection2ID, (int)swPropertyManagerPageControlType_e.swControlType_Selectionbox, caption, (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_Indent, (int)options, "Select an edge, face, vertex, solid body, or a component");

                swSelectType_e[] filters = new swSelectType_e[7];
                filters[0] = swSelectType_e.swSelEDGES;
                filters[1] = swSelectType_e.swSelREFEDGES;
                filters[2] = swSelectType_e.swSelFACES;
                filters[3] = swSelectType_e.swSelVERTICES;
                filters[4] = swSelectType_e.swSelSOLIDBODIES;
                filters[5] = swSelectType_e.swSelCOMPONENTS;
                filters[6] = swSelectType_e.swSelCOMPSDONTOVERRIDE;

                object filterObj = null; filterObj = filters;

                //选择框的样式（）
                pm_Selection.SingleEntityOnly = false;                  //获取或设置此选择框是用于单个项还是多个项
                pm_Selection.AllowMultipleSelectOfSameEntity = true;    //获取或设置是否可以在此选择框中多次选择同一实体。
                pm_Selection.Height = 50;                               //获取和设置此选择框的高度。
                pm_Selection.SetSelectionFilters(filterObj);            //定义用户可以为此选择框选择的对象类型。

                pm_Selection2.SingleEntityOnly = false;
                pm_Selection2.AllowMultipleSelectOfSameEntity = true;
                pm_Selection2.Height = 50;
                pm_Selection2.SetSelectionFilters(filterObj);

                //下拉框Add a combo box
                pm_Combo = (PropertyManagerPageCombobox)pm_Group.AddControl2(ComboID, (int)swPropertyManagerPageControlType_e.swControlType_Combobox, caption, (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_Indent, (int)options, tip);
                if ((pm_Combo != null))
                { //下拉框的样式（）
                    pm_Combo.Height = 50;           //获取和设置控件高度。
                    listItems[0] = "Value 1";
                    listItems[1] = "Value 2";
                    listItems[2] = "Value 3";
                    listItems[3] = "Value 4";
                    pm_Combo.AddItems(listItems);   //设置此控件数据。
                    pm_Combo.CurrentSelection = 0;  //获取或设置当前控件选定的项。
                }

                //列表Add a list box
                pm_List = (PropertyManagerPageListbox)pm_Group.AddControl2(ListID, (int)swPropertyManagerPageControlType_e.swControlType_Listbox, caption, (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_Indent, (int)options, " Multi - select values in the list box");

                if ((pm_List != null))
                {//列表的样式（）
                    pm_List.Style = (int)swPropMgrPageListBoxStyle_e.swPropMgrPageListBoxStyle_MultipleItemSelect;
                    pm_List.Height = 50;
                    listItems[0] = "Value 1";
                    listItems[1] = "Value 2";
                    listItems[2] = "Value 3";
                    listItems[3] = "Value 4";
                    pm_List.AddItems(listItems);
                    pm_List.SetSelectedItem(1, true); //设置在启用多项选择的列表框中是选择还是清除项目
                }

                //文本Add a label
                pm_Label = (PropertyManagerPageLabel)pm_Group.AddControl2(LabelID, (int)swPropertyManagerPageControlType_e.swControlType_Label, "Label", (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_LeftEdge, options, "");

                //滑块调节Add a slider
                pm_Slider = (PropertyManagerPageSlider)pm_Group.AddControl2(SliderID, (int)swPropertyManagerPageControlType_e.swControlType_Slider, "Slider", (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_LeftEdge, options, "Slide");

                //单选按钮Add a radio button
                pm_Radio = (PropertyManagerPageOption)pm_Group.AddControl2(RadioID, (int)swPropertyManagerPageControlType_e.swControlType_Option, "Radio button", (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_LeftEdge, options, "Select");

                //数字输入框Add a number box
                pm_Number = (PropertyManagerPageNumberbox)pm_Group.AddControl2(NumberID, (int)swPropertyManagerPageControlType_e.swControlType_Numberbox, "Number box", (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_LeftEdge, options, "Spin");

                //按钮Add a button
                pm_Button = (PropertyManagerPageButton)pm_Group.AddControl2(ButtonID, (int)swPropertyManagerPageControlType_e.swControlType_Button, "Button", (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_LeftEdge, options, "Click");
                pm_Button2 = (PropertyManagerPageButton)pm_Group.AddControl2(ButtonID2, (int)swPropertyManagerPageControlType_e.swControlType_Button, "Button", (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_LeftEdge, options, "Click");

                //图标按钮Add a bitmap button
                pm_BMPButton = (PropertyManagerPageBitmapButton)pm_Group.AddControl2(BMPButtonID, (int)swPropertyManagerPageControlType_e.swControlType_BitmapButton, "Bitmap button", (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_LeftEdge, options, "Click");
                pm_BMPButton.SetStandardBitmaps((int)swPropertyManagerPageBitmapButtons_e.swBitmapButtonImage_parallel);//SW内置图标使用

                //添加另一个根据计算机分辨率缩放的位图按钮Add another bitmap button that scales with computer's resolution
                pm_BMPButton2 = (PropertyManagerPageBitmapButton)pm_Group.AddControl2(BMPButtonID2, (int)swPropertyManagerPageControlType_e.swControlType_BitmapButton, "Bitmap button", (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_LeftEdge, options, "Click");
                string[] imageList = new string[3];
                string[] imageListMasks = new string[3];
                imageList[0] = "Pathname_to_nxn_image";
                imageList[1] = "Pathname_to_nnxnn_image";
                imageList[2] = "Pathname_to_nnnxnnn_image";
                imageListMasks[0] = "Pathname_to_mask_nxn_image";
                imageListMasks[1] = "Pathname_to_mask_nnxnn_image";
                imageListMasks[2] = "Pathname_to_mask_nnnxnnn_image";
                pm_BMPButton2.SetBitmapsByName3(imageList, imageListMasks); //插入此按钮的指定图像。

                //图片Add a bitmap
                pm_Bitmap = (PropertyManagerPageBitmap)pm_Group.AddControl2(BitmapID, (int)swPropertyManagerPageControlType_e.swControlType_Bitmap, "Bitmap", (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_LeftEdge, options, "Bitmap");
                pm_Bitmap.SetStandardBitmap((int)swBitmapControlStandardTypes_e.swBitmapControl_Volume);

                //添加ActiveX控件Add an ActiveX control
                pm_ActiveX = (PropertyManagerPageActiveX)pm_Group.AddControl2(ActiveXID, (int)swPropertyManagerPageControlType_e.swControlType_ActiveX, "ActiveX", (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_LeftEdge, options, "ActiveX control tip");
                pm_ActiveX.SetClass("ClassID", "LicenseKey");
            }

            else
            {
                //页面创建失败If the page is not created
                System.Windows.Forms.MessageBox.Show("An error occurred while attempting to create the PropertyManager page.");
            }
        }


        #region IPropertyManagerPage2Handler9 Members

        void IPropertyManagerPage2Handler9.AfterActivation()
        {

        }

        void IPropertyManagerPage2Handler9.AfterClose()
        {

        }

        int IPropertyManagerPage2Handler9.OnActiveXControlCreated(int Id, bool Status)
        {
            Debug.Print("ActiveX control created");
            return 0;
        }

        void IPropertyManagerPage2Handler9.OnButtonPress(int Id)
        {
            {
                if (Id == 11) { Debug.Print("Button1 clicked"); System.Windows.Forms.MessageBox.Show("Button1 clicked"); }
                else if (Id == 17) { Debug.Print("Button2 clicked"); System.Windows.Forms.MessageBox.Show("Button2 clicked"); }
            }
        }

        void IPropertyManagerPage2Handler9.OnCheckboxCheck(int Id, bool Checked)

        {

        }

        void IPropertyManagerPage2Handler9.OnClose(int Reason)
        {
            if (Reason == (int)swPropertyManagerPageCloseReasons_e.swPropertyManagerPageClose_Cancel)
            {//关闭属性页操作
                Debug.Print("Cancel button clicked");
            }
            else if (Reason == (int)swPropertyManagerPageCloseReasons_e.swPropertyManagerPageClose_Okay)
            {//确认属性页操作
                Debug.Print("OK button clicked");
            }
        }

        void IPropertyManagerPage2Handler9.OnComboboxEditChanged(int Id, string Text)

        {

        }

        void IPropertyManagerPage2Handler9.OnComboboxSelectionChanged(int Id, int Item)

        {

        }

        void IPropertyManagerPage2Handler9.OnGainedFocus(int Id)
        {
            short[] varArray = null;
            Debug.Print("Control box " + Id + " gained focus");
            varArray = (short[])pm_List.GetSelectedItems();
            pm_Combo.CurrentSelection = varArray[0];
        }

        void IPropertyManagerPage2Handler9.OnGroupCheck(int Id, bool Checked)

        {

        }

        void IPropertyManagerPage2Handler9.OnGroupExpand(int Id, bool Expanded)

        {

        }

        bool IPropertyManagerPage2Handler9.OnHelp()

        {

            return false;

        }

        bool IPropertyManagerPage2Handler9.OnKeystroke(int Wparam, int Message, int Lparam, int Id)

        {

            return false;

        }

        void IPropertyManagerPage2Handler9.OnListboxSelectionChanged(int Id, int Item)

        {

        }

        void IPropertyManagerPage2Handler9.OnLostFocus(int Id)

        {

            Debug.Print("Control box " + Id + " lost focus");

        }

        bool IPropertyManagerPage2Handler9.OnNextPage()

        {

            return false;

        }

        void IPropertyManagerPage2Handler9.OnNumberboxChanged(int Id, double Value)

        {

            Debug.Print("Number box changed");

        }

        void IPropertyManagerPage2Handler9.OnOptionCheck(int Id)

        {

            Debug.Print("Option selected");

        }

        void IPropertyManagerPage2Handler9.OnPopupMenuItem(int Id)

        {

        }

        void IPropertyManagerPage2Handler9.OnPopupMenuItemUpdate(int Id, ref int retval)

        {

        }

        bool IPropertyManagerPage2Handler9.OnPreview()

        {

            return false;

        }

        bool IPropertyManagerPage2Handler9.OnPreviousPage()

        {

            return false;

        }

        void IPropertyManagerPage2Handler9.OnRedo()

        {

        }

        void IPropertyManagerPage2Handler9.OnSelectionboxCalloutCreated(int Id)

        {

        }

        void IPropertyManagerPage2Handler9.OnSelectionboxCalloutDestroyed(int Id)

        {

        }

        void IPropertyManagerPage2Handler9.OnSelectionboxFocusChanged(int Id)

        {

            Debug.Print("The focus moved to selection box " + Id);

        }

        void IPropertyManagerPage2Handler9.OnSelectionboxListChanged(int Id, int Count)

        {

            pm_Page.SetCursor((int)swPropertyManagerPageCursors_e.swPropertyManagerPageCursors_Advance);
            Debug.Print("The list in selection box " + Id + " changed");

        }

        void IPropertyManagerPage2Handler9.OnSliderPositionChanged(int Id, double Value)

        {

            Debug.Print("Slider position changed");

        }

        void IPropertyManagerPage2Handler9.OnSliderTrackingCompleted(int Id, double Value)

        {

        }

        bool IPropertyManagerPage2Handler9.OnSubmitSelection(int Id, object Selection, int SelType, ref string ItemText)

        {

            // This method must return true for selections to occur
            return true;

        }

        bool IPropertyManagerPage2Handler9.OnTabClicked(int Id)

        {

            return false;

        }

        void IPropertyManagerPage2Handler9.OnTextboxChanged(int Id, string Text)

        {

        }

        void IPropertyManagerPage2Handler9.OnUndo()

        {

        }

        void IPropertyManagerPage2Handler9.OnWhatsNew()

        {

        }

        void IPropertyManagerPage2Handler9.OnListboxRMBUp(int Id, int PosX, int PosY)
        {

        }

        int IPropertyManagerPage2Handler9.OnWindowFromHandleControlCreated(int Id, bool Status)
        {

            return 0;

        }

        void IPropertyManagerPage2Handler9.OnNumberBoxTrackingCompleted(int Id, double Value)
        {

        }
        #endregion
    }

    public class show_PropMgr
    {
        private static SldWorks swApp = (SldWorks)SwModelDoc.ConnectToSolidWorks();
        private static clsPropMgr pm;
        public static void Show()
        {
            //设置文档默认用户首选项值。此方法用于可切换的用户首选项。
            swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swStopDebuggingVstaOnExit, false);
            ModelDoc2 swDoc = swApp.ActiveDoc; if (swDoc == null) { return; }
            //创建PropertyManager类的新实例Create a new instance of the PropertyManager class
            pm = new clsPropMgr(swApp);
            pm.Show();
        }
    }
}