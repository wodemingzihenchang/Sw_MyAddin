using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace Sw_MyAddin
{
    class DeFeatrue
    {
        /// <summary>
        /// 删除零件特征,但保留实体,类似于导出
        /// </summary>
        public static void Function(ISldWorks swApp)
        {
            var swModel = (ModelDoc2)swApp.ActiveDoc;
            if (swModel != null)
            {
                PartDoc part = (PartDoc)swModel;

                // 把备份的实体,用于后面CreateFeatureFromBody生成
                var vBodies = GetBodyCopies(part);

                // 删除当前所有的特征
                SelectAllUserFeature(swModel);
                swModel.Extension.DeleteSelection2((int)swDeleteSelectionOptions_e.swDelete_Children + (int)swDeleteSelectionOptions_e.swDelete_Absorbed);

                // 恢复备份的实体
                for (int i = 0; i < vBodies.Length; i++)
                {
                    part.CreateFeatureFromBody3(vBodies[i], false, (int)swCreateFeatureBodyOpts_e.swCreateFeatureBodySimplify);
                }
            }
        }

        private static Body2[] GetBodyCopies(PartDoc partDoc)// 获取零件实体的备份
        {
            object[] vBodies = partDoc.GetBodies2((int)swBodyType_e.swAllBodies, true);

            Body2[] newBodies = new Body2[vBodies.Length];

            for (int i = 0; i < vBodies.Length; i++)
            {
                var swBody2 = (Body2)vBodies[i];
                newBodies[i] = swBody2.Copy();
            }
            return newBodies;
        }

        private static void SelectAllUserFeature(ModelDoc2 modelDoc2)// 选择所有的特征
        {
            modelDoc2.ClearSelection2(true);

            var swFeature = (Feature)modelDoc2.FirstFeature();

            while (swFeature != null)
            {
                if (swFeature.GetTypeName2() == "OriginProfileFeature" || swFeature.GetTypeName2() == "原点") { }
                else { swFeature.Select2(true, 1); }
                swFeature = swFeature.GetNextFeature();
            }
        }
    }
}
