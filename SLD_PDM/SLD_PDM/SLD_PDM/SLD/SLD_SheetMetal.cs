// System
using System;

// SolidWorks
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace SLD_PDM.SLD
{
    public class SLD_SheetMetal
    {
        public SldWorks swApp;
        public IModelDoc2 swModel;

        public SLD_SheetMetal(SldWorks sldWorksApp)
        {
            swApp = sldWorksApp;
        }

        /// <summary>
        /// Verifica se o modelo é uma chapa metálica.
        /// </summary>
        public bool IsSheetMetal(IModelDoc2 model)
        {
            try
            {
                Feature feature = (Feature)model.FirstFeature();

                while (feature != null)
                {
                    if (feature.GetTypeName2().ToUpper() == "SHEETMETAL")
                    {
                        return true; // Retorna verdadeiro ao encontrar a feature SheetMetal
                    }
                    feature = (Feature)feature.GetNextFeature();
                }
            }
            catch (Exception ex)
            {
                LOG.GravarLog($"{nameof(SLD_SheetMetal).ToUpper()}:{nameof(IsSheetMetal)}",
                    "ERRO - Ao verificar se o item é SHEETMETAL. Ative o DEBUG para mais detalhes.", ex);
            }

            return false; // Retorna falso se não encontrar a feature
        }

        /// <summary>
        /// Planifica a chapa metálica, ativando a feature "Flat-Pattern".
        /// </summary>
        public bool PlanificarChapa(IModelDoc2 model)
        {
            try
            {
                Feature feature = (Feature)model.FirstFeature();

                while (feature != null)
                {
                    // Verifica se a feature é do tipo "FlatPattern" (Planificação de chapa metálica)
                    if (feature.GetTypeName2().ToUpper() == "FLATPATTERN")
                    {
                        // Descomprime (planifica) a chapa metálica
                        feature.SetSuppression2((int)swFeatureSuppressionAction_e.swUnSuppressFeature, 2, null);
                        return true; // Planificação bem-sucedida
                    }

                    feature = (Feature)feature.GetNextFeature();
                }
            }
            catch (Exception ex)
            {
                LOG.GravarLog($"{nameof(SLD_SheetMetal).ToUpper()}:{nameof(PlanificarChapa)}",
                    "ERRO - Ao planificar a chapa metálica. Ative o DEBUG para mais detalhes.", ex);
            }

            return false; // Retorna falso caso a planificação não seja possível
        }

        /// <summary>
        /// Desplanifica a chapa metálica, suprimindo a feature "Flat-Pattern".
        /// </summary>
        public void DesplanificarChapa(IModelDoc2 model)
        {
            try
            {
                Feature feature = (Feature)model.FirstFeature();

                while (feature != null)
                {
                    // Verifica se a feature é do tipo "FlatPattern" (Planificação de chapa metálica)
                    if (feature.GetTypeName2().ToUpper() == "FLATPATTERN")
                    {
                        // Suprime (desplanifica) a chapa metálica
                        feature.SetSuppression2((int)swFeatureSuppressionAction_e.swSuppressFeature, 2, null);
                        return; // Desplanificação bem-sucedida
                    }

                    feature = (Feature)feature.GetNextFeature();
                }
            }
            catch (Exception ex)
            {
                LOG.GravarLog($"{nameof(SLD_SheetMetal).ToUpper()}:{nameof(DesplanificarChapa)}",
                    "ERRO - Ao desplanificar a chapa metálica. Ative o DEBUG para mais detalhes.", ex);
            }
        }

        #region  CutListFolder
        public string getComprimento_CutListFolder(IModelDoc2 model)
        {
            try
            {
                string comprimento = "";
                Feature swFeat = (Feature)model.FirstFeature();

                while ((swFeat != null))
                {
                    if (swFeat.GetTypeName2() == "CutListFolder")
                    {
                        if (!swFeat.ExcludeFromCutList && !swFeat.IsSuppressed())
                        {
                            comprimento = getPropridedadeDaCutList(swFeat, "selecionado da Caixa delimitadora", "Bounding Box Length");
                            break;
                        }
                    }
                    swFeat = (Feature)swFeat.GetNextFeature();
                }

                return comprimento;
            }
            catch (Exception ex)
            {
                LOG.GravarLog($"{nameof(SLD_SheetMetal).ToUpper()}:{nameof(getComprimento_CutListFolder)}",
                  "ERRO - Ao obter o comprimento da CutListFolder. Ative o DEBUG para mais detalhes.", ex);
                throw; // Relança a exceção
            }
        }
        public string getLargura_CutListFolder(IModelDoc2 model)
        {
            try
            {
                string largura = "";
                Feature swFeat = (Feature)model.FirstFeature();

                while ((swFeat != null))
                {
                    if (swFeat.GetTypeName2() == "CutListFolder")
                    {
                        if (!swFeat.ExcludeFromCutList && !swFeat.IsSuppressed())
                        {
                            largura = getPropridedadeDaCutList(swFeat, "Largura da Caixa delimitadora", "Bounding Box Width");
                            break;
                        }
                    }
                    swFeat = (Feature)swFeat.GetNextFeature();
                }

                return largura;
            }
            catch (Exception ex)
            {
                LOG.GravarLog($"{nameof(SLD_SheetMetal).ToUpper()}:{nameof(getLargura_CutListFolder)}",
                  "ERRO - Ao obter a largura da CutListFolder. Ative o DEBUG para mais detalhes.", ex);
                throw; // Relança a exceção
            }
        }
        public string getEspessura_CutListFolder(IModelDoc2 model)
        {
            try
            {
                string espessura = "";
                Feature swFeat = (Feature)model.FirstFeature();

                while ((swFeat != null))
                {
                    if (swFeat.GetTypeName2() == "CutListFolder")
                    {
                        if (!swFeat.ExcludeFromCutList && !swFeat.IsSuppressed())
                        {
                            espessura = getPropridedadeDaCutList(swFeat, "Espessura da Chapa metálica", "Sheet Metal Thickness");
                            break;
                        }
                    }
                    swFeat = (Feature)swFeat.GetNextFeature();
                }

                return espessura;
            }
            catch (Exception ex)
            {
                LOG.GravarLog($"{nameof(SLD_SheetMetal).ToUpper()}:{nameof(getEspessura_CutListFolder)}",
                  "ERRO - Ao obter a espessura da CutListFolder. Ative o DEBUG para mais detalhes.", ex);
                throw; // Relança a exceção
            }
        }
        private string getPropridedadeDaCutList(Feature swFeat, string propriedadeBr, string propriedadeEng)
        {
            try
            {
                string valorPropriedade = "";
                string CustomPropVal = "";
                string CustomPropResolvedVal = "";

                CustomPropertyManager CustomPropMgr = (CustomPropertyManager)swFeat.CustomPropertyManager;
                string[] vCustomPropNames = (string[])CustomPropMgr.GetNames();

                if ((vCustomPropNames != null))
                {
                    for (int ii = 0; ii <= (vCustomPropNames.Length - 1); ii++)
                    {
                        string CustomPropName = vCustomPropNames[ii];

                        if (CustomPropName.Contains(propriedadeBr) || CustomPropName.Contains(propriedadeEng))
                        {
                            CustomPropMgr.Get2(CustomPropName, out CustomPropVal, out CustomPropResolvedVal);

                            valorPropriedade = CustomPropResolvedVal.Replace(".", ",");
                            break;
                        }
                    }
                }

                return valorPropriedade;
            }
            catch (Exception ex)
            {
                LOG.GravarLog($"{nameof(SLD_SheetMetal).ToUpper()}:{nameof(getPropridedadeDaCutList)}",
                    "ERRO - Ao obter a propriedade da CutList. Ative o DEBUG para mais detalhes.", ex);
                throw; // Relança a exceção
            }
        }
        #endregion
    }
}
