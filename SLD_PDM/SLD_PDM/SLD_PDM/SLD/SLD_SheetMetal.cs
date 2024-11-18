// System
using System;

// SolidWorks
using SolidWorks.Interop.sldworks;

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

        // Verifica se o modelo é chapa metálica
        public bool isSheetMetal(IModelDoc2 model)
        {
            bool sheetMetal = false;
            try
            {
                ModelDocExtension mdExtension = model.Extension;
                Feature f = (Feature)model.FirstFeature();

                while (f != null)
                {
                    if (f.GetTypeName2().ToUpper() == "SHEETMETAL")
                    {
                        sheetMetal = true;
                        break;
                    }
                    f = (Feature)f.GetNextFeature();
                }
            }
            catch (Exception ex)
            {
                LOG.GravarLog($"{nameof(SLD_SheetMetal).ToUpper()}:{nameof(isSheetMetal)}",
                    "ERRO - Ao verificar se o item é SHEETMETAL. Ative o DEBUG para mais detalhes.", ex);
            }
            return sheetMetal;
        }
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
    }
}