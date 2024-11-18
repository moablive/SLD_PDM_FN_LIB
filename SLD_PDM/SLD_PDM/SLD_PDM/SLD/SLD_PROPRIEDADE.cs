// System
using System;

// Project
using SLD_PDM.SLD.MODEL;

// SolidWorks
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace SLD_PDM.SLD
{
    public class SLD_PROPRIEDADE
    {
        public SldWorks swApp;
        public IModelDoc2 swModel;

        public SLD_PROPRIEDADE(SldWorks sldWorksApp)
        {
            swApp = sldWorksApp;
        }

        public void SetPropriedade(IModelDoc2 model, string valor, string config, string propName)
        {
            try
            {
                var swModelDocExt = model?.Extension;
                var swCustProp = swModelDocExt?.get_CustomPropertyManager(config);

                if (swCustProp == null)
                    throw new Exception("Gerenciador de propriedade personalizado não encontrado.");

                // Tenta deletar a propriedade antes de adicioná-la
                try
                {
                    swCustProp.Delete(propName);
                }
                catch (Exception ex)
                {
                    LOG.GravarLog($"{nameof(SLD_PROPRIEDADE).ToUpper()}:{nameof(SetPropriedade)}",
                        "ERRO - Ao deletar propriedade existente. Ative o DEBUG para mais detalhes.", ex);
                }

                swCustProp.Add3(propName, (int)swCustomInfoType_e.swCustomInfoText, valor, (int)swCustomPropertyAddOption_e.swCustomPropertyReplaceValue);
            }
            catch (Exception ex)
            {
                LOG.GravarLog($"{nameof(SLD_PROPRIEDADE).ToUpper()}:{nameof(SetPropriedade)}",
                    "ERRO - Ao configurar propriedade. Ative o DEBUG para mais detalhes.", ex);
            }
        }
        public string GetPropriedade(string prop, string Config, IModelDoc2 model)
        {
            try
            {
                ModelDocExtension swModelDocExt = model.Extension;
                CustomPropertyManager swCustProp = swModelDocExt.get_CustomPropertyManager(Config);
                bool status = swCustProp.Get4(prop, false, out _, out string swPropAtual);

                if (status)
                {
                    return swPropAtual;
                }
                else
                {
                    LOG.GravarLog($"{nameof(SLD_PROPRIEDADE).ToUpper()}:{nameof(GetPropriedade)}",
                        "ERRO - Propriedade não encontrada ou vazia. Verifique o nome e a configuração da propriedade.");
                    return string.Empty;
                }
            }
            catch (Exception ex)
            {
                LOG.GravarLog($"{nameof(SLD_PROPRIEDADE).ToUpper()}:{nameof(GetPropriedade)}",
                    "ERRO - Ao obter a propriedade personalizada. Ative o DEBUG para mais detalhes.", ex);
                return string.Empty;
            }
        }


        public objCAIXADELIMITADORA ObtemCaixaDelimitadora(IModelDoc2 model, string conf)
        {
            try
            {
                int sStatus = 0;
                ModelDocExtension swModelDocExt = model.Extension;

                // Insere o recurso da caixa delimitadora
                model.FeatureManager.InsertGlobalBoundingBox(1, false, false, out sStatus);

                // Obtém as propriedades da caixa delimitadora usando o método atualizado
                string comp = GetPropriedade("Comprimento total da caixa delimitadora", conf, model);
                string larg = GetPropriedade("Largura total da caixa delimitadora", conf, model);
                string espess = GetPropriedade("Espessura total da caixa delimitadora", conf, model);

                // Cria o objeto caixa delimitadora com as propriedades obtidas
                var caixaDelimitadora = new objCAIXADELIMITADORA
                {
                    comp = comp,
                    larg = larg,
                    espess = espess
                };

                // Loop nas features para selecionar e apagar o recurso WELDMENT
                Feature f = (Feature)model.FirstFeature();

                while (f != null)
                {
                    if (f.GetTypeName2().ToUpper() == "BOUNDINGBOXPROFILEFEAT")
                    {
                        Entity swEntity = (Entity)f;
                        swEntity.Select4(false, null);
                        model.EditDelete(); // Apaga o recurso WELDMENT
                        break;
                    }
                    f = (Feature)f.GetNextFeature();
                }

                return caixaDelimitadora;
            }
            catch (Exception ex)
            {
                LOG.GravarLog($"{nameof(SLD_PROPRIEDADE).ToUpper()}:{nameof(ObtemCaixaDelimitadora)}",
                    "ERRO - Ao obter a caixa delimitadora. Ative o DEBUG para mais detalhes.", ex);
                return null; // Retorna null em caso de erro
            }
        }
    }
}