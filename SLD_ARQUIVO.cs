// System
using System;
using System.Collections.Generic;

// SKA
using DebugSKA;

// DLL SolidWorks
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace rodovale.SLD_PDM
{
    public class arquivoSLD
    {
        private readonly SldWorks swApp;
        private IModelDoc2 swModel;

        public arquivoSLD(SldWorks sldWorksApp)
        {
            swApp = sldWorksApp;
        }

        public IModelDoc2 AbrirDocumento(string caminhoArquivo, int tipoDocumento, bool visivel)
        {
            try
            {
                int errors = 0, warnings = 0;

                swApp.Visible = visivel;
                swModel = swApp.OpenDoc6(caminhoArquivo, tipoDocumento,
                    (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref errors, ref warnings);

                if (swModel == null)
                {
                    throw new Exception("Não foi possível instanciar o arquivo no SolidWorks.");
                }

                return swModel;
            }
            catch (Exception ex)
            {
                Log.GravarLog($"{nameof(arquivoSLD).ToUpper()}:{nameof(AbrirDocumento)}",
                    "ERRO - Ao Abrir Documento. Ative o DEBUG para mais detalhes.", ex);
                return null;
            }
        }

        public string GetActiveConfiguration(IModelDoc2 model)
        {
            try
            {
                return model?.ConfigurationManager.ActiveConfiguration.Name
                       ?? throw new Exception("Configuração ativa não encontrada.");
            }
            catch (Exception ex)
            {
                Log.GravarLog($"{nameof(arquivoSLD).ToUpper()}:{nameof(GetActiveConfiguration)}",
                    "ERRO - Ao Obter a Configuração Ativa. Ative o DEBUG para mais detalhes.", ex);
                return string.Empty;
            }
        }

        public List<string> GetAllConfigurations(IModelDoc2 model)
        {
            var configNamesList = new List<string>();

            try
            {
                var configNames = model?.GetConfigurationNames() as string[];
                if (configNames == null) throw new Exception("Nenhuma configuração encontrada.");

                foreach (var configName in configNames)
                {
                    var config = model.GetConfigurationByName(configName) as Configuration;
                    if (config != null)
                    {
                        configNamesList.Add($"Name: {config.Name}, " +
                                            $"Alternate Name in BOM: {config.UseAlternateNameInBOM}, " +
                                            $"Alternate Name: {config.AlternateName}, " +
                                            $"Comment: {config.Comment}");
                    }
                }
            }
            catch (Exception ex)
            {
                Log.GravarLog($"{nameof(arquivoSLD).ToUpper()}:{nameof(GetAllConfigurations)}",
                    "ERRO - Ao Obter Todas as Configurações. Ative o DEBUG para mais detalhes.", ex);
            }

            return configNamesList;
        }

        public void SetPropriedade(string valor, string config, string propName)
        {
            try
            {
                var swModelDocExt = swModel?.Extension;
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
                    Log.GravarLog($"{nameof(arquivoSLD).ToUpper()}:{nameof(SetPropriedade)}",
                        "ERRO - Ao deletar propriedade existente. Ative o DEBUG para mais detalhes.", ex);
                }

                swCustProp.Add3(propName, (int)swCustomInfoType_e.swCustomInfoText, valor, (int)swCustomPropertyAddOption_e.swCustomPropertyReplaceValue);
            }
            catch (Exception ex)
            {
                Log.GravarLog($"{nameof(arquivoSLD).ToUpper()}:{nameof(SetPropriedade)}",
                    "ERRO - Ao configurar propriedade. Ative o DEBUG para mais detalhes.", ex);
            }
        }


        /// <summary>
        /// Verifica se o modelo é chapa metálica
        /// </summary>
        /// <returns></returns>
        public bool isSheetMetal()
        {
            bool sheetMetal = false;
            try
            {
                ModelDocExtension mdExtension = swModel.Extension;
                Feature f = (Feature)swModel.FirstFeature();

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
                Log.GravarLog($"{nameof(arquivoSLD).ToUpper()}:{nameof(isSheetMetal)}",
                    "ERRO - Ao Verificar se item is SHEETMETAL. Ative o DEBUG para mais detalhes.", ex);
            }
            return sheetMetal;
        }
    }
}
