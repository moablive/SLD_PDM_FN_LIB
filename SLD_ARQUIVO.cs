//System;
using System;

// SolidWorks DLLs
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace SLD
{
    public class arquivoSLD
    {
        private SldWorks swApp;
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

                // Abre o documento no SolidWorks
                swModel = swApp.OpenDoc6(caminhoArquivo, tipoDocumento, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref errors, ref warnings);

                if (swModel == null)
                {
                    throw new Exception("Não foi possível instanciar o arquivo no SolidWorks.");
                }

                return swModel;
            }
            catch (Exception ex)
            {
                Log.GravarLog($"{typeof(arquivoSLD).Name.ToUpper() + ":" + nameof(AbrirDocumento)}",
                    "ERRO - Ao Abrir Documento. Ative o DEBUG para mais detalhes.",
                    ex);
            }
        }

        public string GetActiveConfiguration(IModelDoc2 _swModel)
        {
            try
            {
                return _swModel.ConfigurationManager.ActiveConfiguration.Name;
            }
            catch (Exception ex)
            {
                Log.GravarLog($"{typeof(arquivoSLD).Name.ToUpper() + ":" + nameof(GetActiveConfiguration)}",
                    "ERRO - Ao Obter a Configuracao Ativa. Ative o DEBUG para mais detalhes.",
                    ex);
            }
        }

        public List<string> GetAllConfigurations(IModelDoc2 swModel)
        {
            var configNamesList = new List<string>();

            try
            {
                var configNames = (string[])swModel.GetConfigurationNames();

                foreach (var configName in configNames)
                {
                    var swConfig = (Configuration)swModel.GetConfigurationByName(configName);

                    // Adiciona o nome da configuração e outros detalhes relevantes
                    configNamesList.Add($"Name: {swConfig.Name}, " +
                                        $"Alternate Name in BOM: {swConfig.UseAlternateNameInBOM}, " +
                                        $"Alternate Name: {swConfig.AlternateName}, " +
                                        $"Comment: {swConfig.Comment}");
                }
            }
            catch (Exception ex)
            {
                Log.GravarLog($"{typeof(arquivoSLD).Name.ToUpper() + ":" + nameof(GetAllConfigurations)}",
                    "ERRO - Ao Obter Todas as Configuracoes. Ative o DEBUG para mais detalhes.",
                    ex);
            }

            return configNamesList;
        }

        public void setPropriedade(string Valor, string Config, string pName)
        {
            try
            {
                ModelDocExtension swModelDocExt = default(ModelDocExtension);
                CustomPropertyManager swCustProp = default(CustomPropertyManager);

                swModelDocExt = swModel.Extension;

                // Get the custom property data
                swCustProp = swModelDocExt.get_CustomPropertyManager(Config);
                try
                {
                    swCustProp.Delete(pName);
                }
                catch (Exception ex)
                {
                    Log.GravarLog($"{typeof(arquivoSLD).Name.ToUpper() + ":" + nameof(setPropriedade)}",
                        "ERRO - Ao setPropriedade. Ative o DEBUG para mais detalhes.",
                        ex);
                }

                swCustProp.Add3(pName, 30, Valor, 1);
            }
            catch (Exception ex)
            {
                Log.GravarLog($"{typeof(arquivoSLD).Name.ToUpper() + ":" + nameof(setPropriedade)}",
                    "ERRO - Ao setPropriedade. Ative o DEBUG para mais detalhes.",
                    ex);
            }
        }
    }
}