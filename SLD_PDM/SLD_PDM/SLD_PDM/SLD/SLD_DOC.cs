// System
using System;
using System.Collections.Generic;

// DLL SolidWorks
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace SLD_PDM.SLD
{
    public class SLD_DOC
    {
        #region SLD_DOC
        public SldWorks swApp;
        public IModelDoc2 swModel;

        // Construtor de SLD_DOC
        public SLD_DOC(SldWorks sldWorksApp)
        {
            swApp = sldWorksApp;
        }
        #endregion

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
                LOG.GravarLog($"{nameof(SLD_DOC).ToUpper()}:{nameof(AbrirDocumento)}",
                    "ERRO - Ao Abrir Documento. Ative o DEBUG para mais detalhes.", ex);
                return null;
            }
        }
        public void FecharDocumento(string caminhoArquivo)
        {
            try
            {
                swApp.CloseDoc(caminhoArquivo);
            }
            catch (Exception ex)
            {
                LOG.GravarLog($"{nameof(SLD_DOC).ToUpper()}:{nameof(FecharDocumento)}",
                    "ERRO - Ao Fechar Documento. Ative o DEBUG para mais detalhes.", ex);
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
                LOG.GravarLog($"{nameof(SLD_DOC).ToUpper()}:{nameof(GetActiveConfiguration)}",
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
                LOG.GravarLog($"{nameof(SLD_DOC).ToUpper()}:{nameof(GetAllConfigurations)}",
                    "ERRO - Ao Obter Todas as Configurações. Ative o DEBUG para mais detalhes.", ex);
            }

            return configNamesList;
        }
    }
}