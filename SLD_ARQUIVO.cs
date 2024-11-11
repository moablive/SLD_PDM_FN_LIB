// System
using System;
using System.Collections.Generic;
using System.Windows.Forms;


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

        private string getComprimento_CutListFolder()
        {
            try
            {
                string comprimento = "";

                Feature swFeat;

                swFeat = (Feature)swModel.FirstFeature();

                while ((swFeat != null))
                {
                    string a = swFeat.GetTypeName();
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
                throw new Exception(ex.Message);
            }
        }
        private string getLargura_CutListFolder()
        {
            try
            {
                string largura = "";

                Feature swFeat;

                swFeat = (Feature)swModel.FirstFeature();

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
                throw new Exception(ex.Message);
            }
        }
        private string getEspessura_CutListFolder()
        {
            try
            {
                string espessura = "";

                Feature swFeat;

                swFeat = (Feature)swModel.FirstFeature();

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
                throw new Exception(ex.Message);
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

                string[] vCustomPropNames = vCustomPropNames = (string[])CustomPropMgr.GetNames();

                if ((vCustomPropNames != null))
                {
                    for (int ii = 0; ii <= (vCustomPropNames.Length - 1); ii++)
                    {
                        string CustomPropName = null;

                        CustomPropName = (string)vCustomPropNames[ii];

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
                throw new Exception(ex.Message);
            }
        }

        /// <summary>
        /// 1 = CRIA AUTOMATICO
        /// 2 = ATUALIZA
        /// </summary>
        /// <param name="opcao"></param>
        /// <exception cref="Exception"></exception>
        private void AtualizarCutList(int opcao)
        {
            try
            {
                //Metodo utilizado para atualizar a CUTLIST ou definir como atualização automatica, pois as vezes os dados não são carregados pois estão desatualizados
                BodyFolder swBodyFolder;
                SelectionMgr swSelMgr = default(SelectionMgr);
                swSelMgr = (SelectionMgr)swModel.SelectionManager;

                bool boolstatus = swModel.Extension.SelectByID2("Solid Bodies", "BDYFOLDER", 0, 0, 0, false, 0, null, 0);
                if (!boolstatus)
                {
                    boolstatus = swModel.Extension.SelectByID2("Corpos sólidos", "BDYFOLDER", 0, 0, 0, false, 0, null, 0);
                }
                if (boolstatus)
                {
                    Feature swFeat = (Feature)swSelMgr.GetSelectedObject6(1, -1);
                    swBodyFolder = (BodyFolder)swFeat.GetSpecificFeature2();
                    //Escolher uma opção
                    //1) Aqui define se quer fazer automaticamente
                    if (opcao == 1)
                    {
                        swBodyFolder.SetAutomaticCutList(true);
                        swBodyFolder.SetAutomaticUpdate(true);
                    }
                    else
                    {
                        //2) Aqui se quer somente atualizar sem deixar automatico
                        swBodyFolder.UpdateCutList();
                    }

                }
                else
                {
                    MessageBox.Show("ERRO AO OBTER DADOS DA CUTLIST! \nVERIFIQUE O LOG PARA MAIS DETALHES!", "ERRO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    DebugSKA.Log.GravarLog($"{this.GetType().Name.ToUpper() + ":" + nameof(AtualizarCutList)}", " Erro ao obter os valores da Cut List." +
                        "\nÉ necessário atualizar a lista de corte clicando com o direito em Lista de Corte(CutList), e após clicar em 'Atualizar' ou marcar a opção 'Atualizar Automaticamente'.", null);
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }
    }
}
