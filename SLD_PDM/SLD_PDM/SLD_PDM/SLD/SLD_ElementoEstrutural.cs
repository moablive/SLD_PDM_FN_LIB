// System
using System;
using System.Windows.Forms;

// Project
using SLD_PDM.SLD.MODEL;

// SolidWorks
using SolidWorks.Interop.sldworks;

namespace SLD_PDM.SLD
{
    public class SLD_ElementoEstrutural
    {
        public SldWorks swApp;
        public IModelDoc2 swModel;

        public SLD_ElementoEstrutural(SldWorks sldWorksApp)
        {
            swApp = sldWorksApp;
        }

        // Verifica se o modelo é elemento estrutural
        public bool iselementoEstrutural(IModelDoc2 model)
        {
            try
            {
                bool Weldment = false;

                // Verifica se o documento ativo é uma peça
                if (model is PartDoc)
                {
                    PartDoc partDoc = (PartDoc)model;

                    // Usando o método IsWeldment para verificar
                    Weldment = partDoc.IsWeldment();
                }

                return Weldment;
            }
            catch (Exception ex)
            {
                LOG.GravarLog($"{nameof(SLD_ElementoEstrutural).ToUpper()}:{nameof(iselementoEstrutural)}",
                    "ERRO - Ao verificar se o modelo é um elemento estrutural. Ative o DEBUG para mais detalhes.", ex);
                throw; // Relança a exceção para que seja tratada em níveis superiores
            }
        }

        public objLISTACORTE ElementoEstrutural(IModelDoc2 model)
        {
            objLISTACORTE listaCorte = null;

            try
            {
                Feature swFeat = (Feature)model.FirstFeature();
                CustomPropertyManager customPropMgr = null;

                while (swFeat != null)
                {
                    // Verifica se o tipo é "CutListFolder"
                    if (swFeat.GetTypeName() == "CutListFolder" && !swFeat.ExcludeFromCutList && !swFeat.IsSuppressed())
                    {
                        customPropMgr = (CustomPropertyManager)swFeat.CustomPropertyManager;

                        // Cria o objeto objLISTACORTE com as propriedades obtidas
                        listaCorte = new objLISTACORTE
                        {
                            comprimento = getValorDaPropriedade(customPropMgr, "COMPRIMENTO"),
                            description = getValorDaPropriedade(customPropMgr, "DESCRIPTION"),
                            dimensoes = getValorDaPropriedade(customPropMgr, "DIMENSOES"),
                            peso = getValorDaPropriedade(customPropMgr, "PESO"),
                            quantity = getValorDaPropriedade(customPropMgr, "QUANTITY"),
                            totallength = getValorDaPropriedade(customPropMgr, "TOTAL LENGTH")
                        };

                        // Interrompe o loop após encontrar o primeiro objeto correspondente
                        break;
                    }

                    // Passa para a próxima feature
                    swFeat = (Feature)swFeat.GetNextFeature();
                }
            }
            catch (Exception ex)
            {
                LOG.GravarLog($"{nameof(SLD_ElementoEstrutural).ToUpper()}:{nameof(ElementoEstrutural)}",
                    "ERRO - Ao instanciar o elemento estrutural. Ative o DEBUG para mais detalhes.", ex);
            }

            return listaCorte;
        }

        private string getValorDaPropriedade(CustomPropertyManager CustomPropMgr, string propriedade)
        {
            string CustomPropResolvedVal = string.Empty;

            try
            {
                string CustomPropVal = string.Empty;
                CustomPropMgr.Get2(propriedade, out CustomPropVal, out CustomPropResolvedVal);
            }
            catch (Exception ex)
            {
                LOG.GravarLog($"{nameof(SLD_ElementoEstrutural).ToUpper()}:{nameof(getValorDaPropriedade)}",
                    $"ERRO - Ao obter o valor da propriedade '{propriedade}'. Ative o DEBUG para mais detalhes.", ex);
            }

            return CustomPropResolvedVal;
        }

        /// <summary>
        /// Atualiza ou define a atualização automática da CutList.
        /// </summary>
        /// <param name="opcao">1 para atualização automática, 2 para atualização manual.</param>
        public void AtualizarCutList(int opcao)
        {
            try
            {
                // Método utilizado para atualizar a CUTLIST ou definir como atualização automática
                BodyFolder swBodyFolder;
                SelectionMgr swSelMgr = (SelectionMgr)swModel.SelectionManager;

                bool boolstatus = swModel.Extension.SelectByID2("Solid Bodies", "BDYFOLDER", 0, 0, 0, false, 0, null, 0);
                if (!boolstatus)
                {
                    boolstatus = swModel.Extension.SelectByID2("Corpos sólidos", "BDYFOLDER", 0, 0, 0, false, 0, null, 0);
                }

                if (boolstatus)
                {
                    Feature swFeat = (Feature)swSelMgr.GetSelectedObject6(1, -1);
                    swBodyFolder = (BodyFolder)swFeat.GetSpecificFeature2();

                    // Escolher uma opção
                    if (opcao == 1)
                    {
                        // Atualização automática
                        swBodyFolder.SetAutomaticCutList(true);
                        swBodyFolder.SetAutomaticUpdate(true);
                    }
                    else
                    {
                        // Atualização manual
                        swBodyFolder.UpdateCutList();
                    }
                }
                else
                {
                    LOG.GravarLog($"{nameof(SLD_ElementoEstrutural).ToUpper()}:{nameof(AtualizarCutList)}",
                        "ERRO - Não foi possível selecionar os corpos sólidos para atualizar a CutList.");
                    MessageBox.Show("ERRO AO OBTER DADOS DA CUTLIST! \nVERIFIQUE O LOG PARA MAIS DETALHES!",
                        "ERRO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                LOG.GravarLog($"{nameof(SLD_ElementoEstrutural).ToUpper()}:{nameof(AtualizarCutList)}",
                    "ERRO - Ao atualizar a CutList. Ative o DEBUG para mais detalhes.", ex);
                throw; // Relança a exceção para tratamento superior
            }
        }
    }
}
