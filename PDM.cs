//System
using System;
using System.Collections.Generic;
using System.Windows.Forms;

// DLL PDM
using EPDM.Interop.epdm;

namespace PDM
{
    public static class PDM_FN
    {
        private static List<string> _pdmlstDocumentos = new List<string>(); // LST DE BUSCA PARA RETORNO
        private static IEdmVault7 CofrePDM; // COFRE PARA RETORNO DE LOGIN

        ///<summary>
        /// strVar: Valor a ser consultado
        /// strTipo: Se estiver com o valor de "Path" retorna o nome completo do arquivo. Se estiver com o valor "Name" retorna somente o nome do arquivo
        /// var1: Nome da variável do cartão de dados onde a pesquisa será feita
        /// </summary>
        /// <param name="CofrePDM"></param>
        /// <param name="strVar"></param>
        /// <param name="strTipo"></param>
        /// <param name="var1"></param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        public static List<string> FindDocumentsbyVariable(IEdmVault7 CofrePDM, string strVar, string strTipo, string var1)
        {
            try
            {
                // LIMPA LST PARA NOVA BUSCA
                _pdmlstDocumentos.Clear();

                // Inicializa a busca no PDM
                IEdmSearch8 search = (IEdmSearch8)CofrePDM.CreateSearch();

                // Define a variável e o valor para a busca
                string variavel = var1;
                string valor = strVar;

                // Configura a busca para encontrar apenas arquivos
                search.SetToken(EdmSearchToken.Edmstok_FindFiles, true);

                // Configura para buscar apenas a última versão dos arquivos
                search.SetToken(EdmSearchToken.Edmstok_AllVersions, false);

                // Adiciona o critério de pesquisa
                search.AddVariable(var1, strVar);

                // Inicializa o resultado da busca
                IEdmSearchResult5 result = search.GetFirstResult();

                // Verifica se nenhum resultado foi encontrado
                if (result == null)
                {
                    _pdmlstDocumentos.Add("");
                    return _pdmlstDocumentos; // Retorna lista com item vazio se não houver resultados
                }

                // Itera sobre todos os resultados encontrados
                while (result != null)
                {
                    // Verifica se o tipo do objeto é um arquivo
                    if (result.ObjectType == EdmObjectType.EdmObject_File)
                    {
                        // Adiciona o nome ou caminho do arquivo à lista, dependendo do tipo especificado
                        if (strTipo.Equals("Name", StringComparison.OrdinalIgnoreCase))
                        {
                            _pdmlstDocumentos.Add(result.Name);
                        }
                        else if (strTipo.Equals("Path", StringComparison.OrdinalIgnoreCase))
                        {
                            _pdmlstDocumentos.Add(result.Path);
                        }
                    }

                    // Obtém o próximo resultado
                    result = search.GetNextResult();
                }

                return _pdmlstDocumentos; // Retorna a lista de documentos encontrados
            }
            catch (Exception ex)
            {
                DebugSKA.Log.GravarLog($"{typeof(PDM_FN).Name.ToUpper() + ":" + nameof(FindDocumentsbyVariable)}", "ERRO - BUSCAR NO PDM. Ative o DEBUG para mais detalhes.", ex);

                // Método global de proteção contra erros do PDM
                ERROR_RELOAD.RestartExplorer();

                throw new Exception(ex.Message);
            }
        }

        public static string getVar_Cartao_PDM(IEdmVault7 CofrePDM, string nomeVariavel, string arquivo, string configuracao)
        {
            IEdmFile8 file;
            IEdmFolder5 folder;

            file = (IEdmFile8)CofrePDM.GetFileFromPath((string)arquivo, out folder);

            IEdmEnumeratorVariable8 enumVariable = (IEdmEnumeratorVariable8)file.GetEnumeratorVariable("");
            EdmStrLst5 cfgList = default(EdmStrLst5);
            IEdmPos5 pos = default(IEdmPos5);
            string cfgName = null;
            object res = null;

            try
            {
                cfgList = file.GetConfigurations();
                pos = cfgList.GetHeadPosition();

                if (configuracao != "")
                {
                    enumVariable.GetVar(nomeVariavel, configuracao, out res);
                }
                else
                {
                    while (!pos.IsNull)
                    {
                        cfgName = cfgList.GetNext(pos);

                        enumVariable.GetVar(nomeVariavel, cfgName, out res);

                        if (res != null)
                        {
                            break;
                        }
                    }
                }
                enumVariable.CloseFile(true);
                enumVariable.Flush();
            }
            catch (Exception ex)
            {
                DebugSKA.Log.GravarLog($"{typeof(PDM_FN).Name.ToUpper() + ":" + nameof(getVar_Cartao_PDM)}", "ERRO - Ao Obter Valor da Variavel do CARTAO PDM. Ative o DEBUG para mais detalhes.", ex);

                // Método global de proteção contra erros do PDM
                ERROR_RELOAD.RestartExplorer();

                throw new Exception(ex.Message);
            }

            // Valor localizado
            return res?.ToString();
        }

        public static void setVar_Cartao(IEdmFile13 file, string NomVar, Object sVal)
        {
            try
            {
                EdmStrLst5 cfgList = default(EdmStrLst5);
                cfgList = file.GetConfigurations();

                IEdmPos5 pos = default(IEdmPos5);
                pos = cfgList.GetHeadPosition();
                string cfgName = null;

                while (!pos.IsNull)
                {
                    cfgName = cfgList.GetNext(pos);
                    IEdmEnumeratorVariable5 EnumVar = file.GetEnumeratorVariable("");

                    //Preenche a variável
                    EnumVar.SetVar(NomVar, cfgName, ref sVal, false);

                    //Este variável foi declarada apenas para poder utilizar 'CloseFile'. Não sei por quê disso mas no help está assim
                    IEdmEnumeratorVariable8 EnumVarCod8 = (IEdmEnumeratorVariable8)EnumVar;

                    EnumVarCod8.CloseFile(true);
                    EnumVar.Flush();

                }
            }
            catch (Exception ex)
            {
                DebugSKA.Log.GravarLog($"{typeof(PDM_FN).Name.ToUpper() + ":" + nameof(setVar_Cartao_PDM)}", "ERRO - Ao Gravar Valor da Variavel do CARTAO PDM. Ative o DEBUG para mais detalhes.", ex);

                // Método global de proteção contra erros do PDM
                ERROR_RELOAD.RestartExplorer();

                throw new Exception(ex.Message);
            }

        }

        public static void Checkout(IEdmVault7 CofrePDM, string dirFILE)
        {
            try
            {
                IEdmFile13 file;
                IEdmFolder5 folder;
                file = (IEdmFile13)CofrePDM.GetFileFromPath(dirFILE, out folder);

                if (!file.IsLocked)
                {
                    file.LockFile(folder.ID, 0);
                }
            }
            catch (Exception ex)
            {
                DebugSKA.Log.GravarLog($"{typeof(PDM_FN).Name.ToUpper() + ":" + ":" + nameof(CheckoutFile)}", "ERRO - Ao Fazer Checkout do ARQUIVO do PDM. Ative o DEBUG para mais detalhes.", ex);

                // Método global de proteção contra erros do PDM
                ERROR_RELOAD.RestartExplorer();

                throw new Exception(ex.Message);
            }
        }

        public static void CheckIn(IEdmVault7 CofrePDM, string dirFILE)
        {
            try
            {
                IEdmFile13 file;
                IEdmFolder5 folder;

                file = (IEdmFile13)CofrePDM.GetFileFromPath(dirFILE, out folder);

                if (file.IsLocked)
                {
                    file.UnlockFile(0, "Checkin - SKA VIA API", 8);
                }
            }
            catch (Exception ex)
            {
                DebugSKA.Log.GravarLog($"{typeof(PDM_FN).Name.ToUpper() + ":" + nameof(CheckInFile)}", "ERRO - Ao Fazer Cheking do ARQUIVO do PDM. Ative o DEBUG para mais detalhes.", ex);

                // Método global de proteção contra erros do PDM
                ERROR_RELOAD.RestartExplorer();

                throw new Exception(ex.Message);
            }
        }

        public static IEdmVault7 LoginPDM(string NOMECOFRE)
        {
            try
            {
                // Inicializa o CofrePDM se ainda não estiver instanciado
                if (CofrePDM == null)
                    CofrePDM = new EdmVault5();

                // Tenta logar automaticamente no Cofre PDM
                if (!CofrePDM.IsLoggedIn)
                    CofrePDM.LoginAuto(NOMECOFRE, 0);

                // Verifica se o login foi bem-sucedido
                if (!CofrePDM.IsLoggedIn)
                {
                    MessageBox.Show("Erro ao realizar login no Cofre PDM.", "Erro de Login", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                // Registra o erro no log e lança uma exceção com a mensagem de erro
                DebugSKA.Log.GravarLog($"{typeof(PDM_FN).Name.ToUpper() + ":" + nameof(LoginPDM)}", "ERRO - Ao Logar no PDM. Ative o DEBUG para mais detalhes.", ex);

                throw new Exception(ex.Message);
            }

            return CofrePDM; // Retorna o objeto CofrePDM
        }
    }
}
