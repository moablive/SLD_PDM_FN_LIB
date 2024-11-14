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
        private static string NOMECOFRE = "VAULT_NAME"; // NOME COFRE PDM
        private static IEdmVault7 CofrePDM; // INSTANCIA DO COFRE 
         private static IEdmFolder5 folder; // PDM OUT FOLDER

        #region EXEMPLOS PDM
        // EXEMPLO DE FILE 13: (IEdmFile13)CofrePDM.GetFileFromPath(dirFILE, out folder)
        #endregion

        // PDM => LOGIN
        public static IEdmVault7 Login()
        {
            try
            {
                if (CofrePDM == null)
                    CofrePDM = new EdmVault5();

                if (!CofrePDM.IsLoggedIn)
                    CofrePDM.LoginAuto(NOMECOFRE, 0);

                if (!CofrePDM.IsLoggedIn)
                {
                    MessageBox.Show("Erro ao realizar login no Cofre PDM.", "Erro de Login", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                Log.GravarLog($"{typeof(PDM_FN).Name.ToUpper() + ":" + nameof(Login)}","ERRO - Ao Logar no PDM. Ative o DEBUG para mais detalhes.",ex);
                throw new Exception(ex.Message);
            }

            return CofrePDM;
        }

        // PDM => FILE
        public static void Check_out(IEdmVault7 vault, string filePath)
        {
            try
            {
                // Obtém o arquivo e a pasta associada a partir do caminho do arquivo
                IEdmFolder5 folder;
                IEdmFile13 file = (IEdmFile13)vault.GetFileFromPath(filePath, out folder);

                if (file != null && !file.IsLocked)
                {
                    // Realiza o checkout do arquivo usando o ID da pasta
                    file.LockFile(folder.ID, 0);
                }
                else if (file == null)
                {
                    throw new Exception("Arquivo não encontrado no PDM.");
                }
            }
            catch (Exception ex)
            {
                Log.GravarLog($"{typeof(PDM_FN).Name.ToUpper() + ":" + nameof(Check_out)}",
                            "ERRO - Ao Fazer Checkout do ARQUIVO do PDM. Ative o DEBUG para mais detalhes.",
                            ex);

                ERROR_RELOAD.RestartExplorer();

                throw new Exception(ex.Message);
            }
        }

        public static void Check_In(IEdmFile13 file)
        {
            try
            {
                // Verifica se o arquivo está bloqueado antes de desbloquear
                if (file.IsLocked)
                {
                    // Desbloqueia o arquivo e finaliza o check-in
                    file.UnlockFile(0, "Checkin - VIA API", 8);
                }
            }
            catch (Exception ex)
            {
                Log.GravarLog($"{typeof(PDM_FN).Name.ToUpper() + ":" + nameof(Check_In)}",
                            "ERRO - Ao Fazer Checkin do ARQUIVO do PDM. Ative o DEBUG para mais detalhes.",
                            ex);

                ERROR_RELOAD.RestartExplorer();

                throw new Exception(ex.Message);
            }
        }

        public static void ExcluirArquivoPDM(IEdmVault7 cofre, string caminhoArquivo)
        {
            try
            {
                // Obtém o arquivo e o diretório associado
                IEdmFile5 file;
                IEdmFolder5 folder;
                file = (IEdmFile5)cofre.GetFileFromPath(caminhoArquivo, out folder);

                if (file != null && folder != null)
                {
                    int fileID = file.ID;
                    int folderID = folder.ID;

                    // Tenta excluir o arquivo
                    folder.DeleteFile(0, fileID, true); // `true` para remover a cópia local
                }
                else
                {
                    MessageBox.Show($"Arquivo '{caminhoArquivo}' não encontrado no PDM.", "Arquivo não encontrado", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao excluir o arquivo '{caminhoArquivo}' do PDM: {ex.Message}", "Erro de Exclusão", MessageBoxButtons.OK, MessageBoxIcon.Error);
                DebugSKA.Log.GravarLog($"{typeof(PDM).Name.ToUpper() + ":" + nameof(ExcluirArquivoPDM)}", "ERRO - Ao Excluir Arquivo do PDM. Ative o DEBUG para mais detalhes.", ex);
            }
        }

        // PDM => CARTAO
        public static string getVar_Cartao(IEdmFile13 file, string nomeVariavel, string configuracao)
        {
            object res = null;

            try
            {
                IEdmEnumeratorVariable8 enumVariable = (IEdmEnumeratorVariable8)file.GetEnumeratorVariable("");

                // Verifica se a configuração específica foi passada e tenta obter o valor
                if (!string.IsNullOrEmpty(configuracao))
                {
                    enumVariable.GetVar(nomeVariavel, configuracao, out res);
                }
                else
                {
                    // Obtém a lista de configurações e busca o valor para a primeira configuração encontrada
                    EdmStrLst5 cfgList = file.GetConfigurations();
                    IEdmPos5 pos = cfgList.GetHeadPosition();

                    while (!pos.IsNull)
                    {
                        string cfgName = cfgList.GetNext(pos);

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
                Log.GravarLog($"{typeof(PDM_FN).Name.ToUpper() + ":" + nameof(getVar_Cartao)}",
                            "ERRO - Ao Obter Valor da Variavel do CARTAO PDM. Ative o DEBUG para mais detalhes.",
                            ex);

                ERROR_RELOAD.RestartExplorer();

                throw new Exception(ex.Message);
            }

            return res?.ToString();
        }

        public static void setVar_Cartao(IEdmFile13 file, string nomeVariavel, object valor, string configuracao)
        {
            try
            {
                IEdmEnumeratorVariable5 enumVar = file.GetEnumeratorVariable("");

                // Se a configuração foi especificada, aplica o valor diretamente nela
                if (!string.IsNullOrEmpty(configuracao))
                {
                    enumVar.SetVar(nomeVariavel, configuracao, ref valor, false);
                }
                else
                {
                    // Caso contrário, aplica o valor em todas as configurações
                    EdmStrLst5 cfgList = file.GetConfigurations();
                    IEdmPos5 pos = cfgList.GetHeadPosition();

                    while (!pos.IsNull)
                    {
                        string cfgName = cfgList.GetNext(pos);
                        enumVar.SetVar(nomeVariavel, cfgName, ref valor, false);
                    }
                }

                IEdmEnumeratorVariable8 enumVarCod8 = (IEdmEnumeratorVariable8)enumVar;
                enumVarCod8.CloseFile(true);
                enumVar.Flush();
            }
            catch (Exception ex)
            {
                Log.GravarLog($"{typeof(PDM_FN).Name.ToUpper() + ":" + nameof(setVar_Cartao)}",
                            "ERRO - Ao Gravar Valor da Variavel do CARTAO PDM. Ative o DEBUG para mais detalhes.",
                            ex);

                ERROR_RELOAD.RestartExplorer();

                throw new Exception(ex.Message);
            }
        }

        /// <summary>
        /// Retorna uma lista de caminhos ou nomes de arquivos que contêm uma variável específica com o valor especificado no PDM.
        /// </summary>
        /// <param name="vault">Instância do cofre PDM (IEdmVault7)</param>
        /// <param name="variableValue">Valor que a variável deve possuir</param>
        /// <param name="returnType">Tipo de retorno: "Path" para o caminho completo ou "Name" para apenas o nome do arquivo</param>
        /// <param name="variableName">Nome da variável a ser pesquisada</param>
        /// <returns>Lista de strings contendo o caminho ou nome dos arquivos encontrados</returns>
        /// <exception cref="Exception"></exception>
        public static List<string> FindDocumentsByVariable(IEdmVault7 vault, string variableValue, string returnType, string variableName)
        {
            var documentPaths = new List<string>();

            try
            {
                // Inicializa a busca no PDM
                IEdmSearch8 search = (IEdmSearch8)vault.CreateSearch();

                // Configura a busca para encontrar apenas arquivos
                search.SetToken(EdmSearchToken.Edmstok_FindFiles, true);

                // Configura para buscar apenas a última versão dos arquivos
                search.SetToken(EdmSearchToken.Edmstok_AllVersions, false);

                // Adiciona o critério de pesquisa para a variável e o valor
                search.AddVariable(variableName, variableValue);

                // Inicializa o resultado da busca
                IEdmSearchResult5 result = search.GetFirstResult();

                // Itera sobre todos os resultados encontrados
                while (result != null)
                {
                    // Verifica se o tipo do objeto é um arquivo
                    if (result.ObjectType == EdmObjectType.EdmObject_File)
                    {
                        // Adiciona o nome ou caminho do arquivo à lista, dependendo do tipo especificado
                        if (returnType.Equals("Name", StringComparison.OrdinalIgnoreCase))
                        {
                            documentPaths.Add(result.Name);
                        }
                        else if (returnType.Equals("Path", StringComparison.OrdinalIgnoreCase))
                        {
                            documentPaths.Add(result.Path);
                        }
                    }

                    // Obtém o próximo resultado
                    result = search.GetNextResult();
                }

                // Retorna a lista de documentos encontrados (ou vazia se nenhum for encontrado)
                return documentPaths;
            }
            catch (Exception ex)
            {
                Log.GravarLog($"{typeof(PDM_FN).Name.ToUpper()}:{nameof(FindDocumentsByVariable)}",
                            "ERRO - Ao buscar arquivos no PDM. Ative o DEBUG para mais detalhes.",
                            ex);

                ERROR_RELOAD.RestartExplorer();

                throw new Exception("Erro ao buscar arquivos com a variável e valor especificados no PDM.", ex);
            }
        }
    }
}
