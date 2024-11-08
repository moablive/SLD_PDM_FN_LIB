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
        IEdmFolder5 folder; // PDM OUT FOLDER

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
        public static void Check_out(IEdmFile13 file)
        {
            try
            {
                // Obtém a pasta associada ao arquivo para realizar o bloqueio
                IEdmFolder5 folder = file.GetFolder();

                // Verifica se o arquivo não está bloqueado
                if (!file.IsLocked)
                {
                    // Realiza o bloqueio do arquivo na pasta correspondente
                    file.LockFile(folder.ID, 0);
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

        // RETORNA CAMINHO DO ARQUIVO QUE TEM ESTA VARIAVEL COM ESTE VALOR
        public static List<string> FindDocumentsbyVariable(IEdmVault7 vault, string variableValue, string returnType, string variableName)
        {
            var documentList = new List<string>();

            try
            {
                // Inicializa a busca no PDM
                IEdmSearch8 search = (IEdmSearch8)vault.CreateSearch();

                // Configura a busca para encontrar apenas arquivos
                search.SetToken(EdmSearchToken.Edmstok_FindFiles, true);

                // Configura para buscar apenas a última versão dos arquivos
                search.SetToken(EdmSearchToken.Edmstok_AllVersions, false);

                // Adiciona o critério de pesquisa com a variável e valor
                search.AddVariable(variableName, variableValue);

                // Inicializa o resultado da busca
                IEdmSearchResult5 result = search.GetFirstResult();

                // Verifica se nenhum resultado foi encontrado
                if (result == null)
                {
                    documentList.Add(""); // Retorna lista com item vazio se não houver resultados
                    return documentList;
                }

                // Itera sobre todos os resultados encontrados
                while (result != null)
                {
                    IEdmFile13 file = (IEdmFile13)result;  // Faz o cast direto para IEdmFile13
                    string filePath = returnType.Equals("Path", StringComparison.OrdinalIgnoreCase)
                        ? file.GetLocalPath(file.GetFolder().ID)
                        : file.Name;

                    documentList.Add(filePath); // Adiciona o caminho ou nome à lista

                    // Obtém o próximo resultado
                    result = search.GetNextResult();
                }

                return documentList; // Retorna a lista de documentos encontrados
            }
            catch (Exception ex)
            {
                Log.GravarLog($"{typeof(PDM_FN).Name.ToUpper() + ":" + nameof(FindDocumentsbyVariable)}",
                            "ERRO - BUSCAR NO PDM. Ative o DEBUG para mais detalhes.",
                            ex);

                ERROR_RELOAD.RestartExplorer();

                throw new Exception(ex.Message);
            }
        }
    }
}