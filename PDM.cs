//System
using System;
using System.Collections.Generic;
using System.Windows.Forms;

// DLL PDM
using EPDM.Interop.epdm;

namespace SLD_PDM
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
                Log.GravarLog($"{typeof(PDM_FN).Name.ToUpper() + ":" + nameof(Login)}", "ERRO - Ao Logar no PDM. Ative o DEBUG para mais detalhes.", ex);
                throw new Exception(ex.Message);
            }

            return CofrePDM;
        }

        // PDM => FILE
        public static void Check_out(IEdmVault7 vault, string filePath)
        {
            try
            {
                IEdmFolder5 folder;
                IEdmFile13 file = (IEdmFile13)vault.GetFileFromPath(filePath, out folder);

                if (file != null && !file.IsLocked)
                {
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
                if (file.IsLocked)
                {
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
                IEdmFile5 file;
                IEdmFolder5 folder;
                file = (IEdmFile5)cofre.GetFileFromPath(caminhoArquivo, out folder);

                if (file != null && folder != null)
                {
                    int fileID = file.ID;
                    folder.DeleteFile(0, fileID, true);
                }
                else
                {
                    MessageBox.Show($"Arquivo '{caminhoArquivo}' não encontrado no PDM.", "Arquivo não encontrado", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao excluir o arquivo '{caminhoArquivo}' do PDM: {ex.Message}", "Erro de Exclusão", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Log.GravarLog($"{typeof(PDM_FN).Name.ToUpper() + ":" + nameof(ExcluirArquivoPDM)}", "ERRO - Ao Excluir Arquivo do PDM. Ative o DEBUG para mais detalhes.", ex);
            }
        }

        // PDM => CARTAO
        public static string getVar_Cartao(IEdmFile13 file, string nomeVariavel, string configuracao)
        {
            object res = null;

            try
            {
                IEdmEnumeratorVariable8 enumVariable = (IEdmEnumeratorVariable8)file.GetEnumeratorVariable("");

                if (!string.IsNullOrEmpty(configuracao))
                {
                    enumVariable.GetVar(nomeVariavel, configuracao, out res);
                }
                else
                {
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

                if (!string.IsNullOrEmpty(configuracao))
                {
                    enumVar.SetVar(nomeVariavel, configuracao, ref valor, false);
                }
                else
                {
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

        public static List<string> FindDocumentsByVariable(IEdmVault7 vault, string variableValue, string returnType, string variableName)
        {
            var documentPaths = new List<string>();

            try
            {
                IEdmSearch8 search = (IEdmSearch8)vault.CreateSearch();
                search.SetToken(EdmSearchToken.Edmstok_FindFiles, true);
                search.SetToken(EdmSearchToken.Edmstok_AllVersions, false);
                search.AddVariable(variableName, variableValue);

                IEdmSearchResult5 result = search.GetFirstResult();

                while (result != null)
                {
                    if (result.ObjectType == EdmObjectType.EdmObject_File)
                    {
                        if (returnType.Equals("Name", StringComparison.OrdinalIgnoreCase))
                        {
                            documentPaths.Add(result.Name);
                        }
                        else if (returnType.Equals("Path", StringComparison.OrdinalIgnoreCase))
                        {
                            documentPaths.Add(result.Path);
                        }
                    }

                    result = search.GetNextResult();
                }

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
