// System
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;

// DLL PDM
using EPDM.Interop.epdm;

namespace SLD_PDM.PDM
{
    public static class PDM
    {
        private static IEdmVault7 CofrePDM; // COFRE PARA RETORNO DE LOGIN
        private static string NOMECOFRE = "XXXXX"; // NOME PARA LOGIN

        public static IEdmVault7 LoginPDM()
        {
            try
            {
                LOG.GravarLog(nameof(LoginPDM), "Realizando login no PDM.");
                if (CofrePDM == null)
                    CofrePDM = new EdmVault5();

                if (!CofrePDM.IsLoggedIn)
                    CofrePDM.LoginAuto(NOMECOFRE, 0);

                if (!CofrePDM.IsLoggedIn)
                {
                    throw new Exception("Erro ao realizar login no PDM.");
                }

                return CofrePDM;
            }
            catch (Exception ex)
            {
                ERROR_RELOAD.RestartExplorer(ex);
                LOG.GravarLog(nameof(LoginPDM), "Erro ao realizar login no PDM.", ex);
                throw;
            }
        }

        #region CheckoutFile - CheckInFile
        public static void CheckoutFile(IEdmVault7 CofrePDM, string dirFILE)
        {
            try
            {
                LOG.GravarLog(nameof(CheckoutFile), $"Realizando checkout do arquivo '{dirFILE}'.");
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
                ERROR_RELOAD.RestartExplorer(ex);
                LOG.GravarLog(nameof(CheckoutFile), "Erro ao realizar checkout.", ex);
                throw;
            }
        }

        public static void CheckInFile(IEdmVault7 CofrePDM, string dirFILE)
        {
            try
            {
                LOG.GravarLog(nameof(CheckInFile), $"Realizando check-in do arquivo '{dirFILE}'.");
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
                ERROR_RELOAD.RestartExplorer(ex);
                LOG.GravarLog(nameof(CheckInFile), "Erro ao realizar check-in.", ex);
                throw;
            }
        }
        #endregion

        #region CARTAO
        public static string GetVar(IEdmVault7 CofrePDM, string nomeVariavel, string arquivo, string configuracao)
        {
            try
            {
                LOG.GravarLog(nameof(GetVar), $"Obtendo variável '{nomeVariavel}' para o arquivo '{arquivo}' com a configuração '{configuracao}'.");
                IEdmFile8 file;
                IEdmFolder5 folder;

                file = (IEdmFile8)CofrePDM.GetFileFromPath((string)arquivo, out folder);
                IEdmEnumeratorVariable8 enumVariable = (IEdmEnumeratorVariable8)file.GetEnumeratorVariable("");
                EdmStrLst5 cfgList = file.GetConfigurations();
                string cfgName = null;
                object res = null;

                if (!string.IsNullOrEmpty(configuracao))
                {
                    enumVariable.GetVar(nomeVariavel, configuracao, out res);
                }
                else
                {
                    IEdmPos5 pos = cfgList.GetHeadPosition();
                    while (!pos.IsNull)
                    {
                        cfgName = cfgList.GetNext(pos);
                        enumVariable.GetVar(nomeVariavel, cfgName, out res);
                        if (res != null) break;
                    }
                }

                enumVariable.CloseFile(true);
                enumVariable.Flush();

                return res?.ToString();
            }
            catch (Exception ex)
            {
                ERROR_RELOAD.RestartExplorer(ex);
                LOG.GravarLog(nameof(GetVar), "Erro ao obter variável.", ex);
                throw;
            }
        }

        public static void SetVar(IEdmFile13 file, string NomVar, object sVal, string configuracao)
        {
            try
            {
                LOG.GravarLog(nameof(SetVar), $"Definindo variável '{NomVar}' para o arquivo com configuração '{configuracao}'.");
                IEdmEnumeratorVariable5 enumVar = file.GetEnumeratorVariable("");
                enumVar.SetVar(NomVar, configuracao, ref sVal, false);
                IEdmEnumeratorVariable8 enumVarCod8 = (IEdmEnumeratorVariable8)enumVar;
                enumVarCod8.CloseFile(true);
                enumVar.Flush();
            }
            catch (Exception ex)
            {
                ERROR_RELOAD.RestartExplorer(ex);
                LOG.GravarLog(nameof(SetVar), "Erro ao definir variável.", ex);
                throw;
            }
        }
        #endregion

        #region copiarArquivo
        public static bool copiarArquivo(string nomePastaDestino, string nomeArquivo, string novonome)
        {
            try
            {
                LOG.GravarLog(nameof(copiarArquivo), $"Copiando arquivo '{nomeArquivo}' para '{nomePastaDestino}' como '{novonome}'.");
                IEdmFolder5 folderOrigem = null;
                IEdmFile8 file = (IEdmFile8)CofrePDM.GetFileFromPath(nomeArquivo, out folderOrigem);

                if (file != null)
                {
                    ObtemArquivoLocal((IEdmFile7)file, (IEdmFolder6)folderOrigem);
                    IEdmFolder8 folderDestino = (IEdmFolder8)CofrePDM.GetFolderFromPath(nomePastaDestino);

                    if (folderDestino != null)
                    {
                        folderDestino.CopyFile2(file.ID, folderOrigem.ID, 0, out _, novonome);
                        return true;
                    }
                    else
                    {
                        string filePath = retornaBarraFinal(nomePastaDestino) + novonome;
                        System.IO.File.Copy(nomeArquivo, filePath);
                        alteraReadOnlyfileIO(nomeArquivo, false);
                        return true;
                    }
                }
                else
                {
                    LOG.GravarLog(nameof(copiarArquivo), $"Arquivo '{nomeArquivo}' não encontrado no cofre.");
                    return false;
                }
            }
            catch (Exception ex)
            {
                ERROR_RELOAD.RestartExplorer(ex);
                LOG.GravarLog(nameof(copiarArquivo), "Erro ao copiar arquivo.", ex);
                throw;
            }
        }

        public static bool ObtemArquivoLocal(IEdmFile7 file, IEdmFolder6 folder)
        {
            try
            {
                LOG.GravarLog(nameof(ObtemArquivoLocal), "Obtendo arquivo local para o arquivo no cofre.");

                // Declara as variáveis necessárias para o método
                object poVer = null;
                object poPath = folder.ID;
                int versionNumber = 0;
                string resolvedPath = "";

                // Chama o método GetFileCopy com os parâmetros apropriados
                file.GetFileCopy(versionNumber, ref poVer, ref poPath, 0, resolvedPath);

                LOG.GravarLog(nameof(ObtemArquivoLocal), "Arquivo obtido com sucesso.");
                return true;
            }
            catch (Exception ex)
            {
                ERROR_RELOAD.RestartExplorer(ex);
                LOG.GravarLog(nameof(ObtemArquivoLocal), "Erro ao obter arquivo local.", ex);
                return false;
            }
        }

        public static string retornaBarraFinal(string strCaminho)
        {
            try
            {
                LOG.GravarLog(nameof(retornaBarraFinal), $"Ajustando barra final para o caminho '{strCaminho}'.");
                if (!strCaminho.EndsWith(@"\"))
                {
                    strCaminho += @"\";
                }
                return strCaminho;
            }
            catch (Exception ex)
            {
                ERROR_RELOAD.RestartExplorer(ex);
                LOG.GravarLog(nameof(retornaBarraFinal), "Erro ao ajustar barra final.", ex);
                throw;
            }
        }

        public static void alteraReadOnlyfileIO(string strPath, bool status)
        {
            try
            {
                LOG.GravarLog(nameof(alteraReadOnlyfileIO), $"Alterando status ReadOnly do arquivo '{strPath}' para '{status}'.");
                var fileInfo = new System.IO.FileInfo(strPath);
                fileInfo.IsReadOnly = status;
            }
            catch (Exception ex)
            {
                ERROR_RELOAD.RestartExplorer(ex);
                LOG.GravarLog(nameof(alteraReadOnlyfileIO), "Erro ao alterar atributo ReadOnly.", ex);
                throw;
            }
        }
        #endregion

        public static void ExcluirArquivoPDM(IEdmVault7 cofre, string caminhoArquivo)
        {
            try
            {
                LOG.GravarLog(nameof(ExcluirArquivoPDM), $"Excluindo arquivo '{caminhoArquivo}' do PDM.");
                IEdmFile5 file;
                IEdmFolder5 folder;
                file = (IEdmFile5)cofre.GetFileFromPath(caminhoArquivo, out folder);

                if (file != null && folder != null)
                {
                    folder.DeleteFile(0, file.ID, true);
                }
                else
                {
                    MessageBox.Show($"Arquivo '{caminhoArquivo}' não encontrado.", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                ERROR_RELOAD.RestartExplorer(ex);
                LOG.GravarLog(nameof(ExcluirArquivoPDM), "Erro ao excluir arquivo.", ex);
                throw;
            }
        }

        #region BUSCA
        // BUSCA VALOR DE UMA VARIÁVEL DO CARTÃO
        public static List<string> FindDocumentsByVariable(IEdmVault7 vault, string variableValue, string returnType, string variableName)
        {
            var documentPaths = new List<string>();

            try
            {
                LOG.GravarLog($"{nameof(PDM)}:{nameof(FindDocumentsByVariable)}",
                    $"Iniciando busca de documentos no PDM para variável '{variableName}' com valor '{variableValue}'.");

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

                LOG.GravarLog($"{nameof(PDM)}:{nameof(FindDocumentsByVariable)}",
                    $"Busca concluída. {documentPaths.Count} documentos encontrados.");

                return documentPaths;
            }
            catch (Exception ex)
            {
                LOG.GravarLog($"{nameof(PDM)}:{nameof(FindDocumentsByVariable)}",
                    $"Erro ao buscar arquivos no PDM para variável '{variableName}' com valor '{variableValue}'.",
                    ex);

                ERROR_RELOAD.RestartExplorer(ex);

                throw new Exception("Erro ao buscar arquivos com a variável e valor especificados no PDM.", ex);
            }
        }

        // BUSCA DOCUMENTO NO COFRE
        public static List<string> FindDocuments(IEdmVault7 vault, string searchName, string FILTER_FOLDER)
        {
            var foundFiles = new List<string>();

            try
            {
                IEdmSearch6 search = (IEdmSearch6)vault.CreateSearch();
                if (search == null)
                {
                    throw new Exception("A busca não pôde ser inicializada.");
                }

                // LIMITA BUSCA NO FOLDER DA FILTER_FOLDER
                if (!string.IsNullOrEmpty(FILTER_FOLDER))
                {
                    IEdmFolder5 of = vault.GetFolderFromPath(FILTER_FOLDER);
                    search.StartFolderID = of.ID;
                }

                // Nome do arquivo com wildcard para capturar todas as extensões
                search.FileName = searchName + "%";

                // Captura os resultados da busca
                IEdmSearchResult5 result = search.GetFirstResult();

                // Itera pelos resultados encontrados
                while (result != null)
                {
                    if (result.ObjectType == EdmObjectType.EdmObject_File)
                    {
                        // Adiciona o caminho completo dos arquivos encontrados
                        foundFiles.Add(result.Path);
                    }

                    result = search.GetNextResult();
                }

                return foundFiles;
            }
            catch (COMException ex)
            {
                ERROR_RELOAD.RestartExplorer(ex);

                LOG.GravarLog($"{nameof(PDM)}:{nameof(FindDocuments)}",
                              $"Erro ao buscar arquivos no PDM: HRESULT = 0x{ex.ErrorCode:X} - {ex.Message}",
                              ex);

                throw new Exception("Erro ao buscar arquivos no PDM.", ex);
            }
        }
        #endregion
    }
}
