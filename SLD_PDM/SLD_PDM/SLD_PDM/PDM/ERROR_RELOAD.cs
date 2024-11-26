using System;
using System.Diagnostics;
using System.Reflection;
using System.Threading;
using System.Windows.Forms;

namespace SLD_PDM.PDM
{
    public static class ERROR_RELOAD
    {
        /// <summary>
        /// Método para reiniciar o explorer.exe
        /// </summary>
        public static void RestartExplorer(Exception ex)
        {
            try
            {
                // Obtém o método que originou o erro, se disponível
                string originMethod = ex.TargetSite != null ? ex.TargetSite.Name : "Método desconhecido";
                string fullMessage = $"Ocorreu um erro no PDM no método '{originMethod}': {ex.Message}\n\nDeseja reiniciar o explorer.exe para corrigir o problema?";

                // Exibe uma caixa de diálogo para confirmação do usuário
                DialogResult result = MessageBox.Show(
                    fullMessage,
                    "Reiniciar Explorer",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question
                );

                // Se o usuário escolher "Yes", procede com o reinício do explorer.exe
                if (result == DialogResult.Yes)
                {
                    // Fecha o explorer.exe
                    Process.Start("cmd.exe", "/C taskkill /F /IM explorer.exe");

                    // Pequeno atraso para garantir que o explorer.exe seja encerrado
                    Thread.Sleep(1000);

                    // Reinicia o explorer.exe
                    Process.Start("cmd.exe", "/C start explorer.exe");

                    // Atraso adicional para garantir que o explorer.exe reinicie corretamente
                    Thread.Sleep(500);

                    // Mensagem de confirmação para o usuário
                    MessageBox.Show(
                        "O explorer.exe foi reiniciado com sucesso.",
                        "Operação Concluída",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information
                    );
                }
                else
                {
                    // Mensagem informando que o reinício foi cancelado
                    MessageBox.Show(
                        "A operação foi cancelada pelo usuário. O sistema pode permanecer instável.",
                        "Operação Cancelada",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning
                    );
                }
            }
            catch (Exception innerEx)
            {
                // Loga o erro e exibe informações sobre o método que causou o erro original
                string errorOrigin = ex.TargetSite != null ? ex.TargetSite.Name : "Método desconhecido";
                LOG.GravarLog($"{nameof(ERROR_RELOAD).ToUpper()}:{nameof(RestartExplorer)}", $"ERRO - NO PDM no método '{errorOrigin}'.", innerEx);

                // Exibe uma mensagem de erro para o usuário
                MessageBox.Show(
                    $"Erro ao reiniciar o explorer.exe: {innerEx.Message}\nErro original no método: {errorOrigin}",
                    "Erro",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );

                throw new Exception(innerEx.Message);
            }
        }
    }
}
