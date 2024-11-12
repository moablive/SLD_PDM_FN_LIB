using System;
using System.Diagnostics;
using System.Threading;

namespace SLD_PDM
{
    public static class ERROR_RELOAD
    {
        /// <summary>
        /// Método para reiniciar o explorer.exe
        /// </summary>
        public static void RestartExplorer()
        {
            try
            {
                // Fecha o explorer.exe
                Process.Start("cmd.exe", "/C taskkill /F /IM explorer.exe");

                // Pequeno atraso para garantir que o explorer.exe seja encerrado
                Thread.Sleep(1000);

                // Reinicia o explorer.exe
                Process.Start("cmd.exe", "/C start explorer.exe");

                // Atraso adicional para garantir que o explorer.exe reinicie corretamente
                Thread.Sleep(500);
            }
            catch (Exception ex)
            {
                DebugSKA.Log.GravarLog($"{typeof(PDM_FN).Name.ToUpper() + ":" + ":" + nameof(RestartExplorer)}", "ERRO - Ao Fazer Checkout do ARQUIVO do PDM. Ative o DEBUG para mais detalhes.", ex);
                throw new Exception(ex.Message);
            }
        }
    }
}
