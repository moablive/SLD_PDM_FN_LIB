using System;
using System.Diagnostics;
using System.Threading;

namespace SLD_PDM.PDM
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
                LOG.GravarLog($"{typeof(ERROR_RELOAD).Name.ToUpper() + ":" + ":" + nameof(RestartExplorer)}", "ERRO - NO PDM", ex);
                throw new Exception(ex.Message);
            }
        }
    }
}
