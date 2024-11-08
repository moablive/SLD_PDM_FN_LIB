using System;
using System.IO;

namespace PDM
{
    public static class Log
    {
        // Caminho do arquivo de log
        private static readonly string logFilePath = @"C:\TEMP\moablog.txt";

        /// <summary>
        /// Método para gravar logs em um arquivo de texto estático
        /// </summary>
        /// <param name="metodo">Nome do método onde o log está sendo gerado</param>
        /// <param name="mensagem">Mensagem do log</param>
        /// <param name="excecao">Exceção associada ao log, se houver</param>
        public static void GravarLog(string metodo, string mensagem, Exception excecao = null)
        {
            try
            {
                // Cria o conteúdo do log com data, hora e detalhes do erro
                string logContent = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} | {metodo} | {mensagem}";
                
                if (excecao != null)
                {
                    logContent += $" | Exceção: {excecao.Message} | StackTrace: {excecao.StackTrace}";
                }

                // Escreve o log no arquivo
                File.AppendAllText(logFilePath, logContent + Environment.NewLine);
            }
            catch (Exception ex)
            {
                // Em caso de erro ao gravar o log, você pode implementar uma lógica adicional de fallback se necessário
                Console.WriteLine("Erro ao gravar o log: " + ex.Message);
            }
        }
    }
}
