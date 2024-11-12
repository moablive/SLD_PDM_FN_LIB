// System
using System;

// SKA
using DebugSKA;

// SLD DLL
using SolidWorks.Interop.sldworks;

namespace SLD
{
    public class SLD
    {
        // VAR swApp
        private SldWorks swApp = null;

        // RETURN swApp
        public SldWorks SWApp
        {
            get { return swApp; }
        }

        // ABRE SLD
        public SLD()
        {
            try
            {
                object processSW = Activator.CreateInstance(Type.GetTypeFromProgID("SldWorks.Application"));
                swApp = (SldWorks)processSW;
                swApp.Visible = true; // Deixa o SolidWorks visível
            }
            catch (Exception ex)
            {
                Log.GravarLog($"{typeof(SLD).Name.ToUpper()}:{nameof(SLD)}",
                    "ERRO - Ao Instanciar SldWorks. Ative o DEBUG para mais detalhes.",
                    ex);
            }
        }

        public void FecharSLD()
        {
            try
            {
                if (swApp != null)
                    swApp.ExitApp();
            }
            catch (Exception ex)
            {
                Log.GravarLog($"{typeof(SLD).Name.ToUpper()}:{nameof(FecharSLD)}",
                    "ERRO - Ao Fechar SldWorks. Ative o DEBUG para mais detalhes.",
                    ex);
            }
        }
    }
}
