using System;
using System.Diagnostics;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace SLD
{
    public class InstanciaSW
    {
        private SldWorks swApp = null;

        public InstanciaSW(bool abreSW = true)
        {
            if (abreSW)
            {
                try
                {
                    object processSW = Activator.CreateInstance(Type.GetTypeFromProgID("SldWorks.Application"));
                    swApp = (SldWorks)processSW;

                    if (swApp == null)
                        throw new Exception("Não foi possível instanciar o SOLIDWORKS!");
                }
                catch (Exception ex)
                {
                    throw new Exception($"Erro ao iniciar o SOLIDWORKS: {ex.Message}");
                }
            }
        }

        public int ObterCorInterface()
        {
            return swApp?.GetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swSystemColorsBackground) ?? -1;
        }

        public SldWorks SWApp => swApp;
    }
}