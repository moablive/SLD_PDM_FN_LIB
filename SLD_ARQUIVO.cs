//System;
using System;

// SolidWorks DLLs
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace SLD
{
    public class arquivoSLD
    {
        private SldWorks swApp;
        private IModelDoc2 swModel;

        public arquivoSLD(SldWorks sldWorksApp)
        {
            swApp = sldWorksApp;
        }

        public IModelDoc2 AbrirDocumento(string caminhoArquivo, int tipoDocumento, bool visivel)
        {
            int errors = 0, warnings = 0;

            swApp.Visible = visivel;

            // Abre o documento no SolidWorks
            swModel = swApp.OpenDoc6(caminhoArquivo, tipoDocumento, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref errors, ref warnings);

            if (swModel == null)
            {
                throw new Exception("Não foi possível instanciar o arquivo no SolidWorks.");
            }

            return swModel;
        }

        public void setPropriedade(string Valor, string Config, string pName)
        {
            try
            {
                ModelDocExtension swModelDocExt = default(ModelDocExtension);
                CustomPropertyManager swCustProp = default(CustomPropertyManager);

                swModelDocExt = swModel.Extension;

                // Get the custom property data
                swCustProp = swModelDocExt.get_CustomPropertyManager(Config);
                try
                {
                    swCustProp.Delete(pName);
                }
                catch (Exception ex)
                {
                    Log.GravarLog($"{typeof(arquivoSLD).Name.ToUpper() + ":" + nameof(setPropriedade)}",
                        "ERRO - Ao setPropriedade. Ative o DEBUG para mais detalhes.",
                        ex);
                }

                swCustProp.Add3(pName, 30, Valor, 1);
            }
            catch (Exception ex)
            {
                Log.GravarLog($"{typeof(arquivoSLD).Name.ToUpper() + ":" + nameof(setPropriedade)}",
                    "ERRO - Ao setPropriedade. Ative o DEBUG para mais detalhes.",
                    ex);
            }
        }
    }
}