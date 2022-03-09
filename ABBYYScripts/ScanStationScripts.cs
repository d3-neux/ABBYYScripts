using ABBYY.ScanStation;
using System.IO;

namespace ABBYYScripts
{
    public class ScanStationScripts
    {

        public static IExportBatch SplitAndExportImagesAndMetadata(IExportBatch ExportBatch, string rootPath)
        {
            IImageSavingOptions exportImageSavingOptions = ExportBatch.NewImageSavingOptions();
            exportImageSavingOptions.AddProperFileExt = true;
            exportImageSavingOptions.Format = "tif";
            //exportImageSavingOptions.Resolution = 300;
            exportImageSavingOptions.ShouldOverwrite = true;
            // var pathAFC  = "\\\\192.168.1.15\\_afc\\_Export\\Conecel\\Import\\";
            // var pathAtt = "\\\\192.168.1.15\\_afc\\_Export\\Conecel\\Temp\\" + ExportBatch.Name + "\\";
            //var pathAFC  = "\\\\1s1-flexiapppro\\_afc\\_Import\\Facturas\\FacturasGYE\\";
            //var pathAFC  = "\\\\1s1-flexiapppro\\_afc\\_Import\\Facturas\\FacturasUIO\\";
            // var pathAtt   = "\\\\1s1-flexiapppro\\_afc\\_Import\\Attachments\\" + ExportBatch.Name + "\\";
            // var pathAFC      = "\\\\1s1-flexiapppro\\_afc\\_Import\\Facturas\\" + ExportBatch.Type.Name + "\\";

            //string  pathAtt = "\\\\192.168.100.125\\_afc\\_Import\\Fase 2\\Attachments\\";
            //string  pathAFC = "\\\\192.168.100.125\\_afc\\_Import\\Fase 2\\" + ExportBatch.Type.Name + "\\";

            string pathAtt = $"{rootPath}Attachments\\";
            string pathAFC = $"{rootPath}{ExportBatch.Type.Name}\\";


            string finalPathAFC = pathAFC;

            if (!Directory.Exists(pathAFC))
            {
                Directory.CreateDirectory(pathAFC);
            }

            pathAFC += "EXP_" + ExportBatch.Name + "\\";
            finalPathAFC += ExportBatch.Name + "\\";

            if (!Directory.Exists(pathAFC))
            {
                Directory.CreateDirectory(pathAFC);
            }

            if (!Directory.Exists(pathAtt))
            {
                Directory.CreateDirectory(pathAtt);
            }

            pathAtt += ExportBatch.Name + "\\";

            if (!Directory.Exists(pathAtt))
            {
                Directory.CreateDirectory(pathAtt);
            }


            bool isExportable = true;
            string errStr = "Elementos con error: ";

            //var PARAM_FechaDig   = ExportBatch.RegistrationProperties(0).Value.replace(/[/\\?%*:|"<>]/g, ' ');
            //var PARAM_IDEstacion = ExportBatch.RegistrationProperties(1).Value.replace(/[/\\?%*:|"<>]/g, ' ');
            //var PARAM_UsuarioDig = ExportBatch.RegistrationProperties(2).Value.replace(/[/\\?%*:|"<>]/g, ' ');

            string caja = "";
            string paquete = "";


            using (StreamWriter writer = new StreamWriter(pathAtt + "RegParams.txt", false))
            {
                for (int i = 0; i < ExportBatch.RegistrationProperties.Count; i++)
                {
                    string prop = ExportBatch.RegistrationProperties[i].Value;
                    string name = ExportBatch.RegistrationProperties[i].Name;

                    if (name == "Caja")
                        caja = prop;

                    if (name == "Paquete")
                        paquete = prop;


                    writer.WriteLine(name + "^" + prop);
                }
            }

            string nombreLote = "PRO_" + caja + "_" + paquete;

            if (ExportBatch.Name.Substring(0,15) != nombreLote)
            {

                if (Directory.Exists(pathAtt)) Directory.Delete(pathAtt, true);
                if (Directory.Exists(pathAFC)) Directory.Delete(pathAFC, true);


                ExportBatch.Result.ErrorMessage = "Nombre de lote " + ExportBatch.Name + " incorrecto, debe ser " + nombreLote;
                ExportBatch.Result.Succeeded = false;
                return null;


            }




            for (int i = 0; i < ExportBatch.Children.Count; i++)
            {
                IExportBatchItem item = ExportBatch.Children[i];
                if (!item.IsDocument)
                {
                    isExportable = false;
                    errStr += (i + 1) + ", ";
                }
            }

            if (isExportable == false)
            {
                ExportBatch.Result.ErrorMessage = errStr;
                ExportBatch.Result.Succeeded = false;
            }
            else
            {
                int exportedPageCount = ExportBatch.CalcExportedPageCount();
                ExportBatch.NotifyProcessingProgress(exportedPageCount, ExportBatch.PageCount - exportedPageCount);
                for (int i = 0; i < ExportBatch.Children.Count; i++)
                {
                    IExportBatchItem item = ExportBatch.Children[i];
                    if (item.IsDocument && !item.IsExported)
                    {
                        string docDir = pathAFC;
                        string docDirAtt = pathAtt + item.Name;
                        if (!Directory.Exists(docDirAtt))
                        {
                            Directory.CreateDirectory(docDirAtt);
                        }
                        for (int j = 0; j < item.Children.Count; j++)
                        {
                            string dicExport = "";
                            if (j == 0)
                                dicExport = docDir;
                            else
                                dicExport = docDirAtt;
                            ExportBatch.CheckCanceled();
                            item.IsExported = true;
                            IExportBatchItem page = item.Children[j];
                            if (!page.IsExported)
                            {

                                int pageNumber = j + 1;

                                string pageNumberStr = (pageNumber < 10 ? ("0" + pageNumber.ToString()) : pageNumber.ToString());

                                //string pageName = page.Name + "-" + item.Children.Count;

                                string pageName = pageNumberStr + "-" + item.Children.Count;

                                page.SaveAs(dicExport + "\\" + ExportBatch.Name + "^" + item.Name + "^" + pageName, exportImageSavingOptions);
                                exportedPageCount++;
                                ExportBatch.NotifyProcessingProgress(exportedPageCount, ExportBatch.PageCount - exportedPageCount);
                                page.IsExported = true;
                            }
                        }
                        item.IsExported = true;
                    }
                }

                //SE RENOMBRA EL DIRECTORIO PARA DE ESTA MANERA EVITAR IMPORTACION DE AFC ANTES QUE ESTA EXPORT FINALICE
                Directory.Move(pathAFC, finalPathAFC);

            }
            return ExportBatch;
        }



       



    }
}
