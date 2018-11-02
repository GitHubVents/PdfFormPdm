using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using View = SolidWorks.Interop.sldworks.View;

namespace FormPdf
{
    public class LoadPdf
    {
        public ModelDoc2 SwModel;
        public DrawingDoc SwDraw;

        public string PdfLoad(string filepath, bool deep, string pathpdf)
        {
            try
            {
                SwModel = SolidWorksAdapter.SldWoksAppExemplare.OpenDoc6(filepath, (int)swDocumentTypes_e.swDocDRAWING,
                    (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", 0, 0);
                SwDraw = (DrawingDoc)SwModel;

                SwDraw.ResolveOutOfDateLightWeightComponents();
                SwDraw.ForceRebuild();

                var vSheetName = (string[])SwDraw.GetSheetNames();


                foreach (var name in vSheetName)
                {
                    if (name != null)
                    {
                        SwDraw.ResolveOutOfDateLightWeightComponents();
                        var swSheet = SwDraw.Sheet[name];

                        SwDraw.ActivateSheet(swSheet.GetName());

                        if ((swSheet.IsLoaded()))
                        {
                            var sheetviews = (object[])swSheet.GetViews();

                            if (sheetviews != null)
                            {
                                var firstView = (View)sheetviews[0];

                                firstView.SetLightweightToResolved();
                            }

                        }

                        if (!deep) continue;
                        var views = (object[])swSheet.GetViews();

                        if (views != null)
                        {
                            foreach (var drwView in views.Cast<View>())
                            {
                                drwView.SetLightweightToResolved();
                            }
                        }


                    }

                }

                var errors = 0;
                var warnings = 0;
                var newpath = pathpdf + "\\" + Path.GetFileNameWithoutExtension(SwModel.GetPathName()) + ".pdf";
                //var newpath = Path.GetFullPath(SwModel.GetPathName().Replace("slddrw".ToUpper(), "pdf"));// + Path.GetFileNameWithoutExtension(SwModel.GetPathName()) + ".pdf";
                SwModel.Extension.SaveAs(newpath, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_UpdateInactiveViews, null, ref errors, ref warnings);
                SolidWorksAdapter.SldWoksAppExemplare.CloseDoc(Path.GetFileNameWithoutExtension(new FileInfo(newpath).FullName));
                //SolidWorksAdapter.KillProcsses("SLDWORKS");
                
                return newpath;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
                throw;
            }
        }
    }
}