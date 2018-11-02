//using Microsoft.Office.Interop.Outlook;

using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Windows.Forms;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace FormPdf
{
    public static class SolidWorksAdapter
    {

        private static SldWorks swApp;


        public static void DisposeSOLID()
        {
            swApp = null;
        }


        public static SldWorks SldWoksAppExemplare
        {
            get
            {
                if (swApp == null)
                {
                    //swApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application");
                    InitializeSolidWorks();
                }
                return swApp;
            }
        }


        [DllImport("ole32.dll")]
        static extern int CreateBindCtx(uint reserved, out IBindCtx ppbc);
        [DllImport("ole32.dll")]
        private static extern void GetRunningObjectTable(int reserved, out IRunningObjectTable prot);
        private static void InitializeSolidWorks()
        {
            string monikerName = "SolidWorks_PID_";
            object app;
            IBindCtx context = null;
            IRunningObjectTable rot = null;
            IEnumMoniker monikers = null;

            try
            {
                CreateBindCtx(0, out context);

                context.GetRunningObjectTable(out rot);
                rot.EnumRunning(out monikers);

                IMoniker[] moniker = new IMoniker[1];

                while (monikers.Next(1, moniker, IntPtr.Zero) == 0)
                {
                    var curMoniker = moniker.First();

                    string name = null;

                    if (curMoniker != null)
                    {
                        try
                        {
                            curMoniker.GetDisplayName(context, null, out name);
                        }
                        catch (UnauthorizedAccessException ex)
                        {
                            MessageBox.Show("Failed to get SolidWorks_PID." + "\t" + ex);
                        }
                    }
                    if (name.Contains(monikerName))
                    {
                        rot.GetObject(curMoniker, out app);
                        swApp = (SldWorks)app;
                        swApp.Visible = true;
                        return;
                    }
                }
                string progId = "SldWorks.Application";

                Type progType = Type.GetTypeFromProgID(progId);
                app = Activator.CreateInstance(progType) as SldWorks;
                swApp = (SldWorks)app;
                swApp.Visible = true;
                return;
            }
            finally
            {
                if (monikers != null)
                {
                    Marshal.ReleaseComObject(monikers);
                }
                if (rot != null)
                {
                    Marshal.ReleaseComObject(rot);
                }
                if (context != null)
                {
                    Marshal.ReleaseComObject(context);
                }
            }
        }


        /// <summary>
        /// Closing all opened documents
        /// </summary>
        public static void CloseAllDocuments()
        {
            try
            {
                var modelDocs = SldWoksAppExemplare.GetDocuments();
                foreach (var eachModelDoc in modelDocs)
                {
                    eachModelDoc.Close();
                }
                MessageBox.Show("\t\tClosed all opened documents");
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("\t\tFailed close documents " + ex);
            }
        }

        /// <summary>
        /// Closing opened document
        /// </summary>
        /// <param name="swModel"></param>
        public static void CloseDocument(IModelDoc2 swModel)
        {

            try
            {
                SldWoksAppExemplare.CloseDoc(swModel.GetTitle().ToLower().Contains(".sldprt") ? swModel.GetTitle() : swModel.GetTitle() + ".sldprt");
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("Failed close document " + swModel.GetTitle() + "\t" + ex);
            }
        }

        /// <summary>
        /// Closing all opened documents and exist from SolidWorks Application
        /// </summary>
        public static void CloseAllDocumentsAndExit()
        {
            try
            {
                CloseAllDocuments();
                SldWoksAppExemplare.ExitApp();
                MessageBox.Show("Exit from  SolidWorks Application");
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("Failed exit from  SolidWorks Application");
            }
        }

        /// <summary>
        /// Check is sheet metal part
        /// </summary>
        /// <param name="swPart"></param>
        public static bool IsSheetMetalPart(IPartDoc swPart)
        {
            bool isSheetMetal = false;
            try
            {
                var bodies = swPart.GetBodies2((int)swBodyType_e.swSolidBody, false);
                foreach (Body2 body in bodies)
                {
                    isSheetMetal = body.IsSheetMetal();

                    MessageBox.Show("Check is sheet metal part; returned " + isSheetMetal);

                    if (isSheetMetal)
                    {
                        return true;
                    }
                }
            }
            catch (System.Exception)
            {
                MessageBox.Show("Failed check is sheet metal part; returned " + false);
                return false;
            }
            return isSheetMetal;
        }

        //public static ModelDoc2 OpenDocument(string path, swDocumentTypes_e documentType, string configuration = "")
        //{
        //    if (!File.Exists(path))
        //    {
        //        MessageObserver.Instance.SetMessage($"Error at open solid works document {path} file not exists. Maybe it is virtual document", MessageType.Error);
        //        throw new Exception($"Error at open solid works document {path} file not exists. Maybe it is virtual document" );
        //    }
        //    int errors = 0, warnings = 0;
        //    int openDocOptions = (int)swOpenDocOptions_e.swOpenDocOptions_ReadOnly;
        //    if (documentType == swDocumentTypes_e.swDocASSEMBLY) {  openDocOptions += (int)swOpenDocOptions_e.swOpenDocOptions_LoadModel; }

        //    var SolidWorksDocumentument = SldWoksAppExemplare.OpenDoc6(path, (int)documentType, openDocOptions, configuration, ref errors, ref warnings);

        //    if (errors != 0)
        //    {
        //        MessageObserver.Instance.SetMessage($"Error at open solid works document {path}; error code {errors }, description error { (swFileLoadError_e)errors }" ) ;
        //    }
        //    if (warnings != 0)
        //    {
        //        MessageObserver.Instance.SetMessage("Warning at open solid works document: code {" + warnings + "}, description warning {" + (swFileLoadWarning_e)errors + "}");
        //    }
        //    return SolidWorksDocumentument;
        //}

        public static ModelDoc2 OpenDocument(string path, swDocumentTypes_e documentType, string configuration = "")
        {
            //if (!File.Exists(path))
            //{
            //    MessageObserver.Instance.SetMessage($"Error at open solid works document {path} file not exists. Maybe it is virtual document", MessageType.Error);
            //    throw new System.Exception($"Error at open solid works document {path} file not exists. Maybe it is virtual document");
            //}

            DocumentSpecification swDocSpecification;

            swDocSpecification = SolidWorksAdapter.SldWoksAppExemplare.GetOpenDocSpec(path);
            swDocSpecification.ReadOnly = false;
            swDocSpecification.DocumentType = (int)documentType;
            swDocSpecification.UseLightWeightDefault = false;
            swDocSpecification.LightWeight = false;
            swDocSpecification.Silent = true;
            swDocSpecification.IgnoreHiddenComponents = false;
            swDocSpecification.ViewOnly = false;
            swDocSpecification.InteractiveAdvancedOpen = false;
            swDocSpecification.ConfigurationName = configuration;

            ModelDoc2 SolidWorksDocumentument = null;

            if (swDocSpecification != null)
            {
                SolidWorksDocumentument = SldWoksAppExemplare.OpenDoc7(swDocSpecification);
                if (SolidWorksDocumentument == null)
                {
                    MessageBox.Show("Failed to open document " + path + System.Environment.NewLine + "Error : " + (swFileLoadError_e)swDocSpecification.Error + System.Environment.NewLine +
                                    " Warning: " + (swFileLoadWarning_e)swDocSpecification.Warning);
                }
                else
                {
                    if (documentType == swDocumentTypes_e.swDocASSEMBLY)
                    {
                        int resolved = ((AssemblyDoc)SolidWorksDocumentument).ResolveAllLightWeightComponents(false);
                        MessageBox.Show("Resolve document with result: " + (swComponentResolveStatus_e)resolved);
                    }
                }
            }
            else
            {
                MessageBox.Show("Failed to get specification on " + path + System.Environment.NewLine + "Error : " + (swFileLoadError_e)swDocSpecification.Error + System.Environment.NewLine +
                                " Warning: " + (swFileLoadWarning_e)swDocSpecification.Warning);
            }

            return SolidWorksDocumentument;
        }

        public static ModelDoc2 AcativeteDoc(string docTitle)
        {
            int errors = 0;
            ModelDoc2 modelDoc = SolidWorksAdapter.SldWoksAppExemplare.ActivateDoc3(docTitle, true, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, errors);

            if (errors != 0)
            {
                MessageBox.Show("Exeption at activate solid works document: code {" + errors + "}, description error {" + (swActivateDocError_e)errors + "}");
            }
            return modelDoc;
        }

        /// <summary>
        /// Convert  ModelDoc2 to AssemblyDoc and resolve all light weight components
        /// </summary>
        /// <param name="document"></param>
        /// <returns></returns>
        public static AssemblyDoc ToAssemblyDocument(ModelDoc2 document)
        {
            swComponentResolveStatus_e res = swComponentResolveStatus_e.swResolveOk;
            AssemblyDoc swAsm = null;
            if ((int)swDocumentTypes_e.swDocASSEMBLY == document.GetType())
            {
                swAsm = (AssemblyDoc)document;
                //res = (swComponentResolveStatus_e)swAsm.ResolveAllLightWeightComponents(false);
                //MessageObserver.Instance.SetMessage("Resolve All LightWeight Components: code {" + res + "}");
            }
            else
            {
                MessageBox.Show("Unable to cast SolidWorksDocument to AssemblyDoc type, cause it's paart. " + document.GetTitle());
            }
            return swAsm;
        }

        public static DrawingDoc ToDrawingDoc(ModelDoc2 document)
        {
            DrawingDoc drw = (DrawingDoc)document;
            return drw;
        }

        public static int ToInt(this double value)
        {
            return Convert.ToInt32(value);
        }


        public static bool CheckExcistance(string path)
        {

            if (File.Exists(path))
            {
                MessageBox.Show("Part exists and will be replaced with path: " + path + ". Or ASM exsists and will be opened.");
                return true;
            }

            return false;
        }


        //public static void OutLookSendMeALog(string pathToLogFile, string message)
        //{
        //    Microsoft.Office.Interop.Outlook.Application lookApp = new Microsoft.Office.Interop.Outlook.Application();

        //    MailItem mail = null;
        //    Recipients mailrecipients = null;
        //    Recipient recip = null;


        //    try
        //    {
        //        mail = lookApp.CreateItem(OlItemType.olMailItem);
        //        mail.Subject = message;
        //        mail.Attachments.Add(pathToLogFile, OlAttachmentType.olByValue, 1);
        //        mailrecipients = mail.Recipients;

        //        recip = mailrecipients.Add("d.dagovets@vents.kiev.ua");

        //        recip.Resolve();
        //        mail.Send();
        //    }
        //    catch (System.Exception ex)
        //    {
        //        MessageObserver.Instance.SetMessage("Failed to send Log message.", MessageType.Error);
        //    }
        //    finally
        //    {
        //        if (mailrecipients != null) Marshal.ReleaseComObject(mailrecipients);
        //        if (recip != null) Marshal.ReleaseComObject(recip);
        //        if (mail != null) Marshal.ReleaseComObject(mail);
        //    }
        //}

        public static void KillProcsses(string name)
        {
            var processes = System.Diagnostics.Process.GetProcessesByName(name);
            foreach (var process in processes)
            {
                process.Kill();
            }
        }

    }
}