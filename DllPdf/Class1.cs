using System.Runtime.InteropServices;
using EPDM.Interop.epdm;
using FormPdf;


namespace DllPdf
{
    [Guid("FD25A60D-CB75-43C7-8619-84E6BE0103DA"), ComVisible(true)]
    public class Class1 : IEdmAddIn5
    {
        public void GetAddInInfo(ref EdmAddInInfo poInfo, IEdmVault5 poVault, IEdmCmdMgr5 poCmdMgr)
        {
            poInfo.mbsAddInName = "AddinPdf";
            poInfo.mbsCompany = "SOLIDWORKS Corporation";
            poInfo.mbsDescription = "Adds menu command items";
            poInfo.mlAddInVersion = 1;

            poInfo.mlRequiredVersionMajor = 10;
            poInfo.mlRequiredVersionMinor = 0;

            //poCmdMgr.AddCmd(100, "Выгрузить Pdf-file", (int)EdmMenuFlags.EdmMenu_OnlyMultipleSelection);
            poCmdMgr.AddCmd(100, "Выгрузить Pdf-file", (int)EdmMenuFlags.EdmMenu_OnlySingleSelection);
        }

        public void OnCmd(ref EdmCmd poCmd, ref EdmCmdData[] ppoData)
        {
            switch (poCmd.meCmdType)
            {
                case EdmCmdType.EdmCmd_Menu:
                    OnMenu(ref poCmd, ref ppoData);
                    break;
            }
        }

        private void OnMenu(ref EdmCmd poCmd, ref EdmCmdData[] ppoData)
        {
            int i;
            for (i = 0; i < ppoData.Length; i++)
            {
                if (((EdmCmdData)ppoData.GetValue(0)).mlObjectID1 != 0)
                {
                    var vault = (IEdmVault7)poCmd.mpoVault;


                    var file = (IEdmFile7)vault.GetObject(EdmObjectType.EdmObject_File, ((EdmCmdData)ppoData.GetValue(i)).mlObjectID1);
                    var folder = (IEdmFolder5)vault.GetObject(EdmObjectType.EdmObject_Folder, ((EdmCmdData)ppoData.GetValue(i)).mlObjectID3);

                    Form1 f = new Form1(file, vault);
                    f.ShowDialog();


                }
            }
        }

    }
}
