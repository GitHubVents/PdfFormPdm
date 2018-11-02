using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using EPDM.Interop.epdm;

namespace FormPdf
{
    public partial class Form1 : Form
    {
        public IEdmFile7 File;
        public IEdmVault7 Vault;

        string connectionString = @"Data Source=pdmsrv;Initial Catalog=SWPlusDB;User ID=airventscad;Password=1";

        SolidWorksPdmAdapter sw = new SolidWorksPdmAdapter();
        FolderBrowserDialog fbd = new FolderBrowserDialog();

        public Form1(IEdmFile7 file, IEdmVault7 vault)
        {
            File = file;
            Vault = vault;
            InitializeComponent();
            TextBox();
            GetConfiguration();
        }

        public void TextBox()
        {
            textBox1.Text = File.Name;

            File.GetFileCopy(0, 0, 0, (int)EdmGetFlag.EdmGet_Refs + (int)EdmGetFlag.EdmGet_RefsVerLatest);
        }

        public void GetConfiguration()
        {
            try
            {
                var s = File.GetConfigurations();
                IEdmPos5 pos = default(IEdmPos5);
                pos = s.GetHeadPosition();

                while (!pos.IsNull)
                {
                    string conf = s.GetNext(pos);
                    //MessageBox.Show(conf);
                    sw.GetBomShell(File, conf, Vault);
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
                throw;
            }
        }

        public void FileObj(List<BomShell> boomShellList, string pathpdf)
        {
            try
            {
               foreach (var item in boomShellList)
                {
                    if (item.FileName != "" && item.PartNumber != "")
                    {
                        IEdmFolder5 folder;

                        string pathfile = item.FolderPath + @"\" + item.FileName;
                        IEdmFile5 file = Vault.GetFileFromPath(pathfile, out folder);
                        file.GetFileCopy(0, 0, 0, (int)EdmGetFlag.EdmGet_Simple);

                        string filepath = file.GetLocalPath(folder.ID);

                        string pathdrw = item.FolderPath + @"\" + item.PartNumber + ".SLDDRW";

                        IEdmFile5 filedrw = Vault.GetFileFromPath(pathdrw, out folder);

                        if (filedrw != null)
                        {
                            filedrw.GetFileCopy(0, 0, 0, (int)EdmGetFlag.EdmGet_Simple);
                            var filepathdrw = filedrw.GetLocalPath(folder.ID);
                            int filedrwId = filedrw.ID;

                            if (CheckPdf(filedrwId, filedrw, pathpdf) != 0)
                            {
                                CheckPdf(filedrwId, filedrw, pathpdf);
                            }
                            else
                            {
                                LoadPdf lp = new LoadPdf();
                                string newpath = lp.PdfLoad(filepathdrw, true, pathpdf);

                                byte[] bytes = BinaryPdf(newpath);
                                ProcCheck(file, filedrw, filedrw.CurrentVersion, bytes);
                            }
                        }
                    }
                }

                SolidWorksAdapter.KillProcsses("SLDWORKS");
                SolidWorksAdapter.DisposeSOLID();
                MessageBox.Show("PDF файлы успешно сохранены");
                this.Hide();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
                throw;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (fbd.ShowDialog() == DialogResult.OK)
            {
                string pathpdf = fbd.SelectedPath;
                if (pathpdf != "")
                {
                    FileObj(sw.boomShellList, pathpdf);
                }
                else
                {
                    return;
                }
            }
       }

        public byte[] BinaryPdf(string path)
        {
            try
            {
                byte[] bytes;
                using (FileStream fs = new FileStream(path, FileMode.Open))
                {
                    BinaryReader reader = new BinaryReader(fs);
                    bytes = reader.ReadBytes((int)fs.Length);
                    fs.Close();
                }

                return bytes;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
                throw;
            }
        }

        public void ProcCheck(IEdmFile5 file, IEdmFile5 filedrw, int version, byte[] binar)
        {
            try
            {
                string sqlExpression = "PDFcheck";

                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();

                    SqlCommand command = new SqlCommand(sqlExpression, connection);
                    command.CommandType = CommandType.StoredProcedure;

                    command.Parameters.Add("@DocumentID", SqlDbType.Int).Value = filedrw.ID;
                    command.Parameters.Add("@Version", SqlDbType.Int).Value = filedrw.CurrentVersion;
                    command.Parameters.Add("@Blob", SqlDbType.Image).Value = binar;
                    command.ExecuteNonQuery();
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
                throw;
            }
        }

        public int CheckPdf(int documentId, IEdmFile5 filedrw, string pathpdf)
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();
                    SqlCommand command = new SqlCommand();
                    command.CommandText = "SELECT * FROM PDF WHERE DocumentId = '" + documentId + "'";
                    command.Connection = connection;
                    int temp = Convert.ToInt16(command.ExecuteScalar());

                    if (temp != 0)
                    {
                        var reader = command.ExecuteReader();

                        if (reader.Read())
                        {
                            string filename = filedrw.Name.Replace(".SLDDRW", ".pdf");

                            System.IO.File.WriteAllBytes(pathpdf + @"\" + filename, (byte[])reader["Blob"]);
                        }
                    }
                    return temp;
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
                throw;
            }
        }
    }
}