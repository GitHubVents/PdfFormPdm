using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using EPDM.Interop.epdm;
using EPDM.Interop.EPDMResultCode;

namespace FormPdf
{
    public class SolidWorksPdmAdapter
    {
       public List<BomShell> boomShellList = new List<BomShell>();

        public IEnumerable<BomShell> GetBomShell(IEdmFile7 file, string bomConfiguration, IEdmVault7 vault)
        {
            try
            {
                var bomView = file.GetComputedBOM(22, file.CurrentVersion, bomConfiguration, (int)EdmBomFlag.EdmBf_ShowSelected);

                if (bomView == null)
                {
                    throw new Exception("Computed BOM it can not be null");
                }
                object[] bomRows;
                EdmBomColumn[] bomColumns;
                bomView.GetRows(out bomRows);
                bomView.GetColumns(out bomColumns);


                DataTable bomTable = new DataTable();
                foreach (EdmBomColumn bomColumn in bomColumns)
                {
                    bomTable.Columns.Add(new DataColumn { ColumnName = bomColumn.mbsCaption });
                }
                for (var i = 0; i < bomRows.Length; i++)
                {
                    var cell = (IEdmBomCell)bomRows.GetValue(i);

                    bomTable.Rows.Add();

                    for (var j = 0; j < bomColumns.Length; j++)
                    {
                        EdmBomColumn column = (EdmBomColumn)bomColumns.GetValue(j);
                        object value;
                        object computedValue;
                        string config;
                        bool readOnly;
                        cell.GetVar(column.mlVariableID, column.meType, out value, out computedValue, out config, out readOnly);

                        if (value != null)
                        {
                            bomTable.Rows[i][j] = value;
                        }
                        else
                        {
                            bomTable.Rows[i][j] = null;
                        }
                    }
                }
                return BomTableToBomList(bomTable);
            }
            catch (COMException ex)
            {
                MessageBox.Show("Failed get bom shell " + (EdmResultErrorCodes_e)ex.ErrorCode + ". Укажите вид PDM или тип спецификации");
                throw ex;
            }
        }

        private IEnumerable<BomShell> BomTableToBomList(DataTable table)
        {
            try
            {
                //ColumnAdd(dataGridView2);
                //List<BomShell> boomShellList = new List<BomShell>(dt.Rows.Count);

                boomShellList.AddRange(from DataRow eachRow in table.Rows
                    select eachRow.ItemArray
                    into values
                    select new BomShell
                    {
                        PartNumber = values[0].ToString(),
                        Description = values[1].ToString(),
                        IdPdm = Convert.ToInt32(values[2]),
                        Configuration = values[3].ToString(),
                        Version = Convert.ToInt32(values[4]),
                        FileName = values[5].ToString(),
                        FolderPath = values[6].ToString(),
                        ObjectType = values[7].ToString(),
                        Partition = values[8].ToString()

                    });

                return boomShellList;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
                throw;
            }
        }

    }
}
