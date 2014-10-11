using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Reflection;

namespace DataExportTool
{
    public partial class Form1 : Form
    {
        private int page = 0;
        const int pageCount = 40000;
        private int currentpage = 0;
        public Form1()
        {
            InitializeComponent();
            this.gridControl1.DataSource = myBind;
            this.bindingNavigator1.BindingSource = this.myBind;
            this.barEditItem1.EditValue = DateTime.Now;
            this.barEditItem2.EditValue = DateTime.Now;
        }
        private DataSet myData;
        private BindingSource myBind = new BindingSource();
        private string BOF_LF_RH_SM = "SELECT DISTINCT heat_id as \"物料号\",steel_grade as \"钢种\", product_day as \"生产时间\" FROM sm_bof_heat WHERE heat_id IS NOT NULL and product_day >= to_date('{0}', 'YYYYMMDDHH24MI') and product_day <= to_date('{1}', 'YYYYMMDDHH24MI') order by product_day";
        private string CCM_SM = "select distinct heat_id as \"入口物料号\", slab_no as \"出口物料号\", steel_grade as \"钢种\", cut_date as \"生产时间\", thickness as \"厚度\", width as \"宽度\" from slab_l2_reports where cut_date >= to_date('{0}', 'YYYYMMDDHH24MI') and cut_date <= to_date('{1}', 'YYYYMMDDHH24MI') order by cut_date";
        private string HRM_SM = "select slab_id as \"入口物料号\", coil_id as \"出口物料号\",c_custom_sgc_upd as \"钢种\", rolled_time as \"生产时间\", b.f_fmdelthktarg as \"厚度\", b.f_fmdelwidtarg as \"宽度\" from hrm_l2_coilreports a, hrm_l2_pdi b where a.slab_id = trim(b.c_slabid) and rolled_time >= to_date('{0}', 'YYYYMMDDHH24MI') and rolled_time <= to_date('{1}', 'YYYYMMDDHH24MI') order by rolled_time";
        private string matProcess = " select in_mat_id1, in_mat_id2, out_mat_id, process_code from process_mat_pedigree t ";
        private string workshop = "select device_no,workshop_no,process_code from equipment_code";
        private string HRM_MATTRACK_TIME = "select mat_no, device_no, start_time, stop_time from HRM_MATTRACK_TIME ";
        private string SM_MATTRACK_TIME = "select mat_no, device_no, start_time, stop_time from SM_MATTRACK_TIME ";
        private string CRM_MATTRACK_TIME = "select mat_no, device_no, start_time, stop_time from CRM_MATTRACK_TIME ";
        private string deviceconfig = "select * from device_area_config order by display_num";
        private string SM_BOF_HEAT = "select * from SM_BOF_HEAT";
        private string sm_elem_ana = "SELECT * FROM  sm_elem_ana WHERE SAMPLETIME >= to_date('{0}', 'YYYYMMDDHH24MI') and SAMPLETIME <= to_date('{1}', 'YYYYMMDDHH24MI') order by SAMPLETIME";
        private string SM_LF_HEAT = "SELECT * FROM SM_LF_HEAT";
        private string sm_TEMPTURE = "select * FROM sm_TEMPTURE where MEASURE_TIME >= to_date('{0}', 'YYYYMMDDHH24MI') and MEASURE_TIME <= to_date('{1}', 'YYYYMMDDHH24MI') order by MEASURE_TIME";
        private string SM_RH_HEAT = "SELECT * FROM SM_RH_HEAT";
        private string SLAB_L2_REPORTS = "SELECT * FROM  SLAB_L2_REPORTS";
        private string  sm_ccm_heat = "SELECT * FROM sm_ccm_heat";
        private string hrm_l2_FCEREPORTS ="SELECT * FROM hrm_l2_FCEREPORTS";
        private string HRM_L2_PDI = "SELECT * FROM HRM_L2_PDI";
        private string hrm_coil_setup = "SELECT * FROM hrm_coil_setup";
        private string HRM_L2_COILREPORTS = "select * FROM HRM_L2_COILREPORTS";
        private string hrm_l2_coilsurfreport = "SELECT * FROM hrm_l2_coilsurfreport";
        private string hrm_l2_coildefects = "SELECT * FROM hrm_l2_coildefects ";
        private string hrm_l2_ctcsetup = "SELECT * FROM hrm_l2_ctcsetup";
        private void barButtonItem1_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if(this.barEditItem1.EditValue==null||this.barEditItem2.EditValue==null)
            {
                MessageBox.Show("时间不能为空");
                return;
            }
            myData = new DataSet();
            this.page = 0;

                var temp = getData((DateTime)this.barEditItem1.EditValue, (DateTime)this.barEditItem2.EditValue);
                //temp.TableName = "Page" + page.ToString();
                myData.Tables.Add(temp);


            this.gridView1.Columns.Clear();
            this.myBind.DataSource = myData;
            currentpage = 0;
            this.myBind.DataMember = myData.Tables[currentpage].TableName;
            initSummary();
        }
        private DataTable getData(DateTime start,DateTime end)
        {
            string con = "Provider=MSDAORA.1;Password=lyq;User ID=lyq;Data Source=lyq;Persist Security Info=True";
            DataTable result = new DataTable();
            using (OleDbConnection conn = new OleDbConnection(con))
            {

                OleDbDataAdapter adp = new OleDbDataAdapter(string.Format(hrm_l2_ctcsetup, start.ToString("yyyyMMddHHmm"), end.ToString("yyyyMMddHHmm"), page * pageCount, (page + 1) * pageCount), conn);
                while (true)
                {
                    if (adp.Fill(page * pageCount, pageCount, result)< pageCount)
                        break;
                    page++;
                }
            }
            return result;
        }
        /*
        private void matProcess()
        {
            List<MaterialInfo> track = new List<MaterialInfo>();

            Dictionary<string, string> tabledef = new Dictionary<string, string>();

            tabledef["LY2250"] = "HRM_MATTRACK_TIME";
            tabledef["LY210"] = "SM_MATTRACK_TIME";
            tabledef["LYCRM"] = "CRM_MATTRACK_TIME";

            using (OleDbConnection connection = new OleDbConnection(ConnectionString.LYQ_DB))
            {
                connection.Open();

                string prevMatId = "";

                while (prevMatId != matId)
                {
                    string sql = string.Format(" select in_mat_id1, in_mat_id2, out_mat_id, process_code " +
                                                " from process_mat_pedigree t " +
                                                " where out_mat_id = '{0}' ", matId);

                    try
                    {
                        OleDbCommand command = new OleDbCommand(sql, connection);


                        prevMatId = matId;

                        using (OleDbDataReader reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                matId = reader.GetString(0);

                                MaterialInfo mat = new MaterialInfo();
                                mat.Equipments = new BindingList<EquipmentInfo>();
                                mat.InId = reader["in_mat_id1"].ToString();
                                mat.OutId = reader["out_mat_id"].ToString();
                                mat.Process = reader["process_code"].ToString();

                                track.Add(mat);

                                string workshop;

                                sql = string.Format("select distinct workshop_no from equipment_code " +
                                    " where process_code = '{0}'", mat.Process);

                                command = new OleDbCommand(sql, connection);
                                using (OleDbDataReader read1 = command.ExecuteReader())
                                {
                                    if (read1.Read())
                                    {
                                        workshop = read1["workshop_no"].ToString();
                                    }
                                    else
                                    {
                                        throw new Exception("工序代码错误");
                                    }
                                }

                                string table;

                                if (tabledef.ContainsKey(workshop))
                                {
                                    table = tabledef[workshop];
                                }
                                else
                                {
                                    throw new Exception("工厂代码错误");
                                }

                                sql = string.Format("select mat_no, device_no, start_time, stop_time " +
                                    " from {1} " +
                                    " where mat_no = '{0}' and device_no in " +
                                    " (select device_no from equipment_code where process_code = '{2}')",
                                    mat.InId, table, mat.Process);

                                if (mat.Process == "CCM")
                                {
                                    sql = string.Format("select mat_no, device_no, start_time, stop_time " +
                                        " from {1} " +
                                        " where mat_no = '{0}' and device_no in " +
                                        " (select device_no from equipment_code where process_code = '{2}')",
                                        mat.OutId, table, mat.Process);
                                }

                                command = new OleDbCommand(sql, connection);

                                using (OleDbDataReader equ_reader = command.ExecuteReader())
                                {
                                    while (equ_reader.Read())
                                    {
                                        EquipmentInfo device = new EquipmentInfo();
                                        mat.Equipments.Add(device);

                                        device.MatId = matId;
                                        device.Name = equ_reader["device_no"].ToString();
                                        device.Workshop = workshop;
                                        device.StartTime = System.Convert.ToDateTime(equ_reader["start_time"]);
                                        device.StopTime = System.Convert.ToDateTime(equ_reader["stop_time"].ToString());

                                        sql = string.Format("select * from device_area_config where device_no = '{0}' order by display_num", device.Name);

                                        OleDbCommand cmd = new OleDbCommand(sql, connection);

                                        using (OleDbDataReader areaReader = cmd.ExecuteReader())
                                        {
                                            while (areaReader.Read())
                                            {
                                                EquipmentAreaInfo info = new EquipmentAreaInfo();

                                                info.Name = areaReader["device_no"].ToString();
                                                info.Workshop = areaReader["workshop_no"].ToString();
                                                info.Area = areaReader["area_no"].ToString();
                                                info.DisplaySeq = Convert.ToInt32(areaReader["display_num"].ToString());

                                                device.Areas.Add(info);
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    catch (OleDbException ex)
                    {
                        throw ex;
                    }
                }
            }

            return track;
        }
         */
        private void initSummary()
        {
            var ChooseNeedSummary = new DevExpress.XtraGrid.GridColumnSummaryItem(DevExpress.Data.SummaryItemType.Count, "", "{0}");
            this.gridView1.Columns[0].Summary.Add(ChooseNeedSummary);
        }
        private void barButtonItem3_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            string strName = "";
            try
            {
                if (this.gridView1.RowCount == 0)
                {
                    MessageBox.Show("Grid表格中没有数据，不能导出为Excel");
                    return;
                }
                DateTime MMSDate = DateTime.Now;
                using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                {
                    saveFileDialog.Filter = "Excel 工作簿(*.xlsx)|*.xlsx|Excel 97-2003 工作簿(*.xls)|*.xls|PDF(*.pdf)|*.pdf|Unicode 文本(*.txt)|*.txt";
                    saveFileDialog.FilterIndex = 0;
                    saveFileDialog.RestoreDirectory = true;
                    saveFileDialog.CreatePrompt = true;
                    saveFileDialog.Title = "导出文件保存路径";

                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        DevExpress.XtraPrinting.PrintingSystem ps = new DevExpress.XtraPrinting.PrintingSystem();
                        strName = saveFileDialog.FileName;
                        if (strName.LastIndexOf(".") - strName.LastIndexOf("\\") != 0)
                        {
                            switch (saveFileDialog.FilterIndex)
                            {
                                case 1:
                                    DevExpress.XtraPrinting.XlsxExportOptions xlsx = ps.ExportOptions.Xlsx;
                                    xlsx.ShowGridLines = true;
                                    xlsx.ExportMode = DevExpress.XtraPrinting.XlsxExportMode.SingleFile;
                                    this.gridControl1.ExportToXlsx(strName,xlsx); break;
                                case 2:
                                    DevExpress.XtraPrinting.XlsExportOptions xls = ps.ExportOptions.Xls;
                                    xls.ShowGridLines = true;
                                    this.gridView1.ExportToXls(strName); break;
                                case 3:
                                    this.gridView1.ExportToPdf(strName); break;
                                case 4:
                                    this.gridView1.ExportToText(strName); break;
                            }
                            MessageBox.Show("导出成功");
                        }
                        else
                        {
                            MessageBox.Show("保存的名称不能为空");
                        }

                    }
                }
            }
            catch (System.Exception msg)
            {
                MessageBox.Show(msg.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
        private void bindingNavigatorMoveLastItem_Click(object sender, EventArgs e)
        {
            if (currentpage < myData.Tables.Count - 1)
            {
                currentpage++;
                this.myBind.DataMember = myData.Tables[currentpage].TableName;
                this.gridView1.Columns.Clear();
                this.gridControl1.DataSource = null;
                this.gridControl1.DataSource = this.myBind;
            }
        }

        private void bindingNavigatorMoveFirstItem_Click(object sender, EventArgs e)
        {
            if (currentpage >0)
            {
                currentpage--;
                this.myBind.DataMember = myData.Tables[currentpage].TableName;
                this.gridView1.Columns.Clear();
                this.gridControl1.DataSource = null;
                this.gridControl1.DataSource = this.myBind;
            }
        }
    }
}
