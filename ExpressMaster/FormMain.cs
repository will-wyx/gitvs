using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Windows.Forms;
using System.Web.Script.Serialization;
using System.Text;

namespace ExpressMaster
{
    public partial class FormMain : Form
    {
        public FormMain()
        {
            InitializeComponent();
        }

        private string sourceFolderName = "Source纸质单";
        private string resultFolderName = "Result纸质单";
        private string type = "";
        private string company = "";
        private ExcelHelper exhlp = new ExcelHelper();

        private void FormMain_Load(object sender, EventArgs e)
        {
            /* 加载配置 */
            LoadSettings();
        }

        List<Data4Cfg> data = new List<Data4Cfg>();

        JavaScriptSerializer jss = new JavaScriptSerializer();
        string profileName = "profile.json";

        /// <summary>
        /// 排序
        /// </summary>
        /// <param name="p1"></param>
        /// <param name="p2"></param>
        /// <returns></returns>
        private static int SortProfileEntity(ProfileEntity p1, ProfileEntity p2)
        {
            return (p1.PinyinInitials == p2.PinyinInitials ? 0 : (p1.PinyinInitials > p2.PinyinInitials ? -1 : 1));
        }

        List<ProfileEntity> profileEntitys /* 配置项 */;
        /// <summary>
        /// 加载配置
        /// </summary>
        private void LoadSettings()
        {
            tsslInfo.Text = "Loading...";
            /* ready */
            tsslInfo.Text = Properties.Settings.Default.LoadState;
            /* export */

            /* load profile */
            FileInfo profile = new FileInfo(profileName);

            if (!profile.Exists)
            {
                /* 文件不存在 */
                profileEntitys = new List<ProfileEntity>();

                string key1 = "default", key2 = "北京|天津|河北";

                profileEntitys.Add(new ProfileEntity
                {
                    Name = "红牡丹",
                    PinyinInitials = 'H',
                    Values = new ValuesEntity[]
                    {
                        new ValuesEntity
                        {
                            Items = new Data4Cfg[]
                            {
                                new Data4Cfg { Key = key1, FirstWeight = 1,FirstAmount = 1, OtherAmount = 1 },
                                new Data4Cfg { Key = key2, FirstWeight = 1,FirstAmount = 1, OtherAmount = 1 }
                            }
                        }
                    }
                });
                #region Profiledata

                profileEntitys.Add(new ProfileEntity
                {
                    Name = "小孟",
                    PinyinInitials = 'X',
                    Values = new ValuesEntity[]
                    {
                        new ValuesEntity
                        {
                            Items = new Data4Cfg[]
                            {
                                new Data4Cfg { Key = key1, FirstWeight = 1,FirstAmount = 1, OtherAmount = 1 },
                                new Data4Cfg { Key = key2, FirstWeight = 1,FirstAmount = 1, OtherAmount = 1 }
                            }
                        }
                    }
                });

                profileEntitys.Add(new ProfileEntity
                {
                    Name = "芊芊家",
                    PinyinInitials = 'Q',
                    Values = new ValuesEntity[]
                    {
                        new ValuesEntity
                        {
                            Items = new Data4Cfg[]
                            {
                                new Data4Cfg { Key = key1, FirstWeight = 1,FirstAmount = 1, OtherAmount = 1 },
                                new Data4Cfg { Key = key2, FirstWeight = 1,FirstAmount = 2, OtherAmount = 1 }
                            }
                        }
                    }
                });

                profileEntitys.Add(new ProfileEntity
                {
                    Name = "宇旺",
                    PinyinInitials = 'Y',
                    Values = new ValuesEntity[]
                    {
                        new ValuesEntity
                        {
                            Items = new Data4Cfg[]
                            {
                                new Data4Cfg { Key = key1, FirstWeight = 1,FirstAmount = 1, OtherAmount = 1 },
                                new Data4Cfg { Key = key2, FirstWeight = 1,FirstAmount = 2, OtherAmount = 1 }
                            }
                        }
                    }
                });

                profileEntitys.Add(new ProfileEntity
                {
                    Name = "利忠",
                    PinyinInitials = 'L',
                    Values = new ValuesEntity[]
                    {
                        new ValuesEntity
                        {
                            Items = new Data4Cfg[]
                            {
                                new Data4Cfg { Key = key1, FirstWeight = 1,FirstAmount = 1, OtherAmount = 1 },
                                new Data4Cfg { Key = key2, FirstWeight = 1,FirstAmount = 2, OtherAmount = 1 }
                            }
                        }
                    }
                });

                profileEntitys.Add(new ProfileEntity
                {
                    Name = "智远",
                    PinyinInitials = 'Z',
                    Values = new ValuesEntity[]
                    {
                        new ValuesEntity
                        {
                            Items = new Data4Cfg[]
                            {
                                new Data4Cfg { Key = key1, FirstWeight = 1,FirstAmount = 2, OtherAmount = 1 },
                                new Data4Cfg { Key = key2, FirstWeight = 1,FirstAmount = 2, OtherAmount = 1 }
                            }
                        }
                    }
                });

                profileEntitys.Add(new ProfileEntity
                {
                    Name = "美姿",
                    PinyinInitials = 'M',
                    Values = new ValuesEntity[]
                    {
                        new ValuesEntity
                        {
                            Items = new Data4Cfg[]
                            {
                                new Data4Cfg { Key = key1, FirstWeight = 1,FirstAmount = 2, OtherAmount = 1 },
                                new Data4Cfg { Key = key2, FirstWeight = 1,FirstAmount = 2, OtherAmount = 1 }
                            }
                        }
                    }
                });

                profileEntitys.Add(new ProfileEntity
                {
                    Name = "刘磊",
                    PinyinInitials = 'L',
                    Values = new ValuesEntity[]
                    {
                        new ValuesEntity
                        {
                            Items = new Data4Cfg[]
                            {
                                new Data4Cfg { Key = key1, FirstWeight = 1,FirstAmount = 2, OtherAmount = 1 },
                                new Data4Cfg { Key = key2, FirstWeight = 1,FirstAmount = 2, OtherAmount = 1 }
                            }
                        }
                    }
                });

                profileEntitys.Add(new ProfileEntity
                {
                    Name = "正新（途乐）",
                    PinyinInitials = 'Z',
                    Values = new ValuesEntity[]
                    {
                        new ValuesEntity
                        {
                            Items = new Data4Cfg[]
                            {
                                new Data4Cfg { Key = key1, FirstWeight = 1,FirstAmount = 2, OtherAmount = 1 },
                                new Data4Cfg { Key = key2, FirstWeight = 1,FirstAmount = 2, OtherAmount = 1 }
                            }
                        }
                    }
                });

                profileEntitys.Add(new ProfileEntity
                {
                    Name = "马家沟",
                    PinyinInitials = 'M',
                    Values = new ValuesEntity[]
                    {
                        new ValuesEntity
                        {
                            Items = new Data4Cfg[]
                            {
                                new Data4Cfg { Key = key1, FirstWeight = 1,FirstAmount = 2, OtherAmount = 1 },
                                new Data4Cfg { Key = key2, FirstWeight = 1,FirstAmount = 2, OtherAmount = 1 }
                            }
                        }
                    }
                });

                profileEntitys.Add(new ProfileEntity
                {
                    Name = "龙邦",
                    PinyinInitials = 'L',
                    Values = new ValuesEntity[]
                    {
                        new ValuesEntity
                        {
                            Items = new Data4Cfg[]
                            {
                                new Data4Cfg { Key = key1, FirstWeight = 1,FirstAmount = 2, OtherAmount = 1 },
                                new Data4Cfg { Key = key2, FirstWeight = 1,FirstAmount = 2, OtherAmount = 1 }
                            }
                        }
                    }
                });
                #endregion

                string json = jss.Serialize(profileEntitys);
                FileStream fs = profile.Create();
                byte[] buffer = Encoding.UTF8.GetBytes(json);
                fs.Write(buffer, 0, buffer.Length);
                fs.Close();
            }
            else
            {
                /* 存在配置文件 */
                FileStream fs = profile.OpenRead();
                int fileLength = (int)fs.Length;
                byte[] buffer = new byte[fileLength];
                fs.Read(buffer, 0, fileLength);
                fs.Close();
                string json = Encoding.UTF8.GetString(buffer);
                profileEntitys = jss.Deserialize<List<ProfileEntity>>(json);
            }

            //profileEntitys.Sort(SortProfileEntity);

            //System.Collections.Specialized.StringCollection
            //    formulas1 = Properties.Settings.Default.Formula1,
            //    formulas2 = Properties.Settings.Default.Formula2;

            //foreach (string item in formulas1)
            //{
            //    Data4Cfg d4c = new Data4Cfg();
            //    string[] itemArr = item.Split(',');
            //    d4c.Key = itemArr[0];
            //    d4c.FirstWeight = Convert.ToDecimal(itemArr[1]);
            //    d4c.FirstAmount = Convert.ToDecimal(itemArr[2]);
            //    d4c.OtherAmount = Convert.ToDecimal(itemArr[3]);
            //    data1.Add(d4c);
            //}

            //foreach (string item in formulas2)
            //{
            //    Data4Cfg d4c = new Data4Cfg();
            //    string[] itemArr = item.Split(',');
            //    d4c.Key = itemArr[0];
            //    d4c.FirstWeight = Convert.ToDecimal(itemArr[1]);
            //    d4c.FirstAmount = Convert.ToDecimal(itemArr[2]);
            //    d4c.OtherAmount = Convert.ToDecimal(itemArr[3]);
            //    data2.Add(d4c);
            //}

            //dgvMain.DataSource = data;
            /* load combobox */
            CreateCbxMainDataSource();
            CreateCbxMinorDataSource();

            //BindingDGV();
            tsslInfo.Text = "Completed";

        }

        /// <summary>
        /// 绑定数据
        /// </summary>
        private void BindingDGV()
        {

            data.Clear();
            ProfileEntity pe = profileEntitys.Find(p => p.Name.Equals(company));
            if (pe != null)
            {
                foreach (ValuesEntity ve in pe.Values)
                {
                    data.AddRange(ve.Items);
                }
                dgvMain.DataSource = new List<Data4Cfg>();
                dgvMain.DataSource = data;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        private void CreateCbxMainDataSource()
        {
            DataTable dt = new DataTable("t1");
            DataColumn valueCol = new DataColumn("value", System.Type.GetType("System.String"));
            dt.Columns.Add(valueCol);
            DataColumn textCol = new DataColumn("text", System.Type.GetType("System.String"));
            dt.Columns.Add(textCol);

            DataRow dr1 = dt.NewRow();
            dr1["value"] = "纸质单";
            dr1["text"] = "纸质单";
            DataRow dr2 = dt.NewRow();
            dr2["value"] = "电子单";
            dr2["text"] = "电子单";
            DataRow dr3 = dt.NewRow();
            dr3["value"] = "菜鸟";
            dr3["text"] = "菜鸟";
            dt.Rows.Add(dr1);
            dt.Rows.Add(dr2);
            dt.Rows.Add(dr3);
            cbxMain.DataSource = dt;
            cbxMain.DisplayMember = "text";
            cbxMain.ValueMember = "value";
            cbxMain_SelectedIndexChanged(this, null);
        }

        /// <summary>
        /// 
        /// </summary>
        private void CreateCbxMinorDataSource()
        {
            DataTable dt = new DataTable("t2");
            DataColumn valueCol = new DataColumn("value", System.Type.GetType("System.String"));
            dt.Columns.Add(valueCol);
            DataColumn textCol = new DataColumn("text", System.Type.GetType("System.String"));
            dt.Columns.Add(textCol);
            foreach (ProfileEntity pe in profileEntitys)
            {
                DataRow dr = dt.NewRow();
                dr["value"] = pe.Name;
                dr["text"] = pe.Name;
                dt.Rows.Add(dr);
            }
            cbxMinor.DataSource = dt;
            cbxMinor.DisplayMember = "text";
            cbxMinor.ValueMember = "value";
            cbxMinor_SelectedIndexChanged(this, null);
        }

        /// <summary>
        /// 导出
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnExport_Click(object sender, EventArgs e)
        {
            exhlp.Data = data;

            ProfileEntity pe = profileEntitys.Find(p => p.Name.Equals(company));
            foreach (ValuesEntity ve in pe.Values)
            {
                ve.Items = data.ToArray();
            }
            FileInfo profile = new FileInfo(profileName);
            string json = jss.Serialize(profileEntitys);
            FileStream fs = profile.Create();
            byte[] buffer = Encoding.UTF8.GetBytes(json);
            fs.Write(buffer, 0, buffer.Length);
            fs.Close();
            Process();
        }

        /// <summary>
        /// 处理并导出
        /// </summary>
        private void Process()
        {
            /* 所有公式 - data */
            /* 取得所有文件名 */
            DirectoryInfo dirSource = new DirectoryInfo(sourceFolderName);
            DirectoryInfo dirResult = new DirectoryInfo(resultFolderName);
            if (!dirResult.Exists)
            {
                Directory.CreateDirectory(resultFolderName);
            }
            if (dirSource.Exists)
            {
                FileInfo[] files = dirSource.GetFiles();
                ProcessFileHandler handler = new ProcessFileHandler(ProcessFile);
                List<IAsyncResult> results = new List<IAsyncResult>();
                tsslInfo.Text = "正在处理...";
                bool typeFlag = type.Equals("纸质单");
                if (typeFlag)
                {
                    //exhlp.list.Clear();
                    /* 加载单号 */
                    DirectoryInfo dirData = new DirectoryInfo("Data单号");
                    if (dirData.Exists)
                    {
                        FileInfo[] dataFiles = dirData.GetFiles();
                        exhlp.DataOrder.Clear();
                        exhlp.DataOrderUnBindCount = 0;
                        foreach (var file in dataFiles)
                        {
                            List<ExpressEntity> lee = exhlp.LoadOrderNumber(file);
                        }
                    }
                    else
                    {
                        MessageBox.Show("没有单号文件夹");
                    }
                }
                foreach (var file in files)
                {
                    IAsyncResult result = handler.BeginInvoke(file, null, null);
                    results.Add(result);
                }
                foreach (var result in results)
                {
                    bool flag = handler.EndInvoke(result);
                }

                if (typeFlag && exhlp.DataOrder.Count > 0)
                {
                    string toPath = resultFolderName + "/" + exhlp.filename;
                    FileStream toFile = File.Create(toPath);
                    exhlp.SaveExcelC(toFile);
                }
            }
            tsslInfo.Text = "完成";
            MessageBox.Show("completed");
        }

        public delegate bool ProcessFileHandler(FileInfo file);

        /// <summary>
        /// 处理文件
        /// </summary>
        /// <param name="file"></param>
        /// <returns></returns>
        private bool ProcessFile(FileInfo file)
        {
            /* 单类型 */
            FileStream fromFile = file.OpenRead();

            string toPath = "";
            FileStream toFile = null;
            switch (type)
            {
                case "纸质单":
                    exhlp.ProcessExcelTemplateC(fromFile);
                    break;
                case "电子单":
                    toPath = resultFolderName + "/" + file.Name;
                    toFile = File.Create(toPath);
                    exhlp.ProcessExcelTemplateB(fromFile, toFile);
                    toFile.Close();
                    break;
                case "菜鸟":
                    toPath = resultFolderName + "/" + file.Name;
                    toFile = File.Create(toPath);
                    exhlp.ProcessExcelTemplateA(fromFile, toFile);
                    toFile.Close();
                    break;
            }
            fromFile.Close();
            return true;
        }

        private void btnAddRow_Click(object sender, EventArgs e)
        {
            data.Add(new Data4Cfg
            {
                Key = "",
                FirstWeight = 0,
                FirstAmount = 0,
                OtherAmount = 0
            });
            dgvMain.DataSource = new List<Data4Cfg>();
            dgvMain.DataSource = data;
        }

        private void tsmiRemove_Click(object sender, EventArgs e)
        {
            var rows = dgvMain.SelectedRows;
            if (rows.Count > 0)
            {
                DataGridViewRow row = rows[0];
                data.Remove((Data4Cfg)row.DataBoundItem);
                dgvMain.DataSource = new List<Data4Cfg>();
                dgvMain.DataSource = data;
            }
        }

        private void cmsMain_Opening(object sender, CancelEventArgs e)
        {

        }

        private void cbxMain_SelectedIndexChanged(object sender, EventArgs e)
        {
            sourceFolderName = "Source" + cbxMain.SelectedValue.ToString();
            resultFolderName = "Result" + cbxMain.SelectedValue.ToString();
            type = cbxMain.SelectedValue.ToString();
            BindingDGV();
        }

        private void cbxMinor_SelectedIndexChanged(object sender, EventArgs e)
        {
            company = cbxMinor.SelectedValue.ToString();
            BindingDGV();
        }
    }
}
