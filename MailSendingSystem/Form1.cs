using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MailSendingSystem
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            EMailEvent += XXX;
            InitializeComponent();
        }


        private string path = "";
        private static DataTable dtResult;
        /// <summary>
        /// 第几行是列名
        /// </summary>
        private int firstRowColumn = 2;
        /// <summary>
        /// 第几列是邮箱地址
        /// </summary>
        private int mailColumn = 3;


        private List<string> fails = new List<string>();
        private int successNum = 0;
        


        private void panel1_Paint(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics, ClientRectangle, Color.Ivory, ButtonBorderStyle.Solid);
        }

        private void button_xuanzewenjian_Click(object sender, EventArgs e)
        {
            //首先根据打开文件对话框，选择要打开的文件
            OpenFileDialog ofd = new OpenFileDialog();
            //打开文件对话框筛选器，默认显示文件类型
            ofd.Filter = "Excel表格|*.xlsx|Excel97-2003表格|*.xls|所有文件|*.*";
            //定义文件路径
            string strPath;
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    strPath = ofd.FileName;
                    path = strPath;
                    dtResult = ReadExcelToDataTable(strPath, null, firstRowColumn);
                    var num = dtResult.Rows.Count;
                    label6.Text = $"已加载{num}条数据";
                }
                catch (Exception ex)
                {
                    MessageBox.Show("导入Excel失败：" + ex.Message);//捕捉异常
                }
            }
        }



        /// <summary>
        /// 将excel文件内容读取到DataTable数据表中
        /// </summary>
        /// <param name="fileName">文件完整路径名</param>
        /// <param name="sheetName">指定读取excel工作薄sheet的名称</param>
        /// <param name="firstRowColumn">第几行是DataTable的列名：从0开始</param>
        /// <returns>DataTable数据表</returns>
        public static DataTable ReadExcelToDataTable(string fileName, string sheetName = null, int firstRowColumn = 0)
        {
            //定义返回值
            DataTable dt = new DataTable();
            //定义WookBook
            IWorkbook workbook = null;
            //Sheet页
            ISheet sheet = null;
            //定义数据起始行
            int startDataRow = 0;
            try
            {
                //指定文件是否存在
                if (!File.Exists(fileName))
                {
                    return null;
                }
                //根据文件流创建excel数据结构
                using (FileStream fileStream = new FileStream(fileName, FileMode.Open, FileAccess.Read))
                {
                    workbook = WorkbookFactory.Create(fileStream);
                }
                //获取指定名称的Sheet页，若果未指定名称，则获取第一个sheet页
                if (string.IsNullOrEmpty(sheetName))
                {
                    sheet = workbook.GetSheetAt(0);
                }
                else
                {
                    sheet = workbook.GetSheet(sheetName);
                }
                //如果没有Sheet页，则返回
                if (sheet == null)
                {
                    return null;
                }
                IRow firstRow = sheet.GetRow(firstRowColumn);

                //添加标题
                for (int i = firstRow.FirstCellNum; i < firstRow.LastCellNum; i++)
                {
                    ICell cell = firstRow.GetCell(i);
                    if (cell == null)
                    {
                        continue;
                    }
                    var cellValue = cell.StringCellValue;
                    if (cellValue == null)
                    {
                        continue;
                    }
                    DataColumn column = new DataColumn(cellValue);
                    dt.Columns.Add(column);
                }
                startDataRow = firstRowColumn + 1;

                //添加数据
                for (int i = startDataRow; i <= sheet.LastRowNum; i++)
                {
                    IRow row = sheet.GetRow(i);
                    if (row == null || row.Count()<=0)
                    {
                        continue;
                    }
                    DataRow dataRow = dt.NewRow();
                    for (int j = row.FirstCellNum; j < firstRow.LastCellNum; ++j)
                    {
                        ICell cell = row.GetCell(j);
                        if (cell == null)
                        {
                            continue;
                        }

                        dataRow[j] = cell.ToString();

                        //if (DateTime.TryParse(cell.StringCellValue, out DateTime date))
                        //{
                        //    dataRow[j] = date;
                        //}
                        //else
                        //{
                        //    dataRow[j] = cell.ToString();
                        //}

                        //dataRow[j] = GetValueType(cell);
                    }
                    dt.Rows.Add(dataRow);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return dt;
        }

        /// <summary>
        /// 获取单元格类型
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        private static object GetValueType(ICell cell)
        {
            if (cell == null)
                return null;
            switch (cell.CellType)
            {
                case CellType.Blank: //BLANK:  
                    return "";
                case CellType.Boolean: //BOOLEAN:  
                    return cell.BooleanCellValue;
                case CellType.Numeric: //NUMERIC:  
                    short format = cell.CellStyle.DataFormat;
                    if (format != 0) { return cell.DateCellValue; } else { return cell.NumericCellValue; }
                case CellType.String: //STRING:  
                    return cell.StringCellValue;
                case CellType.Error: //ERROR:  
                    return cell.ErrorCellValue;
                case CellType.Formula: //FORMULA:  
                default:
                    return "=" + cell.CellFormula;
            }
        }


        /// <summary>
        /// 发送邮件
        /// </summary>
        private void BtnSend_Click(object sender, EventArgs e)
        {
            try
            {
                string userName = textBox_user.Text;
                string password = textBox_password.Text;
                string subject = textBox_title.Text;
                string body = textBox_body.Text;
                //if (string.IsNullOrEmpty(this.tbUserName.Text))
                //{
                //    MessageBox.Show("请输入发件人账号！");
                //    return;
                //}
                //if (string.IsNullOrEmpty(this.tbPassword.Text))
                //{
                //    MessageBox.Show("请输入发件人密码！");
                //    return;
                //}

                var tempPath = System.AppDomain.CurrentDomain.BaseDirectory + @"\Template.html";
                var tempStr = File.ReadAllText(tempPath,Encoding.UTF8);
                body = body.Replace("【表格】", tempStr);


                EmailHandle emailHandle = new EmailHandle();
                emailHandle.From = userName;//"chen.liu@ronds.com.cn";
                fails = new List<string>();
                successNum = 0;
                _menuUpdate(false);
                textBox5.Text = "开始发送...\r\n";
                Task.Factory.StartNew(() =>
                {
                    foreach (DataRow row in dtResult.Rows)
                    {
                        try
                        {
                            //拼接邮件内容
                            var emilStr = body;
                            int itemIndex = 0;
                            foreach (object item in row.ItemArray)
                            {
                                if (item != null)
                                    emilStr = emilStr.Replace($"【*{itemIndex.ToString()}*】", item.ToString());
                                itemIndex++;
                            }

                            var list = new List<string>();
                            list.Add(row[mailColumn]?.ToString());
                            emailHandle.To = list;
                            emailHandle.Subject = subject;
                            emailHandle.Body = emilStr;
                            emailHandle.IsBodyHtml = true;
                            emailHandle.UserName = userName;//"chen.liu@ronds.com.cn";
                            emailHandle.Password = password;
                            emailHandle.Host = "smtp.mxhichina.com";
                            emailHandle.Send();

                            successNum += 1;
                            EMailEvent.Invoke(1, null);
                        }
                        catch(Exception ex)
                        {
                            string err = $"{row[mailColumn]?.ToString()}发送失败，{ex.Message}";
                            fails.Add(err);
                            EMailEvent.Invoke(err, null);
                        }
                    }
                    EMailEvent.Invoke("已完成", new EventArgs());
                });
                
                //MessageBox.Show("全部发送成功!");
            }
            catch (Exception ex)
            {
                MessageBox.Show("发送失败:" + ex.Message.ToString());
            }
        }

        private void _menuUpdate(bool v)
        {
            textBox_user.Enabled = v;
            textBox_password.Enabled = v;
            textBox_title.Enabled = v;
            textBox_body.Enabled = v;
            button_xuanzewenjian.Enabled = v;
            BtnSend.Enabled = v;
        }

        #region 事件
        /// <summary>
        /// 用event关键字声明事件对象(如果EventHandler不够用，可先用delegate定义一个处理事件)
        /// </summary>
        public event EventHandler<EventArgs> EMailEvent;

        private void XXX(object sender, EventArgs e)
        {
            this.Invoke((EventHandler)delegate
            {
                if (e == null)
                {
                    if (sender.ToString() == "1")//成功
                    {
                        label7.Text = $"已发送：{successNum}";
                    }
                    else
                    {
                        textBox5.Text += sender.ToString() + "\r\n";
                    }
                }
                else
                {
                    textBox5.Text += "发送结束\r\n";
                    MessageBox.Show(this, $"邮件发送完毕，成功：{successNum}，失败{fails.Count}");
                    _menuUpdate(true);
                }
            });
        }
        #endregion

        private void pictureBox1_MouseMove(object sender, MouseEventArgs e)
        {
            Control control = sender as Control;
            control.BackColor = System.Drawing.SystemColors.ScrollBar;
        }

        private void pictureBox1_MouseLeave(object sender, EventArgs e)
        {
            Control control = sender as Control;
            control.BackColor = System.Drawing.SystemColors.Control;
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            if (textBox_password.UseSystemPasswordChar)
            {
                textBox_password.UseSystemPasswordChar = false;
            }
            else
            {
                textBox_password.UseSystemPasswordChar = true;
            }
        }
    }
}
