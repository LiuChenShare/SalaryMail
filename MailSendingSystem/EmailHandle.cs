using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;

namespace MailSendingSystem
{
    public class EmailHandle
    {
        private string _serviceType = "SMTP";
        private string _host;
        private object LogManager;

        /// <summary>
        /// 发送者邮箱
        /// </summary>
        public string From { get; set; }

        /// <summary>
        /// 接收者邮箱列表
        /// </summary>
        public List<string> To { get; set; }

        /// <summary>
        /// 抄送者邮箱列表
        /// </summary>
        public string[] Cc { get; set; }

        /// <summary>
        /// 秘抄者邮箱列表
        /// </summary>
        public string[] Bcc { get; set; }

        /// <summary>
        /// 邮件主题
        /// </summary>
        public string Subject { get; set; }

        /// <summary>
        /// 邮件内容
        /// </summary>
        public string Body { get; set; }

        /// <summary>
        /// 是否是HTML格式
        /// </summary>
        public bool IsBodyHtml { get; set; }

        /// <summary>
        /// 附件列表
        /// </summary>
        public string[] Attachments { get; set; }

        /// <summary>
        /// 邮箱服务类型(Pop3,SMTP,IMAP,MAIL等)，默认为SMTP
        /// </summary>
        public string ServiceType
        {
            get { return _serviceType; }
            set { _serviceType = value; }
        }

        /// <summary>
        /// 邮箱服务器，如果没有定义邮箱服务器，则根据serviceType和Sender组成邮箱服务器
        /// </summary>
        public string Host
        {
            get { return _host; }
            set { _host = value; }
        }

        /// <summary>
        /// 邮箱账号(默认为发送者邮箱的账号)
        /// </summary>
        public string UserName { get; set; }

        /// <summary>
        /// 邮箱密码(默认为发送者邮箱的密码)，默认格式GB2312
        /// </summary>
        public string Password { get; set; }

        /// <summary>
        /// 邮箱优先级
        /// </summary>
        public MailPriority MailPriority { get; set; }

        /// <summary>
        ///  邮件正文编码格式
        /// </summary>
        public Encoding Encoding { get; set; }

        /// <summary>
        /// 构造参数，发送邮件，使用方法备注：公开方法中调用
        /// </summary>
        public int Send()
        {
            var mailMessage = new MailMessage();
            //读取To  接收者邮箱列表
            try
            {
                if (this.To != null && this.To.Count > 0)
                {
                    foreach (string to in this.To)
                    {
                        if (string.IsNullOrEmpty(to)) continue;
                        mailMessage.To.Add(new MailAddress(to.Trim()));
                    }
                }
                //读取Cc  抄送者邮件地址
                if (this.Cc != null && this.Cc.Length > 0)
                {
                    foreach (var cc in this.Cc)
                    {
                        if (string.IsNullOrEmpty(cc)) continue;
                        mailMessage.CC.Add(new MailAddress(cc.Trim()));
                    }
                }
                //读取Attachments 邮件附件
                if (this.Attachments != null && this.Attachments.Length > 0)
                {
                    foreach (var attachment in this.Attachments)
                    {
                        if (string.IsNullOrEmpty(attachment)) continue;
                        mailMessage.Attachments.Add(new Attachment(attachment));
                    }
                }
                //读取Bcc 秘抄人地址
                if (this.Bcc != null && this.Bcc.Length > 0)
                {
                    foreach (var bcc in this.Bcc)
                    {
                        if (string.IsNullOrEmpty(bcc)) continue;
                        mailMessage.Bcc.Add(new MailAddress(bcc.Trim()));
                    }
                }
                //读取From 发送人地址
                mailMessage.From = new MailAddress(this.From);
                //邮件标题
                Encoding encoding = Encoding.GetEncoding("GB2312");
                mailMessage.Subject = this.Subject;
                //邮件正文是否为HTML格式
                mailMessage.IsBodyHtml = this.IsBodyHtml;
                //邮件正文
                mailMessage.Body = this.Body;
                mailMessage.BodyEncoding = this.Encoding;
                //邮件优先级
                mailMessage.Priority = this.MailPriority;
                //发送邮件代码实现
                var smtpClient = new SmtpClient
                {
                    Host = this.Host,
                    EnableSsl = true,
                    Credentials = new NetworkCredential(this.UserName, this.Password)
                };
                //加这段之前用公司邮箱发送报错：根据验证过程，远程证书无效
                //加上后解决问题
                ServicePointManager.ServerCertificateValidationCallback = delegate (Object obj, X509Certificate certificate, X509Chain chain, SslPolicyErrors errors) { return true; };
                //认证
                smtpClient.Send(mailMessage);
                return 1;
            }
            catch (Exception ex)
            {
                //return -1;
                throw ex;
            }
        }
    }

}
