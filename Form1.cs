using SMT.Foundation.Log;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml;

namespace 会员管理
{
    public partial class Form1 : Form
    {
        private static string AdminToken = string.Empty;
        private string strPostDate = string.Empty;
        private string cookie_cfduid = string.Empty;
        private string cookie_AspNetSessionId = string.Empty;
        private string cookie_SessionCode = string.Empty;


        private string cookie_SessionUser = string.Empty;

        /// <summary>
        /// 每次登录会变化
        /// </summary>
        private string cookie_ASPSESSIONIDName = string.Empty;
        private string cookie_ASPSESSIONIDValue = string.Empty;
        private string cookie_month_champion = string.Empty;

        private string CustomerMoblie = string.Empty;

        // __cfduid=d2003215804a6ba74df96de4bad4eb6251480646889
        //ASP.NET_SessionId=g121cj1fgk2wwed2z4vnegz0
        //sessionCode_=sessionCode_8de96a83-20e1-4852-a7f0-08f82e4fb852

        public Form1()
        {
            InitializeComponent();
            try
            {
                setCookie1();
                setCookie2();
                timer1.Interval = 3000;//3s
                
            }
            catch (Exception ex)
            {
                log(ex.ToString());
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            button1.Enabled = false;
            try
            {
                login();
                setCookie4();
                GetCustomerList();
                getSaleListCookie();
            }catch(Exception ex)
            {
                log(ex.ToString());
            }
            finally
            {
                button1.Enabled = true;
            }
        }

        //第一步通过http://portal.idcicp.com/login.html 设置cookie __cfduid
        private void setCookie1()
        {
            string url = "http://portal.idcicp.com/login.html";

            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
            HttpWebResponse response = (HttpWebResponse)request.GetResponse();


            string CookieValue = response.Headers.Get("Set-Cookie");
            string[] all = CookieValue.Split(';');
            cookie_cfduid = all[0].Split('=')[1];
        }

        /// <summary>
        /// 第二步：通过http://portal.idcicp.com/tools/Verify_code.ashx 设置: ASP.NET_SessionId sessionCode
        /// </summary>
        private void setCookie2()
        {
            //获取ASP.NET_SessionId sessionCode
            string url = "http://portal.idcicp.com/tools/Verify_code.ashx";

            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
            HttpWebResponse response = (HttpWebResponse)request.GetResponse();
            Stream responseStream = response.GetResponseStream();
            Bitmap btimg = new Bitmap(responseStream);
            pictureBox1.Image = btimg;

            StreamReader reader = new StreamReader(responseStream, Encoding.UTF8);
            string content = reader.ReadToEnd();
            reader.Close();
            responseStream.Close();
            string CookieValue = response.Headers.Get("Set-Cookie");

            string[] all = CookieValue.Split(';');
            foreach (string s in all)
            {
                if (s.Contains("SessionId"))
                {
                    cookie_AspNetSessionId = s.Split(',')[1].Split('=')[1];
                }
                if (s.Contains("sessionCode"))
                {
                    //HttpOnly,sessionCode_=sessionCode_ca1d0a12-5619-40ff-8986-57471d611d29; 
                    cookie_SessionCode = s.Split(',')[1].Split('=')[1];
                }

            }
        }


        /// <summary>
        /// 第三步 登录获取cooike_sesssionUser
        /// </summary>
        /// <returns></returns>
        public string login()
        {
            string strLogin = "http://portal.idcicp.com/login_center.aspx?action=login";
            strPostDate = "ref=&loginname=zwd312146008&loginpass=36916266&twopass=idcicp123topdatazwd312146008&logincode=" + textBox1.Text;
            if (!string.IsNullOrEmpty(AdminToken))
            {
                return AdminToken;
            }
            else
            {
                string strResult = "";

                try
                {
                    //myRequest.CookieContainer.Add()

                    HttpWebRequest myRequest = (HttpWebRequest)WebRequest.Create(strLogin);
                    myRequest.Method = "POST";
                    myRequest.ContentType = "application/x-www-form-urlencoded";
                    myRequest.AllowAutoRedirect = false;
                    Cookie ck1 = new Cookie("__cfduid", cookie_cfduid, "/", ".idcicp.com");
                    Cookie ck2 = new Cookie("ASP.NET_SessionId", cookie_AspNetSessionId, "/", ".idcicp.com");
                    Cookie ck3 = new Cookie("sessionCode_", cookie_SessionCode, "/", ".idcicp.com");

                    CookieCollection cks = new CookieCollection();
                    cks.Add(ck1);
                    cks.Add(ck2);
                    cks.Add(ck3);
                    myRequest.CookieContainer = new CookieContainer();
                    myRequest.CookieContainer.Add(cks);

                    string requestBody = strPostDate;

                    byte[] aryBuf = Encoding.UTF8.GetBytes(requestBody);
                    myRequest.ContentLength = aryBuf.Length;
                    using (Stream writer = myRequest.GetRequestStream())
                    {
                        writer.Write(aryBuf, 0, aryBuf.Length);
                        writer.Close();
                        writer.Dispose();
                    }

                    HttpWebResponse response = (HttpWebResponse)myRequest.GetResponse();
                    Stream responseStream = response.GetResponseStream();
                    StreamReader reader = new StreamReader(responseStream, Encoding.UTF8);
                    string content = reader.ReadToEnd();
                    reader.Close();
                    responseStream.Close();



                    string CookieValue = response.Headers.Get("Set-Cookie");

                    string[] all = CookieValue.Split(';');
                    foreach (string s in all)
                    {
                        if (s.Contains("sessionUser_"))
                        {
                            cookie_SessionUser = s.Split('=')[1];
                        }

                    }
                }
                catch (Exception ex)
                {
                    log("环信GetAdminToken异常：" + ex.ToString());
                }
                return strResult;
            }

        }

        //第四步，获取cookie_ASPSESSIONIDASRQTDSA
        private void setCookie4()
        {
            string url = "http://portal.idcicp.com/depart_login.asp?action=login&loginname=zwd312146008&loginpass=538b4badb72d48b215437e07b2502f78&ref=";

            HttpWebRequest myRequest = (HttpWebRequest)WebRequest.Create(url);
            myRequest.Method = "GET";
            myRequest.ContentType = "application/x-www-form-urlencoded";
            myRequest.AllowAutoRedirect = false;
            Cookie ck1 = new Cookie("__cfduid", cookie_cfduid, "/", ".idcicp.com");
            Cookie ck2 = new Cookie("ASP.NET_SessionId", cookie_AspNetSessionId, "/", ".idcicp.com");
            Cookie ck3 = new Cookie("sessionCode_", cookie_SessionCode, "/", ".idcicp.com");
            Cookie ck4 = new Cookie("sessionUser_", cookie_SessionUser, "/", ".idcicp.com");

            CookieCollection cks = new CookieCollection();
            cks.Add(ck1);
            cks.Add(ck2);
            cks.Add(ck3);
            cks.Add(ck4);
            myRequest.CookieContainer = new CookieContainer();
            myRequest.CookieContainer.Add(cks);


            HttpWebResponse response = (HttpWebResponse)myRequest.GetResponse();
            Stream responseStream = response.GetResponseStream();
            StreamReader reader = new StreamReader(responseStream, Encoding.UTF8);
            string content = reader.ReadToEnd();
            reader.Close();
            responseStream.Close();



            string CookieValue = response.Headers.Get("Set-Cookie");

            string[] all = CookieValue.Split(';');
            foreach (string s in all)
            {
                if (s.Contains("ASPSESSION"))
                {
                    cookie_ASPSESSIONIDName = s.Split('=')[0];
                    cookie_ASPSESSIONIDValue = s.Split('=')[1];
                }

            }

        }

        //第五步，获取客户列表信息
        private void GetCustomerList()
        {
            try
            {
                string url = "http://portal.idcicp.com/tophost_mg/userlist.asp?act=list";

                HttpWebRequest myRequest = (HttpWebRequest)WebRequest.Create(url);
                myRequest.Method = "GET";
                myRequest.ContentType = "application/x-www-form-urlencoded";
                myRequest.AllowAutoRedirect = false;
                Cookie ck1 = new Cookie("__cfduid", cookie_cfduid, "/", ".idcicp.com");
                Cookie ck2 = new Cookie("ASP.NET_SessionId", cookie_AspNetSessionId, "/", ".idcicp.com");
                Cookie ck3 = new Cookie("sessionCode_", cookie_SessionCode, "/", ".idcicp.com");
                Cookie ck4 = new Cookie("sessionUser_", cookie_SessionUser, "/", ".idcicp.com");
                Cookie ck5 = new Cookie(cookie_ASPSESSIONIDName, cookie_ASPSESSIONIDValue, "/", ".idcicp.com");

                CookieCollection cks = new CookieCollection();
                cks.Add(ck1);
                cks.Add(ck2);
                cks.Add(ck3);
                cks.Add(ck4);
                cks.Add(ck5);
                myRequest.CookieContainer = new CookieContainer();
                myRequest.CookieContainer.Add(cks);

                HttpWebResponse response = (HttpWebResponse)myRequest.GetResponse();
                Stream responseStream = response.GetResponseStream();
                StreamReader reader = new StreamReader(responseStream, Encoding.GetEncoding("gb2312"));
                string content = reader.ReadToEnd();
                reader.Close();
                responseStream.Close();

                webBrowser1.DocumentText = content;

                string trPattern = @"<tr>[\s\S]*?</tr>";
                //第一个tr是搜索，第二个tr才是客户信息
                Match trmatch = Regex.Match(content, trPattern).NextMatch();
                string tdPattern = @"<td[\s\S]*?</td>";
                string tdall = trmatch.Value;

                //重置了ASPSESSIONIDASQTRDTB cookie
                if(string.IsNullOrEmpty(tdall))
                {
                    string CookieValue = response.Headers.Get("Set-Cookie");
                    string[] all = CookieValue.Split(';');
                    foreach (string s in all)
                    {
                        if (s.Contains("ASPSESSION"))
                        {
                            cookie_ASPSESSIONIDName = s.Split('=')[0];
                            cookie_ASPSESSIONIDValue = s.Split('=')[1];
                        }

                    }
                    return;
                }
                MatchCollection tdmatch = Regex.Matches(tdall, tdPattern);

                Match MobileMatch = Regex.Match(tdmatch[5].Value, @"(?is)(?<=<td[^>]*?>).*?(?=</td>)");
                string strMobile = MobileMatch.Value.Replace("&nbsp;", "");

                Match kefuMatch = Regex.Match(tdmatch[8].Value, @"(?is)(?<=<td[^>]*?>).*?(?=</td>)");
                string strKefu = kefuMatch.Value.Replace("&nbsp;", "");

                //strKefu = txtMoblie.Text;

                if (string.IsNullOrEmpty(strKefu))//没人认领的新用户，马上启动录入
                {
                    log("发现新客户：" + strMobile + "开始添加客户");
                    CustomerMoblie = strMobile;
                    try
                    {
                        AddCustomer();
                        log("添加新客户：" + strMobile + "成功！");
                    }
                    catch (Exception ex)
                    {
                        log("添加新客户失败：" + strMobile + " 错误信息：" + ex.ToString());
                    }
                    finally
                    {
                        CustomerMoblie = string.Empty;
                    }
                }
            }catch(Exception ex)
            {
                log(ex.ToString());
            }
            finally
            {
                string str= "已获取到最新用户列表信息";
                log(str);
                timer1.Start();
            }
        }

        private void log(string str)
        {
            if(txtMsg.Text.Length>=100)
            {
                txtMsg.Text = string.Empty;
            }
            txtMsg.Text +=DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")+"  "+ str + System.Environment.NewLine;
            Tracer.Debug(str);
        }



        /// <summary>
        /// 查询新增需要用到此cookie
        /// </summary>
        private void getSaleListCookie()
        {
            string strUrl = "http://portal.idcicp.com/customer/list.aspx?act=mine";

            HttpWebRequest myRequest = (HttpWebRequest)WebRequest.Create(strUrl);
            myRequest.Method = "GET";
            myRequest.ContentType = "application/x-www-form-urlencoded";
            myRequest.AllowAutoRedirect = false;
            Cookie ck1 = new Cookie("__cfduid", cookie_cfduid, "/", ".idcicp.com");
            Cookie ck2 = new Cookie("ASP.NET_SessionId", cookie_AspNetSessionId, "/", ".idcicp.com");
            Cookie ck3 = new Cookie("sessionCode_", cookie_SessionCode, "/", ".idcicp.com");
            Cookie ck4 = new Cookie("sessionUser_", cookie_SessionUser, "/", ".idcicp.com");
            Cookie ck5 = new Cookie(cookie_ASPSESSIONIDName, cookie_ASPSESSIONIDValue, "/", ".idcicp.com");

            CookieCollection cks = new CookieCollection();
            cks.Add(ck1);
            cks.Add(ck2);
            cks.Add(ck3);
            cks.Add(ck4);
            cks.Add(ck5);
            myRequest.CookieContainer = new CookieContainer();
            myRequest.CookieContainer.Add(cks);


            HttpWebResponse response = (HttpWebResponse)myRequest.GetResponse();
            Stream responseStream = response.GetResponseStream();
            StreamReader reader = new StreamReader(responseStream, Encoding.UTF8);
            string content = reader.ReadToEnd();
            reader.Close();
            responseStream.Close();



            string CookieValue = response.Headers.Get("Set-Cookie");

            string[] all = CookieValue.Split(';');
            foreach (string s in all)
            {
                if (s.Contains("month_champion"))
                {
                    cookie_month_champion = s.Split('=')[1];
                }

            }

        }

        //第六步，添加客户列表信息
        private void AddCustomer()
        {
            string url = "http://portal.idcicp.com/customer/add.aspx";

            HttpWebRequest myRequest = (HttpWebRequest)WebRequest.Create(url);
            myRequest.Method = "POST";
            myRequest.ContentType = "application/x-www-form-urlencoded";
            myRequest.AllowAutoRedirect = false;
            Cookie ck1 = new Cookie("__cfduid", cookie_cfduid, "/", ".idcicp.com");
            Cookie ck2 = new Cookie("ASP.NET_SessionId", cookie_AspNetSessionId, "/", ".idcicp.com");
            Cookie ck3 = new Cookie("sessionCode_", cookie_SessionCode, "/", ".idcicp.com");
            Cookie ck4 = new Cookie("sessionUser_", cookie_SessionUser, "/", ".idcicp.com");
            Cookie ck5 = new Cookie(cookie_ASPSESSIONIDName, cookie_ASPSESSIONIDValue, "/", ".idcicp.com");
            if (string.IsNullOrEmpty(cookie_month_champion))
            {
                getSaleListCookie();
            }
            Cookie ck6 = new Cookie("month_champion", cookie_month_champion, "/", ".idcicp.com");

            CookieCollection cks = new CookieCollection();
            cks.Add(ck1);
            cks.Add(ck2);
            cks.Add(ck3);
            cks.Add(ck4);
            cks.Add(ck5);
            cks.Add(ck6);
            myRequest.CookieContainer = new CookieContainer();
            myRequest.CookieContainer.Add(cks);



            string requestBody = "custname=a&custcontactname=a&custorderfor=%E5%9F%9F%E5%90%8D%E4%BA%A7%E5%93%81&type2=%E5%9F%9F%E5%90%8D%E4%BA%A7%E5%93%81&sel_CustT_Name=0&custmembername=&custmoblie=" + CustomerMoblie + "&custtel=&custqq=&custwechat=&custemail="; ;

            byte[] aryBuf = Encoding.UTF8.GetBytes(requestBody);
            myRequest.ContentLength = aryBuf.Length;
            using (Stream writer = myRequest.GetRequestStream())
            {
                writer.Write(aryBuf, 0, aryBuf.Length);
                writer.Close();
                writer.Dispose();
            }

            HttpWebResponse response = (HttpWebResponse)myRequest.GetResponse();
            Stream responseStream = response.GetResponseStream();
            StreamReader reader = new StreamReader(responseStream, Encoding.GetEncoding("gb2312"));
            string content = reader.ReadToEnd();
            reader.Close();
            responseStream.Close();
            webBrowser1.DocumentText = content;
        }

        public class ACCESS_TOKEN
        {
            public string access_token { get; set; }
            public string expires_in { get; set; }
            public string application { get; set; }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            webBrowser1.ScriptErrorsSuppressed = true;
            //pictureBox1.ImageLocation = "http://portal.idcicp.com/tools/Verify_code.ashx";
        }

        private void btnRefresh_Click(object sender, EventArgs e)
        {
            timer1.Start();
        }

        private string labMsg = string.Empty;
        //private int iLater = 4;
        private void timer1_Tick(object sender, EventArgs e)
        {
            log("开始获取最新注册用户信息，请稍等......");
            timer1.Stop();//先停止，待获取完成再启动
            txtParmater.Text = "当前上下文参数：" + System.Environment.NewLine
                + "__cfduid =" + cookie_cfduid + System.Environment.NewLine
                 + "ASP.NET_SessionId =" + cookie_AspNetSessionId + System.Environment.NewLine
                  + "sessionCode_ =" + cookie_SessionCode + System.Environment.NewLine
                   + "sessionUser_ =" + cookie_SessionUser + System.Environment.NewLine
                    + cookie_ASPSESSIONIDName + "=" + cookie_ASPSESSIONIDValue + System.Environment.NewLine
                     + "month_champion =" + cookie_month_champion;


            GetCustomerList();
        }

    }
}
