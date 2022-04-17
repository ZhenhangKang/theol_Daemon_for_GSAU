using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Text;
using System.Diagnostics;
using System.Threading.Tasks;
using System.Windows.Forms;
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using System.IO;
using System.Threading;
using System.Collections.ObjectModel;

namespace THEOL_daemon_sharp
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void button1_Click(object sender, EventArgs e)
        {
           
            if (textBox1.Text == "" || textBox2.Text == "")
            {
                MessageBox.Show("请输入正确的账号和密码！");
            }
            else if(numericUpDown5.Value <= 0)
            {
                numericUpDown5.Value = 1;
                MessageBox.Show("视频循环次数有误，请重新设置！");
            }
            else
            {
                var browser = new FirefoxDriver();
                //Selenium 登录
                browser.Navigate().GoToUrl("http://wljx.gsau.edu.cn");
                string loginBtn = "//*[@id=\"loginbtn\"]";
                string userName = "//*[@id=\"userName\"]";
                string pwd = "//*[@id=\"passWord\"]";
                Thread.Sleep(3000);
                //登录阶段开始
                string account = textBox1.Text;
                string password = textBox2.Text;
                string result = "";
                loginProcess(browser, loginBtn, userName, pwd, account, password, out result);
                textBox10.AppendText(DateTime.Now.ToString() + result + "\r\n");
                Thread.Sleep(1000);
                //进入选中课程
                browser.FindElement(By.XPath(@"/html/body/div[1]/div[3]/div/div/div/div[3]/div[3]/div/div/div[2]/ul/li[6]/div[2]/div/a[1]")).Click();

                Thread.Sleep(1000);
                int loops = Convert.ToInt32(numericUpDown5.Value);
                ReadOnlyCollection<string> windowsHandles;
                //刷时间线程
                Thread Thread_1 = new Thread(() =>
                {
                    for (int j = 0; j < loops; j++)
                    {
                        windowsHandles = browser.WindowHandles;
                        browser.SwitchTo().Window(windowsHandles[1]);
                        textBox10.AppendText(DateTime.Now.ToString() + "正在进入学习..." + "\r\n");

                        try
                        {
                            //课程界面刷课件在线时间
                            browser.FindElement(By.XPath(@"/html/body/div[3]/div[1]/div[2]/div[2]/ul/li[3]/a")).Click();
                            Thread.Sleep(2000);
                            browser.FindElement(By.XPath(@"/html/body/div[3]/div[2]/div[1]/div[2]/div[2]/ul[3]/li[1]/a")).Click();
                            Thread.Sleep(2000);
                            browser.FindElement(By.XPath(@"/html/body/div[3]/div[2]/div[1]/div[2]/div[2]/ul[3]/li[1]/ul/li[1]/a")).Click();
                            Thread.Sleep(Convert.ToInt32(numericUpDown1.Value) * 60 * 1000);
                            textBox10.AppendText(DateTime.Now.ToString() + "课件学习时间结束，将进行下一过程" + "\r\n");
                        }
                        catch
                        {
                            textBox10.AppendText(DateTime.Now.ToString() + "课件学习异常！！" + "\r\n");
                        }
                        //刷视频
                        for (int i = 0; i < 3; i++)
                        {
                            try
                            {
                                listBox1.SelectedIndex = i;
                                string video_link;
                                video_link = "window.open()";
                                textBox10.AppendText(DateTime.Now.ToString() + "跳转到 " + listBox1.SelectedItem.ToString() + "\r\n");
                                browser.ExecuteScript(video_link);
                                windowsHandles = browser.WindowHandles;
                                browser.SwitchTo().Window(windowsHandles[i + 2]);
                                textBox10.AppendText(DateTime.Now.ToString() + "成功切换到新窗口" + "\r\n");
                                try
                                {
                                    //视频播放 创建按钮点击
                                    string redirUrl = "return window.location.href = '" + listBox1.SelectedItem.ToString() + "';";
                                    browser.ExecuteScript(redirUrl);
                                    textBox10.AppendText("已经载入页面");
                                    string createBtn = "setTimeout(function(){var htmlObj=document.getElementsByClassName('wrap')[0],submitButton=document.createElement('button');submitButton.setAttribute('type','button'),submitButton.setAttribute('id','actBtn'),submitButton.innerText='test';htmlObj.appendChild(submitButton)},1000);";
                                    browser.ExecuteScript(createBtn);
                                    // 用 selenium 点击 id 为 actBtn 的按钮（等待元素加载出来再点击）
                                    Thread.Sleep(10000);
                                    browser.FindElement(By.Id("actBtn")).Click();
                                    string playMedia = "setInterval(function(){document.querySelector('video').play();},1000);";
                                    browser.ExecuteScript(playMedia);
                                }
                                catch
                                {
                                    textBox10.AppendText(DateTime.Now.ToString() + "鼠标移动失败！" + "\r\n");
                                }

                                if (i == 0)
                                {
                                    int time_1 = (int)numericUpDown2.Value;
                                    Thread.Sleep(time_1 * 60 * 1000);
                                }
                                else if (i == 1)
                                {
                                    int time_2 = (int)numericUpDown3.Value;
                                    Thread.Sleep(time_2 * 60 * 1000);
                                }
                                else if (i == 2)
                                {
                                    int time_3 = (int)numericUpDown4.Value;
                                    Thread.Sleep(time_3 * 60 * 1000);
                                }
                                int count = i + 1;
                                string str = "成功学习第" + count + "个视频";
                                textBox10.AppendText(DateTime.Now.ToString() + str + "\r\n");

                            }
                            catch
                            {
                                textBox10.AppendText(DateTime.Now.ToString() + "视频学习错误！" + "\r\n");
                            }

                        }
                        windowsClose(browser, windowsHandles);
                        int count_loops = j + 1;
                        textBox10.AppendText(DateTime.Now.ToString() + "第" + count_loops + "循环已完成" + "\r\n");
                    }
                });
                Thread_1.Start();
                textBox10.AppendText(DateTime.Now.ToString() + "线程状态：" + Thread_1.ThreadState.ToString() + "\r\n");
            }

        }
        //登录Umooc
        protected static void loginProcess(FirefoxDriver browser,string loginBtn,string userName,string pwd,string account,string password,out string result)
        {
            try
            {
                browser.FindElement(By.XPath(loginBtn)).Click();
                browser.FindElement(By.XPath(userName)).SendKeys(account);
                browser.FindElement(By.XPath(pwd)).SendKeys(password);
                browser.FindElement(By.XPath(@"/html/body/div[1]/div[1]/div/div[3]/div/form/div[2]/input")).Click();
                result = "登录成功";
                
            }
            catch
            {
                result = "登录异常";
                MessageBox.Show("登录异常");
            }

        }
        //关闭视频选项卡
        void windowsClose(FirefoxDriver browser, ReadOnlyCollection<string> windowsHandles)
        {
            browser.Close();
            windowsHandles = browser.WindowHandles;
            browser.SwitchTo().Window(windowsHandles[3]);
            browser.Close();
            windowsHandles = browser.WindowHandles;
            browser.SwitchTo().Window(windowsHandles[2]);
            browser.Close();
         
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            Control.CheckForIllegalCrossThreadCalls = false;
            listBox1.Enabled = false;
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            MessageBox.Show("作者：ZhenhengKang" + "\r\n" + "\r\n" + "基于Selenium和GeckoDriver开发" + "\r\n" + "根据以往Python语言版本的C#重构" + "\r\n" + "需要安装FireFox浏览器" + "\r\n" +"2022年4月");
        }
        //写入配置文件
        private void button3_Click(object sender, EventArgs e)
        {
            try 
            {
                FileStream fsWriter = new FileStream(@"parameters.txt", FileMode.OpenOrCreate, FileAccess.Write);
                string writeString = textBox1.Text + "\r\n" + textBox2.Text + "\r\n" + numericUpDown1.Value + "\r\n" + numericUpDown2.Value + "\r\n"
                    + numericUpDown3.Value + "\r\n" + numericUpDown4.Value + "\r\n" + numericUpDown5.Value;
                byte[] buffer = new byte[1024 * 1024 * 1];
                buffer = Encoding.UTF8.GetBytes(writeString);
                fsWriter.Write(buffer, 0, buffer.Length);
                MessageBox.Show("写入成功！");
            }
            catch
            {
                MessageBox.Show("文件或被占用，写入失败！");
            }
        }
        //读取配置文件
        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                string[] lines = File.ReadAllLines(@"parameters.txt");
                textBox1.Text = lines[0];
                textBox2.Text = lines[1];
                numericUpDown1.Value = Convert.ToDecimal(lines[2]);
                numericUpDown2.Value = Convert.ToDecimal(lines[3]);
                numericUpDown3.Value = Convert.ToDecimal(lines[4]);
                numericUpDown4.Value = Convert.ToDecimal(lines[5]);
                numericUpDown5.Value = Convert.ToDecimal(lines[6]);
                MessageBox.Show("读取成功！");
            }
            catch
            {
                MessageBox.Show("未在相同文件夹下找到参数文件！" + "\r\n" + "请重新检查是否有parameter.txt文件");
            }
            
        }

       
    }
    
}
