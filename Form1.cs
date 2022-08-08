using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;


/*
 A. 輸入單一關鍵字情況
 Step 1: 根據關鍵字搜尋所有廠商的網站
 Step 2: 分析與關鍵字相似度
 Step 3: 將分析的內容顯示在視窗上
 
 B. 單純更新商品列表
 Step 1: 讀取檔案
 Step 2: 根據檔案內容去搜尋目標網站
 Step 3: 分析目標網站內容        
 Step 4: 將分析的內容重新在寫入檔案
 Step 5: 根據分析的內容發佈在電商平台
 
 */

/*
* 讀取 excel
* 
* 1: 條碼
* 2. 辨識用商品型號
* 3. 商品名稱 日文
* 4. 商品名稱 中文
* 5. 價格 日幣
* 6. 成本
* 7. MJ 價格
* 8. 自己價格
* 9. 狀態
* 
*/



namespace onlineshop
{
    public partial class Form1 : Form
    {
        const string rutenaccount = "daniel51215";
        const string rutenpassword = "d7410852";

        string[,] exceltable;               // 追蹤商品     
        string[,] stocktable;               // 庫存商品
        int rowcount;                       // exceltable row
        int columnCount;                    // exceltable column
        int rowcount_stock;                 // stock row
        int columncount_stock;              // stock column       
        string[,] tomixtable;               // Tomix 官網全商品
        int number_tomixproduct;            // Tomix 官網商品數量

        int number_all;
        int number_case = 0;
        int id_updateproduct = 0;

        Thread threadexceltable;            // 追蹤商品處理執行緒
        Thread threadstocktable;            // 庫存商品處理執行緒
        string casestr;                     // 處理階段

        Thread threadcatchpic;              //抓取圖片執行緒
        Thread threaddoc;

        /*declaration of html element*/
        HtmlElement headermybid;
        HtmlElementCollection acollection;
        HtmlElementCollection formcollection;
        HtmlElementCollection inputcollection;
        HtmlElementCollection useridcollection;
        HtmlElementCollection passwordcollection;
        HtmlElement userid;
        HtmlElement password;

        HtmlElementCollection form;
        HtmlElement search;
        HtmlElement btn_search;

        HtmlElement top;
        HtmlElementCollection rt_tablecollection;
        HtmlElement rt_table;
        HtmlElementCollection rt_tbodycollection;
        HtmlElement rt_tbody;
        HtmlElementCollection trcollection;
        HtmlElement tr;
        HtmlElementCollection tdcollection;
        HtmlElement td;
        HtmlElementCollection acollcection;
        HtmlElement a;

        HtmlElementCollection tablecollection;
        HtmlElement producttable;
        HtmlElement tradetable;
        HtmlElementCollection productinputcollection;
        HtmlElementCollection tradeinputcollection;

        HtmlElement btn_submit;
        bool loading = false;
        bool picweb = false;
        // static readonly HttpClient client = new HttpClient();

        public Form1()
        {
            InitializeComponent();
            webBrowser1.DocumentCompleted += new WebBrowserDocumentCompletedEventHandler(webdoccomplete);
            webBrowser1.ScriptErrorsSuppressed = true;
            threadexceltable = new Thread(methodforproductexcel);
            threadstocktable = new Thread(methodforstockexcel);
            threaddoc = new Thread(methodfordocstatus);
            threadcatchpic = new Thread(methodforpic);
            
        }
        /* From1_Load 功能說明
         * 當視窗載入時輸入:
         *  a. 追蹤之 Tomix 商品檔案
         *  b. 現有庫存之檔案
         *  c. 觸發 Thread 讀取追蹤 Tomix 商品檔案
         *  d. 觸發 Thread 讀取庫存之檔案
         *  e. 將上述兩個讀取檔案轉換到 Table
         */
        private void Form1_Load(object sender, EventArgs e)
        {
            txtbx_inputfile.Text = @"C:\Users\weiting\Desktop\Tomix.xlsx";
            txtbx_inputfilestock.Text = @"C:\Users\weiting\Desktop\STOCK.xlsx";
            // btn_analyze.PerformClick();
            // Application.DoEvents();
            /* try
             {
                 // HttpResponseMessage response = await client.GetAsync("https://mybid.ruten.com.tw/master/my.php?l_type=sel_wait_confirm");
                 // response.EnsureSuccessStatusCode();
                 // string responsebody = await response.Content.ReadAsStringAsync();
                 // Above three lines can be replaced with new helper method below
                 // string responseBody = await client.GetStringAsync(uri);

                 //   Console.WriteLine(responsebody);

                 webBrowser1.Navigate("https://member.ruten.com.tw/user/login.htm");                
             }
             catch (Exception ex)
             {
                 MessageBox.Show(ex.ToString());
             }
             */
            btn_filecheck.PerformClick();
            btn_filecheckstock.PerformClick();

        }
        /* btn_fileview_Click 功能說明
         * 
         * 按下" 瀏覽 " 按鈕 觸發新視窗 並將原始位置設定在 C:\Users\weiting\Desktop\
         */
        private void btn_fileview_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.OpenFileDialog openFileDialog = new System.Windows.Forms.OpenFileDialog();

            openFileDialog.InitialDirectory = @"C:\Users\weiting\Desktop\";

            openFileDialog.DefaultExt = ".txt";
            openFileDialog.Filter = "Text Document (.txt)|*.txt";

            openFileDialog.Multiselect = true;
            openFileDialog.Filter = "All flie(*.*)|*.*";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                txtbx_inputfile.Text = openFileDialog.FileName;
            }
        }
        /* btn_fileviewstock_Click 功能說明
         * 
         * 按下" 瀏覽 " 按鈕 觸發新視窗 並將原始位置設定在 C:\Users\weiting\Desktop\
         */
        private void btn_fileviewstock_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.OpenFileDialog openFileDialog = new System.Windows.Forms.OpenFileDialog();

            openFileDialog.InitialDirectory = @"C:\Users\weiting\Desktop\";

            openFileDialog.DefaultExt = ".txt";
            openFileDialog.Filter = "Text Document (.txt)|*.txt";

            openFileDialog.Multiselect = true;
            openFileDialog.Filter = "All flie(*.*)|*.*";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                txtbx_inputfile.Text = openFileDialog.FileName;
            }
        }
        /* btn_filecheck_Click 功能說明
         * 
         * 按下" 確認 " 按鈕 觸發新視窗 並將原始位置設定在 C:\Users\weiting\Desktop\
         * 為觸發執行緒 功能改寫在 methodforproductexcel() 函式裡
         * 下方為原功能
         * try
         *   {
         *      string path = txtbx_inputfile.Text;
         *      Excel.Application application = new Excel.Application();
         *      Excel.Workbook workbook = application.Workbooks.Open(path);
         *      Excel.Worksheet worksheet = workbook.Sheets[1];
         *      Excel.Range range = worksheet.UsedRange;
         *
         *      rowCount = range.Rows.Count;
         *      columnCount = range.Columns.Count;
         *
         *      exceltable = new string[rowCount, columnCount];
         *
         *      for (int i = 1; i < rowCount; i++)
         *      {
         *          for (int j = 1; j < columnCount; j++)
         *          {
         *              if (j == 1) txtbx_data.Text += Environment.NewLine;
         *              if (range.Cells[i, j] != null && range.Cells[i, j].Value2 != null)
         *              {
         *                 txtbx_data.Text += "[ " + i.ToString() + ", " + j.ToString() + " ] = ";
         *                 txtbx_data.Text += range.Cells[i, j].Value2.ToString() + Environment.NewLine;
         *                 exceltable[i, j] = range.Cells[i, j].Value2.ToString();
         *
         *              }
         *              else
         *              {
         *                 exceltable[i, j] = " ";
         *                 txtbx_data.Text += "[ " + i.ToString() + ", " + j.ToString() + " ] = ";
         *                //txtbx_data.Text += allproduct[i, j] + Environment.NewLine;
         *              }
         *          }
         *      }
         * }
         * catch (Exception ex)
         * {
         *     throw new NotImplementedException(); ;
         * }
         */
        private void btn_filecheck_Click(object sender, EventArgs e)
        {
            threadexceltable.Start();
        }
        /* btn_filecheckstock_Click 功能說明
         * 
         * 按下" 確認 " 按鈕 觸發新視窗 並將原始位置設定在 C:\Users\weiting\Desktop\
         * 為觸發執行緒 功能改寫在 methodforproductexcel() 函式裡
         * 下方為原功能
         *
         *   string path = txtbx_inputfilestock.Text;
         *   Excel.Application application = new Excel.Application();
         *   Excel.Workbook workbook = application.Workbooks.Open(path);
         *   Excel.Worksheet worksheet = workbook.Sheets[1];
         *   Excel.Range range = worksheet.UsedRange;
         *
         *   rowcount_stock = range.Rows.Count;
         *   columncount_stock = range.Columns.Count;
         *
         *   stocktable = new string[rowcount_stock, columncount_stock];
         *
         *   for (int i = 1; i < rowcount_stock; i++)
         *  {
         *       for (int j = 1; j < columncount_stock; j++)
         *      {
         *          if (j == 1) txtbx_data.Text += Environment.NewLine;
         *          if (range.Cells[i, j] != null && range.Cells[i, j].Value2 != null)
         *          {
         *               stocktable[i, j] = range.Cells[i, j].Value2.ToString();
         *           }
         *          else
         *          {
         *              stocktable[i, j] = " ";
         *
         *          }
         *      }
         *  }
         * 
         *
         */
        private void btn_filecheckstock_Click(object sender, EventArgs e)
        {
            threadstocktable.Start();
        }
        /* btn_webview 功能說明
         * 將 webbrowser 導向 Tomix 庫存網站
         *
         *下方程式碼曾測試過 但不理想
         * 
               HttpResponseMessage response = await client.GetAsync("https://www.tomytec.co.jp/tomix/stocklist/detail.cgi");
               response.EnsureSuccessStatusCode();
               string responsebody = await response.Content.ReadAsStringAsync();                
               // Above three lines can be replaced with new helper method below
               // string responseBody = await client.GetStringAsync(uri);

               Console.WriteLine(responsebody);
               txtbx_web.Text = responsebody;
               alldata = responsebody;

         */
        private void btn_webview(object sender, EventArgs e)
        {
            try
            {               
                webBrowser1.Navigate("https://www.tomytec.co.jp/tomix/stocklist/detail.cgi");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }            
        }
        /* btn_analyze_Click 功能說明
         * 
         * 分析 Tomix 庫存官網 細部解說再函式裡有說明
         *
         *  下方為簡易測試功能 測試 Sendkeys.Send() 及 button.InvokeMember()
         *  The goal of following eight line is auto enter the value up to what you want to find                         
            webBrowser1.Document.GetElementsByTagName("button")[0].MouseDown += new HtmlElementEventHandler(ClickTest);           
            HtmlElementCollection buttoncollection = webBrowser1.Document.GetElementsByTagName("button");
            HtmlElement button = buttoncollection[0];
            webBrowser1.Document.GetElementsByTagName("input")[0].SetAttribute("value", "");
            webBrowser1.Update();            
            txtbx_data.Text = webBrowser1.DocumentText.ToString();      
            button.InvokeMember("click");
            SendKeys.Send("{Enter}");
         */
        private void btn_analyze_Click(object sender, EventArgs e)
        {            
            /*
             * The goal of following line is to transfer the table caught in website to array    
             * The data is caught is up to the label of html
             * First HtmlElementCollection is an array , it contains all the tag of tbody in the website
             * Second HtmlElementCollection is also an array , it contains all the tag of tr in the specific tbody
             */

            HtmlElementCollection tbodycollection = webBrowser1.Document.GetElementsByTagName("tbody");

            int total = tbodycollection.Count;

            HtmlElement[] tbody = new HtmlElement[total];
            for (int i = 0; i < total; i++)
            {
                tbody[i] = tbodycollection[i];
                //txtbx1.Text += tbodycollection[i].InnerText.ToString();
            }
            trcollection = tbody[1].GetElementsByTagName("tr");
            int trtotal = trcollection.Count;
            number_tomixproduct = trcollection.Count;
            number_all = trcollection.Count;
            txtbx_data.Text += "trtotal:" + trtotal.ToString() + Environment.NewLine;
            HtmlElement[] tr = new HtmlElement[trtotal];
            for (int j = 0; j < trtotal; j++)
            {
                tr[j] = trcollection[j];
                //  txtbx_data.Text += "j = " + j.ToString() + "\t";
                // txtbx_data.Text += tr[j].InnerText.ToString()+Environment.NewLine;
                //  txtbx_data.Update();

            }
            /*
             * The goal of following code is try to split all the tag <td> in every tag <tr> and write into 2 dimension array
             */
            HtmlElement[,] td = new HtmlElement[trtotal, 8];
            tomixtable = new string[trtotal, 8];
            for (int z = 0; z < trtotal; z++)
            {
                HtmlElementCollection tdcollection = tr[z].GetElementsByTagName("td");
                if (tdcollection.Count == 0)
                {
                    for (int k = 0; k < 8; k++)
                    {
                        tomixtable[z, k] = "";
                    }
                }
                else
                {
                    for (int k = 0; k < tdcollection.Count; k++)
                    {
                        td[z, k] = tdcollection[k];
                        if (td[z, k].InnerText == null)
                        {
                            tomixtable[z, k] = "";
                        }
                        else
                        {
                            tomixtable[z, k] = td[z, k].InnerText.ToString();
                        }
                        // txtbx_data.Text += table[z, k];
                        // txtbx_data.Text += ", ";
                    }
                }
                // txtbx_data.Text += Environment.NewLine;
                // txtbx_data.Update();
            }
            //txtbx_data.Text += "j = " + 1.ToString() + "\t";
            //txtbx_data.Text += tr[1].InnerText.ToString() + Environment.NewLine;
            Console.WriteLine("FINISH TOMIX WEBSITE ANALYSIS");
        }
        /* btn_updateexcel_Click 功能說明
         * 前置條件
         * webbrowser 必須在 Tomix 庫存查詢
         * 更新追蹤之 Tomix 商品的狀態
         * 
         * 步驟如下
         * 1. 讀取 EXCEL 
         * 2. 分析網頁
         * 3. 根據網頁更新 EXCEL
         *
         */
        private void btn_updateexcel_Click(object sender, EventArgs e)
        {
            btn_website.PerformClick();
            //btn_website.Update();
            while (true)
            {
                if (casestr == "")
                {
                    Application.DoEvents();
                    continue;
                }
                else if (casestr == null)
                {
                    Application.DoEvents();
                    continue;
                }
                else if (casestr == "TOMIX")
                {
                    btn_analyze.PerformClick();
                    break;
                }
            }
            
            string path = txtbx_inputfile.Text;
            Excel.Application application = new Excel.Application();
            Excel.Workbook workbook = application.Workbooks.Open(path);
            Excel.Worksheet worksheet = workbook.Sheets[1];
            Excel.Range range = worksheet.UsedRange;

            int excelrowCount = range.Rows.Count;
            int excelrowColumn = range.Columns.Count;

            exceltable = new string[excelrowCount, excelrowColumn];
            /*
             * 讀取 excel
             * 
             * 1: 條碼
             * 2. 辨識用商品型號
             * 3. 商品名稱 日文
             * 4. 商品名稱 中文
             * 5. 價格 日幣
             * 6. 成本
             * 7. MJ 價格
             * 8. 自己價格
             * 9. 狀態
             * 
             */
            Console.WriteLine("讀取追蹤之 excel ...");
            for (int i = 1; i < excelrowCount; i++)
            {
                for (int j = 1; j < excelrowColumn; j++)
                {
                    //      if (j == 1) txtbx_data.Text += Environment.NewLine;
                    if (range.Cells[i, j] != null && range.Cells[i, j].Value2 != null)
                    {
                        //  txtbx_data.Text += "[ " + i.ToString() + ", " + j.ToString() + " ] = ";
                        //   txtbx_data.Text += range.Cells[i, j].Value2.ToString() + Environment.NewLine;
                        exceltable[i, j] = range.Cells[i, j].Value2.ToString();
                        /* MessageBox.Show("i = " + i.ToString() + Environment.NewLine
                                        + "j = " + j.ToString() + Environment.NewLine
                                        + "Cells[" + i.ToString() + "," + j.ToString() + "] =" + range.Cells[i, j].Value2.ToString());
                        */
                    }
                    else
                    {
                        exceltable[i, j] = " ";
                        //  txtbx_data.Text += "[ "+i.ToString()+", " +  j.ToString()+" ] = ";
                        //  txtbx_data.Text += exceltable[i, j] + Environment.NewLine;
                    }
                }
            }

            /*
             * 分析網頁
             * 抓取 已經有的商品型號 將 日幣 條碼 庫存 預計生產時間              
             * table 為抓取網頁的內容
             * 0: 商品號碼
             * 1: 商品名稱
             * 2: 日幣價格
             * 3: 庫存顯示
             * 4: 當月
             * 5: 下個月
             * 6: 下下月庫存
             * 7: 條碼
             *
             *△ = 在庫あり
             *▲ = 在庫品薄
             *末○ = 8月末出荷予定
             *○ = 生産予定あり
             *NEW  = 新製品
             */

            /*
             *Follwing code is to match the string in excel of product number and the string in web of product number.
             *If both of two is the same , update the cells in excel of 
             * 1. product name in japanese
             * 2. barcode of product
             * 3. the price of product in yen
             * 4. the stock status of product
             */
            Console.WriteLine("讀取分析 Tomix 網頁  ...");
            int excelcount = 0;
            int tablecount = 0;
            while (excelcount < excelrowCount)
            {
                while (tablecount < number_all)
                {
                    if (exceltable[excelcount, 2] == "5518")
                    {
                        if (tomixtable[tablecount, 0] == "5518")
                        {
                            //MessageBox.Show("TEST");
                        }
                    }
                    if (tomixtable[tablecount, 0].Equals(exceltable[excelcount, 2]))
                    {
                        /**MessageBox.Show("table[tablecount,0]=" + table[tablecount, 0] + Environment.NewLine +
                                        "exceltable[excelcount, 2] ="+exceltable[excelcount, 2]);
                        */
                        // Product name in Japanese
                        exceltable[excelcount, 3] = tomixtable[tablecount, 1];
                        // Barcode
                        exceltable[excelcount, 1] = tomixtable[tablecount, 7];
                        // Price in yen
                        exceltable[excelcount, 5] = tomixtable[tablecount, 2];
                        // Stock Status
                        if (tomixtable[tablecount, 3] != "")
                        {
                            if (tomixtable[tablecount, 3] == "NEW")
                            {
                                exceltable[excelcount, 9] = "新品";
                            }
                            else
                            {
                                exceltable[excelcount, 9] = "有庫存";
                            }
                        }
                        else if (tomixtable[tablecount, 4] != "")
                        {
                            exceltable[excelcount, 9] = "有庫存";
                        }
                        else if (tomixtable[tablecount, 5] != "")
                        {
                            exceltable[excelcount, 9] = "有庫存";
                        }
                        else if (tomixtable[tablecount, 6] != "")
                        {
                            exceltable[excelcount, 9] = tomixtable[tablecount, 7];
                        }
                        else
                        {
                            exceltable[excelcount, 9] = "近期不會有貨";
                        }

                        tablecount++;
                        break;
                    }
                    else
                    {
                        tablecount++;
                        continue;
                    }
                }
                excelcount++;
                tablecount = 0;
            }



            /*
             *  Rewrite the information into the excel                          
             */

            Console.WriteLine("重新寫入 excel 中...");
            for (int i = 1; i < excelrowCount; i++)
            {
                if (i == 10)
                {
                    // MessageBox.Show("i = 10");
                }
                for (int j = 1; j < excelrowColumn; j++)
                {
                    if (exceltable[i, j] == null)
                    {
                        continue;
                    }
                    else
                    {
                        range.Cells[i, j].Value2 = exceltable[i, j];
                    }
                }
            }
            application.DisplayAlerts = false;
            //   workbook.SaveAs("updateTomix.xlsx");                                               
            workbook.Save();//備份
            workbook.SaveAs(@"C:\Users\weiting\Desktop\Tomix.xlsx");
            workbook.Close();
            application.Quit();
            Console.WriteLine("FINISH 更新追蹤商品excel");
        }
        /* btn_updateproduct_Click 功能說明
         * 
         * Following function is update the ruten product according to the excel
         * Step 1: use the webbrowser connect to the ruten platform
         * Step 2: On the basis of the textbook or the information saved in the computer, login the ruten
         * Step 3: Go to the product in the ruten and update the status
         * 
         * There are five status of webbrowser loading
         * 
         * WebBrowserReadyState.Complete  = all controls have loaded
         * WebBrowserReadyState.Interactive  = have already loaded most of the controls ,and can interact with user strictly
         * WebBrowserReadyState.Loaded  = already loaded a new documentand initialize new document but not receive all the document
         * WebBrowserReadyState.Loading  = on loading new document
         * WebBrowserReadyState.Uninitialized  = nothing loaded
         *
         */
        private void btn_updateproduct_Click(object sender, EventArgs e)
        {
            number_case = 0;
            try
            {
                webBrowser1.Navigate(new Uri("https://member.ruten.com.tw/user/login.htm"));

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        /* webdoccomplete 功能說明
         *
         * 針對跑完網頁後不同的 Url 有不同的動作
         * 有針對 webbbrowser 做測試 只有同時 Url 和 ReadyState 都是正確的狀態時
         * 才能認定 webbrowser 完成
         *
         * 下方式測試 webbrowser 完成的測試碼
         *  //  txtbx_web.Text += "call:" + sender.ToString() + Environment.NewLine;
            //  txtbx_data.Text += count_controls.ToString() + ",";
            //webBrowser1.Document.Body.InnerHtml.ToString();
            
            if(webBrowser1.ReadyState == WebBrowserReadyState.Complete)
            {
                txtbx_data.Text += "Complete" + Environment.NewLine;
            }
            else if (webBrowser1.ReadyState == WebBrowserReadyState.Interactive)
            {
                txtbx_data.Text += "Interactive" + Environment.NewLine;
            }
            else if (webBrowser1.ReadyState == WebBrowserReadyState.Loaded)
            {
                txtbx_data.Text += "Loaded" + Environment.NewLine;
            }
            else if (webBrowser1.ReadyState == WebBrowserReadyState.Loading)
            {
                txtbx_data.Text += "Loading" + Environment.NewLine;
            }
            else if (webBrowser1.ReadyState == WebBrowserReadyState.Uninitialized)
            {
                txtbx_data.Text += "Uninitialized" + Environment.NewLine;
            }
            else
            {
                txtbx_data.Text += "No Matching" + Environment.NewLine;                
            }
            
            if (!webBrowser1.IsBusy)
            {
                if (webBrowser1.ReadyState == WebBrowserReadyState.Complete)
                {
                }           
            }
            Console.WriteLine("e:"+e.Url.ToString());
            Console.WriteLine("DOC:"+webBrowser1.Document.Url.ToString());
         */
        private void webdoccomplete(object sender, WebBrowserDocumentCompletedEventArgs e)
        {          
            if ((webBrowser1.Document.Url == e.Url) && (webBrowser1.ReadyState == WebBrowserReadyState.Complete))
            {
                Console.WriteLine("complete:" + e.Url.ToString());
                if (webBrowser1.Document.Title == "露天拍賣-台灣 NO.1 拍賣網站")
                {
                    if (number_case == 0)
                    {
                        casestr = "HOMEPAGE";
                    }
                    else if (number_case == 1)
                    {
                        casestr = "MYBID";
                    }
                    else
                    {
                        casestr = "";
                    }

                }
                else if (webBrowser1.Document.Title == "露天會員登入 | 露天拍賣 - 台灣 NO.1 拍賣網站")
                {
                    casestr = "LOGIN";
                }
                else if (webBrowser1.Document.Title == "出售中商品 | 露天拍賣-台灣 NO.1 拍賣網站")
                {
                    if (number_case == 5)
                    {
                        casestr = "UPDATE_STEP1";
                    }
                    else if (number_case == 6)
                    {

                        casestr = "UPDATE_STEP2";
                    }
                    else
                    {
                        casestr = "";
                    }
                }
                else if (webBrowser1.Document.Title == "商品編輯 | 露天拍賣-台灣 NO.1 拍賣網站")
                {
                    casestr = "EDITPAGE";
                }
                else if (webBrowser1.Document.Title.Equals("製品リスト｜鉄道模型 TOMIX 公式サイト｜株式会社トミーテック"))
                {
                    casestr = "TOMIX";
                }
                else
                {
                    casestr = "SUCCESS";
                }

                switch (casestr)
                {
                    case "HOMEPAGE":
                        txtbx_data.Text += "HomePage" + Environment.NewLine;
                        headermybid = webBrowser1.Document.GetElementById("header_mybid");
                        headermybid.InvokeMember("click");
                        Thread.Sleep(1000);
                        casestr = "";

                        number_case++;
                        view_parameter();
                        break;

                    case "MYBID":
                        txtbx_data.Text += "MYBID" + Environment.NewLine;
                        //   MessageBox.Show(webBrowser1.Document.GetElementsByTagName("a").Count.ToString());
                        acollection = webBrowser1.Document.GetElementsByTagName("a");
                        //HtmlElement a = acollection[0].GetAttribute("")
                        for (int i = 0; i < acollection.Count; i++)
                        {
                            //txtbx_data.Text += acollection[i].GetAttribute("href").ToString() + Environment.NewLine;
                            if (acollection[i].GetAttribute("href").ToString().Equals("https://mybid.ruten.com.tw/master/my.php?l_type=sel_selling"))
                            {
                                acollection[i].InvokeMember("click");
                                Thread.Sleep(1000);
                                number_case++;
                                continue;
                            }
                        }
                        number_case++;
                        view_parameter();
                        break;

                    case "LOGIN":
                        txtbx_data.Text += "LOGIN" + Environment.NewLine;
                        formcollection = webBrowser1.Document.GetElementsByTagName("form");
                        inputcollection = formcollection[0].GetElementsByTagName("input");
                        useridcollection = formcollection.GetElementsByName("userid");
                        passwordcollection = formcollection.GetElementsByName("userpass");

                        userid = webBrowser1.Document.GetElementById("userid");
                        password = webBrowser1.Document.GetElementById("userpass");

                        /*  MessageBox.Show(formcollection.Count.ToString() + ","
                                          + inputcollection.Count.ToString() + ","
                                          + useridcollection.Count.ToString() + ","
                                          + passwordcollection.Count.ToString()
                                          + userid.Id.ToString() + ","
                                          + password.Id.ToString() + ","
                                          + inputcollection[0].Name.ToString() + ","
                                          + inputcollection[1].Name.ToString() + ","
                                          + inputcollection[2].Name.ToString() + ","

                                          );
                        */
                        userid.InnerText = rutenaccount;
                        password.InnerText = rutenpassword;
                        //inputcollection[0].InnerText = rutenaccount;
                        //inputcollection[1].InnerText = rutenpassword;
                        webBrowser1.Update();
                        inputcollection[2].InvokeMember("click");
                        Thread.Sleep(1000);
                        view_parameter();
                        break;

                    case "UPDATE_STEP1":
                        txtbx_data.Text += "UPDATE_STEP1" + Environment.NewLine;
                        form = webBrowser1.Document.Forms;
                        search = webBrowser1.Document.GetElementById("s_content");
                        btn_search = webBrowser1.Document.GetElementById("search_button");


                        /*
                         * Following start to update the ruten website according to the excel tabel
                         * Step 1 : choose the product name
                         * Step 2 : update the product name to filter
                         * Step 3 : trigger the click event of search button
                         * Step 4 : trigger the edit button of the product
                         * Step 5 : anaylze the html of detail of product and update
                         * Step 6 : turn into step until all the product have update
                         */

                        //for(int i = 0; i < rowCount; i++)

                        search.InnerText = stocktable[2, 2];
                        //search.InnerText = "5518";

                        webBrowser1.Update();

                        btn_search.Focus();
                        SendKeys.Send("{Enter}");
                        // btn_search.InvokeMember("click");
                        // btn_search.InvokeMember("post");
                        Thread.Sleep(100);
                        number_case++;
                        view_parameter();
                        break;

                    case "UPDATE_STEP2":
                        txtbx_data.Text += "UPDATE_STEP2" + Environment.NewLine;
                        top = webBrowser1.Document.GetElementById("top");
                        rt_tablecollection = top.GetElementsByTagName("table");
                        rt_table = rt_tablecollection[1];
                        rt_tbodycollection = rt_table.GetElementsByTagName("tbody");
                        rt_tbody = rt_tbodycollection[0];
                        trcollection = rt_tbody.GetElementsByTagName("tr");
                        tr = trcollection[0];
                        tdcollection = tr.GetElementsByTagName("td");
                        td = tdcollection[1];
                        acollcection = td.GetElementsByTagName("a");
                        a = acollcection[0];
                        // HtmlElementCollection tdtrcollection = td.GetElementsByTagName("tr");
                        //HtmlElement tdtr = tdtrcollection[1];
                        a.InvokeMember("click");

                        //td.InvokeMember("onclick");
                        //HtmlElement btn_rewrite = td.GetElementsByTagName("tr")[1].InvokeMember("click");
                        webBrowser1.Update();
                        // MessageBox.Show(rt_tablecollection.Count.ToString()
                        //                 + ", " + tdcollection.Count.ToString());                                                      );                        
                        number_case++;
                        view_parameter();
                        break;

                    case "EDITPAGE":

                        tablecollection = webBrowser1.Document.GetElementsByTagName("table");
                        producttable = tablecollection[0];
                        tradetable = tablecollection[1];
                        productinputcollection = producttable.GetElementsByTagName("input");
                        tradeinputcollection = tradetable.GetElementsByTagName("td");
                        inputcollection = webBrowser1.Document.GetElementsByTagName("input");
                        btn_submit = inputcollection[97];



                        //物品名稱 確認是否有貨 沒貨可否預購
                        if (stocktable[id_updateproduct + 1, 3] == 0.ToString())
                        {
                            productinputcollection[0].InnerText = "預購中" + stocktable[id_updateproduct + 1, 2];
                        }
                        else
                        {
                            productinputcollection[0].InnerText = stocktable[id_updateproduct + 2, 2];
                        }
                        webBrowser1.Update();
                        //個人買場自訂分類                對於每個分類編輯到 exceltable 中
                        //商品價格及數量                  確認 exceltable 價錢及庫存
                        //賣家自用料號                    XXXXX
                        //刊登時間設定                    XXXXX
                        //物品新舊                        設定為全新品
                        //備貨狀態                        根據上面判定的狀態 有貨->備有現貨 沒貨->較長備貨
                        //影片                            XXXXX
                        //物品說明                        對於說明也寫入 exceltable        
                        //特別醒目標籤                    XXXXX
                        //物品所在地                      新北市
                        //買家下標限制                    XXXXX

                        //預設值                          XXXXX
                        //運費規定                        買價
                        //收款方式 PChomePay支付連        取決物品大小 exceltable 多設定一個狀態代表可否超商寄送 
                        //收款方式 其他收款方式           可 銀行或郵局轉帳 郵局無摺存款 貨到付款的部分要根據價錢判斷
                        //運送方式                        確認是否有上述超商寄送的狀態 郵局取決於大小
                        //收款及運送方式 組合說明         XXXXX
                        //合併運送運費                    一律勾選
                        //儲存預設值                      一律不勾

                        webBrowser1.Update();
                        number_case++;
                        break;

                    default:
                        // MessageBox.Show("No matching case");
                        number_case++;
                        view_parameter();
                        break;
                }
                //}

                //}
                loading = false;

            }
            else if (webBrowser1.Document.Title == "404 Not Found")
            {
                casestr = "404NOTFOUND";
                loading = false;
            }
            else
            {
                casestr = " ";
                loading = false;
            }

            //--count_navigated;
            // MessageBox.Show("Controls:"+webBrowser1.Controls.Count.ToString());
            // MessageBox.Show("count_navigated:" + count_navigated.ToString());
        }
        /* btn_getpicture_Click 功能說明
         * 
         * 
         */
        //[Obsolete]
        /*
         * 
         *
         */

        private void btn_getpicture_Click(object sender, EventArgs e)
        {            
            threadcatchpic.Start();
            
            /*
            string str_url;
            WebClient wc = new WebClient();
            byte[] bytes;
            int width;
            int height;
            int num_img = 0;
            int i;
            Bitmap bmp;
            for (int id_updateproduct = 0; id_updateproduct <= number_tomixproduct-1; id_updateproduct++)
            {
                
                Console.WriteLine("Start:" + (id_updateproduct + 1).ToString());
                Console.WriteLine("keyword:" + tomixtable[id_updateproduct + 1, 0]);
                str_url = @"https://www.tomytec.co.jp/tomix/products/n/" + tomixtable[id_updateproduct + 1, 0] + ".html";
                // string str_url = @"https://www.tomytec.co.jp/tomix/products/n/92373" + ".html";
                Console.WriteLine("url:" + str_url);
                 webBrowser1.Update();
                //  webnav(str_url).Wait();
              //  var t = Task.Run(() => webnav(str_url));
               // t.Wait(60000);
                webBrowser1.Navigate(str_url);
                if (threaddoc.ThreadState == ThreadState.Suspended)
                {
                    threaddoc.Resume();
                }
                else
                {
                    threaddoc.Start();
                }
                //Task.Run(() => threaddoc);
                //var t = Task.Run(() => threaddoc); 

                //Task.Delay(10000);
                //webBrowser1.Update();
                // Application.DoEvents();

                //Task.Delay(100000);

                //Thread.Sleep(5000);                     

                /*
                                while (casestr != "SUCCESS")
                                {
                                    //Console.WriteLine(casestr);
                                    Application.DoEvents();
                                    if (casestr == "404NOTFOUND")
                                    {
                                        Console.WriteLine("404NOTFOUND");
                                        break;
                                    }
                                    else
                                    {
                                        // webBrowser1.Update();
                                        Thread.Sleep(1);
                                        continue;
                                    }

                                }
                                if (casestr == "404NOTFOUND") continue;
                
                while (!picweb)
                {
                    Application.DoEvents();
                    //Console.WriteLine("Loading");
                    // Thread.Sleep(10);
                    continue;
                }
                webBrowser1.Update();
                picweb = false;
                
                threaddoc.Suspend();
                
                num_img = 0;

                HtmlElement detail = webBrowser1.Document.GetElementById("detail");
                HtmlElementCollection imgcollection = detail.GetElementsByTagName("img");

                Console.WriteLine("img:" + imgcollection.Count.ToString());

                for (i = 0; i < imgcollection.Count; i++)
                {
                    Console.WriteLine("src:" + imgcollection[i].GetAttribute("src"));
                    if(i == 0)Console.WriteLine("table:" + tomixtable[id_updateproduct + 1, 0]);
                    if (imgcollection[i].GetAttribute("src").ToString().Contains(tomixtable[id_updateproduct + 1, 0]))
                    {
                        string saveurl = imgcollection[i].GetAttribute("src").ToString();
                        bytes = wc.DownloadData(saveurl);
                        Console.WriteLine(saveurl);
                        Console.WriteLine(imgcollection[i].GetAttribute("width"));
                        Console.WriteLine(imgcollection[i].GetAttribute("height"));
                        width = int.Parse(imgcollection[i].GetAttribute("width").ToString());
                        height = int.Parse(imgcollection[i].GetAttribute("height").ToString());

                        bmp = new Bitmap(new MemoryStream(bytes));

                        bmp.Save(@"D:\阿立圓山\picture\Tomix\" + tomixtable[id_updateproduct + 1, 0] + "_" + num_img.ToString() + ".jpg");
                        num_img++;

                    }
                    else
                    {

                    }
                }
                Console.WriteLine("actual img:" + num_img.ToString());
                Console.WriteLine("FINISH:" + (id_updateproduct + 1).ToString());
                Console.WriteLine("");
                webBrowser1.Select();

            }
            Console.WriteLine("FINSIH PICTURE");
            */
        }
           
           
        private void view_parameter()
        {
            /*
            txtbx_data.Text += "navigated: " + call_navigated.ToString() + Environment.NewLine +
                           "navigating: " + call_navigating.ToString() + Environment.NewLine +
                           "completed: " + call_complete.ToString() + Environment.NewLine+Environment.NewLine;
            call_navigated = 0;
            call_navigating = 0;
            call_complete = 0;
            */
        }
        private void webBrowser1_Navigated(object sender, WebBrowserNavigatedEventArgs e)
        {
            while (loading)
            {
                Application.DoEvents();
            }
        }
        private void webBrowser1_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            /*  txtbx_data.Text = webBrowser1.DocumentText.ToString();
              var htmldocumet = (mshtml.HTMLDocument)webBrowser1.Document.DomDocument;
              string gethtml = htmldocumet.documentElement.outerHTML;

              txtbx_web.Text = gethtml;
            */
            /*
             * txtbx_data.Text += webBrowser1.Document.Url.ToString()+Environment.NewLine;
            txtbx_web.Text += e.Url.ToString() + Environment.NewLine;
            */
        }

        //method only for thread
        private void methodforproductexcel(object obj)
        {
            string path = txtbx_inputfile.Text;
            Excel.Application application = new Excel.Application();
            Excel.Workbook workbook = application.Workbooks.Open(path);
            Excel.Worksheet worksheet = workbook.Sheets[1];
            Excel.Range range = worksheet.UsedRange;

            rowcount = range.Rows.Count;
            columnCount = range.Columns.Count;

            exceltable = new string[rowcount, columnCount];

            try
            {

                for (int i = 1; i < rowcount; i++)
                {
                    for (int j = 1; j < columnCount; j++)
                    {
                        if (j == 1) //txtbx_data.Text += Environment.NewLine;
                            if (range.Cells[i, j] != null && range.Cells[i, j].Value2 != null)
                            {
                                exceltable[i, j] = range.Cells[i, j].Value2.ToString();

                            }
                            else
                            {
                                exceltable[i, j] = " ";

                            }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            // workbook.Save();
            workbook.Close();
            application.Quit();
            
            Console.WriteLine("FINISH 追蹤商品");
        }

        private void methodforstockexcel(object obj)
        {
            string path = txtbx_inputfilestock.Text;
            Excel.Application application = new Excel.Application();
            Excel.Workbook workbook = application.Workbooks.Open(path);
            Excel.Worksheet worksheet = workbook.Sheets[1];
            Excel.Range range = worksheet.UsedRange;

            rowcount_stock = range.Rows.Count;
            columncount_stock = range.Columns.Count;
            // Console.WriteLine(rowcount_stock.ToString()+","+ columncount_stock.ToString());
            stocktable = new string[rowcount_stock, columncount_stock];

            for (int i = 1; i < rowcount_stock; i++)
            {
                for (int j = 1; j < columncount_stock; j++)
                {
                    if (range.Cells[i, j] != null && range.Cells[i, j].Value2 != null)
                    {
                        stocktable[i, j] = range.Cells[i, j].Value2.ToString();
                        //   Console.WriteLine(stocktable[i, j].ToString());
                    }
                    else
                    {
                        stocktable[i, j] = " ";
                        //   Console.WriteLine(stocktable[i, j].ToString());

                    }
                }
            }
            //  workbook.Save();
            workbook.Close();
            application.Quit();
            Console.WriteLine("FINISH 庫存商品");
        }
 
        private void methodforpic(object obj)
        {
            string str_url;
            WebClient wc = new WebClient();
            byte[] bytes;
            int width;
            int height;
            int num_img = 0;
            int i;
            Bitmap bmp;
           // WebBrowser webBrowser2 = new WebBrowser();
          //  webBrowser2.DocumentCompleted += new WebBrowserDocumentCompletedEventHandler(webdoccomplete);
            while (true)
            {

                for (int id_updateproduct = 0; id_updateproduct <= number_tomixproduct - 1; id_updateproduct++)
                {

                    Console.WriteLine("Start:" + (id_updateproduct + 1).ToString());
                    Console.WriteLine("keyword:" + tomixtable[id_updateproduct + 1, 0]);
                    str_url = @"https://www.tomytec.co.jp/tomix/products/n/" + tomixtable[id_updateproduct + 1, 0] + ".html";
                    Console.WriteLine("url:" + str_url);
                    webBrowser1.Navigate(str_url);


                    Task.Delay(10000);
                    //webBrowser1.Update();
                    Application.DoEvents();

                    // Task.Wait(100000);

                    //Thread.Sleep(5000);                     


                    while (casestr != "SUCCESS")
                    {
                        //Console.WriteLine(casestr);
                        Application.DoEvents();
                        if (casestr == "404NOTFOUND")
                        {
                            Console.WriteLine("404NOTFOUND");
                            break;
                        }
                        else
                        {
                            // webBrowser1.Update();
                            Thread.Sleep(1);
                            continue;
                        }

                    }
                    if (casestr == "404NOTFOUND") continue;


                    //webBrowser1.Update();



                    num_img = 0;


                    //HtmlElement detail = webBrowser1.Document.GetElementById("detail");
                    HtmlElement detail;
                    HtmlElementCollection imgcollection;
                    this.Invoke(new Action(() =>
                    {
                        detail = webBrowser1.Document.GetElementById("detail");
                        imgcollection = detail.GetElementsByTagName("img");
                        Console.WriteLine("img:" + imgcollection.Count.ToString());
                    
                   

                    for (i = 0; i < imgcollection.Count; i++)
                    {
                        Console.WriteLine("src:" + imgcollection[i].GetAttribute("src"));
                        if (i == 0) Console.WriteLine("table:" + tomixtable[id_updateproduct + 1, 0]);
                        if (imgcollection[i].GetAttribute("src").ToString().Contains(tomixtable[id_updateproduct + 1, 0]))
                        {
                            string saveurl = imgcollection[i].GetAttribute("src").ToString();
                            bytes = wc.DownloadData(saveurl);
                            Console.WriteLine(saveurl);
                            Console.WriteLine(imgcollection[i].GetAttribute("width"));
                            Console.WriteLine(imgcollection[i].GetAttribute("height"));
                            width = int.Parse(imgcollection[i].GetAttribute("width").ToString());
                            height = int.Parse(imgcollection[i].GetAttribute("height").ToString());

                            bmp = new Bitmap(new MemoryStream(bytes));

                            bmp.Save(@"D:\阿立圓山\picture\Tomix\" + tomixtable[id_updateproduct + 1, 0] + "_" + num_img.ToString() + ".jpg");
                            num_img++;

                        }
                        else
                        {

                        }
                    }
                    Console.WriteLine("actual img:" + num_img.ToString());
                    Console.WriteLine("FINISH:" + (id_updateproduct + 1).ToString());
                    Console.WriteLine("");
                        // webBrowser1.Select();
                    }));
               
                }
                Console.WriteLine("FINSIH PICTURE" + ", id:" + (id_updateproduct + 1).ToString() + ", total:" + number_tomixproduct.ToString());
                //Console.WriteLine("FINSIH PICTURE");
                Thread.Sleep(1);
            }
        }

        private void methodfordocstatus(object obj)
        {
            while (true)
            {
                if (casestr == "SUCCESS")
                {
                    picweb = true;
                    continue;
                }
                else
                {
                    picweb = false;
                    continue;
                }
            }
        }
        private async Task webnav(string url)
        {
           // webBrowser1.Navigate(url);
           // var task_status;


            //return task_status;
            //webBrowser1.Update();
        }


    } 
}

