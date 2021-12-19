using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using Spire.Xls;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace word
{
    public partial class Form1 : Form
    {

        //文件数量
        private int Count = 0;
        //扫描到的文件路径list
        private List<String> filelist = new List<String>();
        //读取的word文本内容
        private StringBuilder filecontent;
        //要处理的文件名称及路径
        private String handledfile = "";
        //要处理的文件备份
        private String handledfileBackup = "";
        //要扫扫描的文件夹
        private String scanfilepath = "";
        //要统计的文件的次数
        private int WordCount = 0;
        //要统计的第一阶段的词数
        private int WordCount1 = 0;
        //要统计的第二阶段的词数
        private int WordCount2 = 0;
        //要统计的第三阶段的词数
        private int WordCount3 = 0;

        struct aWord
        {
            public string word;
            public int Cisu;
        }
        List<aWord> words;

        public Form1()
        {
            InitializeComponent();
        }

        //选择扫描文件夹按钮
        private void gb1_button1_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.Description = "请选择文件路径";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                //clear
                richTextBox1.Text = "";

                scanfilepath = (dialog.SelectedPath).ToString();
                //扫描文件夹
                filelist = new List<String>();
                ScanPath(scanfilepath);
                //将结果显示在richtextbox
                foreach (String filepath in filelist)
                {
                    richTextBox1.Text += filepath + "\r\n";
                }
            }
        }

        //更新词频按钮
        private void gb1_button2_Click(object sender, EventArgs e)
        {
            if (scanfilepath == "")
            {
                MessageBox.Show("请选择扫描文件夹");
                return;
            }


            gb1_button2.Text = "处理中";
            gb1_button2.Enabled = false;

            //clear
            richTextBox2.Text = "";

            //设置进度条最大值
            gb1_progressBar1.Maximum = 2;

            //清空读取内容
            filecontent = new StringBuilder();

            //逐个处理word
            HandleWord();
            gb1_progressBar1.Value = 1;

            //统计词频
            StatisticsWords2();

        }

        //递归扫描指定文件夹下的所有文件并过滤
        public void ScanPath(string path)
        {
            DirectoryInfo root = new DirectoryInfo(path);

            foreach (FileInfo f in root.GetFiles())
            {
                if (Path.GetExtension(f.FullName) == ".docx" || Path.GetExtension(f.FullName) == ".doc")
                {
                    filelist.Add(f.FullName);
                    Count += 1;
                }
            }
            foreach (DirectoryInfo d in root.GetDirectories())
            {
                ScanPath(d.FullName);
            }
        }

        //处理word文档
        public void HandleWord()
        {
            if (filelist == null || filelist.Count() == 0) return;

            Document doc = new Document();

            foreach (String filepath in filelist)
            {
                doc.LoadFromFile(filepath);

                //使用GetText方法获取文档中的所有文本
                filecontent.Append(doc.GetText());
            }
        }

        //旧的词频统计方法
        public void StatisticsWords(string path)
        {
            if (!File.Exists(path))
            {
                Console.WriteLine("文件不存在！");
                return;
            }

            Hashtable ht = new Hashtable(StringComparer.OrdinalIgnoreCase);
            StreamReader sr = new StreamReader(path, System.Text.Encoding.UTF8);
            string line = sr.ReadLine();

            string[] wordArr = null;
            int num = 0;
            while (line.Length > 0)
            {
                //   MatchCollection mc =  Regex.Matches(line, @"\b[a-z]+", RegexOptions.Compiled | RegexOptions.IgnoreCase);
                //foreach (Match m in mc)
                //{
                //    if (ht.ContainsKey(m.Value))
                //    {
                //        num = Convert.ToInt32(ht[m.Value]) + 1;
                //        ht[m.Value] = num;
                //    }
                //    else
                //    {
                //        ht.Add(m.Value, 1);
                //    }
                //}
                //line = sr.ReadLine();

                wordArr = line.Split(' ');
                foreach (string s in wordArr)
                {
                    if (s.Length == 0)
                        continue;
                    //去除标点
                    line = Regex.Replace(line, @"[\p{P}*]", "", RegexOptions.Compiled);
                    //将单词加入哈希表
                    if (ht.ContainsKey(s))
                    {
                        num = Convert.ToInt32(ht[s]) + 1;
                        ht[s] = num;
                    }
                    else
                    {
                        ht.Add(s, 1);
                    }
                }
                line = sr.ReadLine();
            }

            ArrayList keysList = new ArrayList(ht.Keys);
            //对Hashtable中的Keys按字母序排列
            keysList.Sort();
            //按次数进行插入排序【稳定排序】，所以相同次数的单词依旧是字母序
            string tmp = String.Empty;
            int valueTmp = 0;
            for (int i = 1; i < keysList.Count; i++)
            {
                tmp = keysList[i].ToString();
                valueTmp = (int)ht[keysList[i]];//次数
                int j = i;
                while (j > 0 && valueTmp > (int)ht[keysList[j - 1]])
                {
                    keysList[j] = keysList[j - 1];
                    j--;
                }
                keysList[j] = tmp;//j=0
            }
            //打印出来
            foreach (object item in keysList)
            {
                Console.WriteLine(item + " <------------> " + ht[item]);
            }
        }

        //新的词频统计方法
        public void StatisticsWords2()
        {
            String str = (filecontent.ToString()).ToLower();
            //去除音标
            str = Regex.Replace(str, @"\[.*\]", " ");
            //其它处理
            str = Regex.Replace(str, "[^a-z A-Z]", " ");
            str = Regex.Replace(str, "\\s{2,}", " ");
            string[] s = str.Split(' ');
            words = new List<aWord>();
            foreach (string sf in s)
            {
                if (words.Exists(a => a.word == sf))
                {
                    int u = words.FindIndex(a => a.word == sf);
                    aWord w = words[u];
                    w.word = sf;
                    w.Cisu += 1;
                    words[u] = w;
                }
                else
                {
                    aWord word = new aWord();
                    word.word = sf;
                    word.Cisu = 1;
                    words.Add(word);
                }
            }

            List<aWord> NewWords = words.OrderByDescending(a => a.Cisu).ToList();

            //初始化
            Workbook workbook = new Workbook();
            Worksheet sheet1 = workbook.Worksheets[0];

            //编辑表头
            //sheet1.Range["A1"].Text = "序号";
            //sheet1.Range["B1"].Text = "单词";
            //sheet1.Range["C1"].Text = "词频";

            int sp = 1;
            foreach (aWord t in NewWords)
            {
                if (t.word.Trim() == "") continue;

                //写入Excel
                sheet1.Range["A" + sp].Text = sp.ToString();
                sheet1.Range["B" + sp].Text = t.word;
                sheet1.Range["C" + sp].Text = t.Cisu.ToString();
                sp++;
            }

            //保存文件
            workbook.SaveToFile(Environment.CurrentDirectory + "\\statistics.xlsx", Spire.Xls.FileFormat.Version2013);



            if (radioButton1.Checked == true)
            {
                gb1_button2.Enabled = true;
                gb1_button2.Text = "内容加载中";
                gb1_button2.Enabled = false;

                //到此数据已经处理完成 下面将数据加载到文本框
                sp = 1;
                foreach (aWord t in NewWords)
                {
                    if (t.word.Trim() == "") continue;

                    //显示在richTextbox
                    richTextBox2.Text += sp.ToString() + " ";
                    richTextBox2.Text += t.word + " ";
                    richTextBox2.Text += t.Cisu.ToString() + " ";
                    richTextBox2.Text += "\r\n";

                    sp++;
                }

                //进度条结束
                gb1_progressBar1.Value = 2;

                gb1_button2.Enabled = true;
                gb1_button2.Text = "更新词频";
            }
            else
            {
                //进度条达到最大值
                gb1_progressBar1.Value = 2;
                //恢复按钮
                gb1_button2.Enabled = true;
                gb1_button2.Text = "更新词频";
            }
        }

        //选择被处理文件
        private void gb2_button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "docx文件|*.docx|doc文件|*.doc";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                handledfile = dialog.FileName;

                //读取word内容到文本框
                Document doc = new Document();
                doc.LoadFromFile(handledfile);
                richTextBox3.Text = doc.GetText();
            }
        }

        //开始处理文件
        private void gb2_button2_Click(object sender, EventArgs e)
        {
            if (handledfile == "")
            {
                MessageBox.Show("请选择要处理的文件");
                return;
            }

            if (!File.Exists(Environment.CurrentDirectory + "\\statistics.xlsx"))
            {
                MessageBox.Show("词频文件不存在 请先进行词频统计");
                return;
            }

            //判断输入的数字是否合理
            if (textBox1.Text.Trim() == "" || textBox2.Text.Trim() == "" || textBox3.Text.Trim() == "" || textBox4.Text.Trim() == "" || textBox5.Text.Trim() == "" || textBox6.Text.Trim() == "")
            {
                MessageBox.Show("数值不能为空");
                return;
            }
            int begin1 = int.Parse(textBox1.Text.Trim());
            int end1 = int.Parse(textBox2.Text.Trim());
            int begin2 = int.Parse(textBox4.Text.Trim());
            int end2 = int.Parse(textBox3.Text.Trim());
            int begin3 = int.Parse(textBox6.Text.Trim());
            int end3 = int.Parse(textBox5.Text.Trim());

            if (begin1 < 1)
            {
                MessageBox.Show("第一阶段的开始值不能小于1");
                return;
            }

            if (begin1 >= end1 || begin2 >= end2 || begin3 >= end3)
            {
                MessageBox.Show("每个阶段的开始值不能大于或等于结束值");
                return;
            }
            if (end1 >= begin2 || end2 >= begin3)
            {
                MessageBox.Show("上个阶段的结束值不能大于或等于下个阶段的开始值");
                return;
            }

            gb2_button2.Text = "处理中...";
            gb2_button2.Enabled = false;

            //进度条置0
            gb2_progressBar1.Value = 0;

            //备份文件
            handledfileBackup = Path.GetDirectoryName(handledfile) + @"\" + Path.GetFileNameWithoutExtension(handledfile) + "_backup" + Path.GetExtension(handledfile);
            Document backdoc = new Document();
            backdoc.LoadFromFile(handledfile);
            backdoc.SaveToFile(handledfileBackup);

            //创建Workbook对象
            Workbook wb = new Workbook();
            //加载Excel文档
            wb.LoadFromFile(Environment.CurrentDirectory + "\\statistics.xlsx");
            //获取第一个工作表
            Worksheet sheet = wb.Worksheets[0];
            String index, word, count;

            //新建一个word文档对象并加载文档
            Document document = new Document();
            document.LoadFromFile(handledfile, Spire.Doc.FileFormat.Docx2010);

            gb2_progressBar1.Maximum = 4;

            //词频1-500设置颜色
            WordCount1 = 0;
            for (int i = begin1; i <= end1; i++)
            {
                //获取单元格值
                index = sheet.Range["A" + i].Value2.ToString();
                word = sheet.Range["B" + i].Value2.ToString();
                count = sheet.Range["C" + i].Value2.ToString();

                //判断数量
                if (index == "") break;

                //判断内容是否为空
                if (word == "") continue;

                //查找文档中所有符合条件的字符串
                TextSelection[] text = document.FindAllString(word, false, true);
                if (text == null) continue;

                //设置高亮颜色
                foreach (TextSelection seletion in text)
                {
                    //seletion.GetAsOneRange().CharacterFormat.HighlightColor = Color.Green;
                    seletion.GetAsOneRange().CharacterFormat.TextColor = Color.Blue;
                }

                WordCount1++;
            }
            gb2_progressBar1.Value = 1;

            //词频500-1500设置颜色
            WordCount2 = 0;
            for (int i = begin2; i <= end2; i++)
            {
                //获取单元格值
                index = sheet.Range["A" + i].Value2.ToString();
                word = sheet.Range["B" + i].Value2.ToString();
                count = sheet.Range["C" + i].Value2.ToString();

                //判断数量
                if (index == "") break;

                //判断内容是否为空
                if (word == "") continue;

                //查找文档中所有符合条件的字符串
                TextSelection[] text = document.FindAllString(word, false, true);
                if (text == null) continue;

                //设置高亮颜色
                foreach (TextSelection seletion in text)
                {
                    //seletion.GetAsOneRange().CharacterFormat.HighlightColor = Color.Blue;
                    seletion.GetAsOneRange().CharacterFormat.TextColor = Color.Green;
                }

                WordCount2++;
            }
            gb2_progressBar1.Value = 2;

            //词频1500以上设置颜色
            WordCount3 = 0;
            for (int i = begin3; i <= end3; i++)
            {
                //获取单元格值
                index = sheet.Range["A" + i].Value2.ToString();
                word = sheet.Range["B" + i].Value2.ToString();
                count = sheet.Range["C" + i].Value2.ToString();
                //判断数量
                if (index == "") break;
                //判断内容是否为空
                if (word == "") continue;
                //查找文档中所有符合条件的字符串
                TextSelection[] text = document.FindAllString(word, false, true);
                if (text == null) continue;
                //设置高亮颜色
                foreach (TextSelection seletion in text)
                {
                    //seletion.GetAsOneRange().CharacterFormat.HighlightColor = Color.Red;//
                    seletion.GetAsOneRange().CharacterFormat.TextColor = Color.Red;
                }

                WordCount3++;
            }
            gb2_progressBar1.Value = 3;

            //保存文档
            document.SaveToFile(handledfile, Spire.Doc.FileFormat.Docx2010);

            //统计信息
            StatisticsWordInfo();

            label_word_count.Text = WordCount.ToString();
            label_word_count1.Text = WordCount1.ToString();
            label_word_count2.Text = WordCount2.ToString();
            label_word_count3.Text = WordCount3.ToString();

            gb2_progressBar1.Value = 4;

            gb2_button2.Enabled = true;
            gb2_button2.Text = "开始处理";
        }

        //统计word文件的词数
        public void StatisticsWordInfo()
        {
            String str = richTextBox3.Text.ToLower();
            //去除音标
            str = Regex.Replace(str, @"\[.*\]", " ");
            //其它处理
            str = Regex.Replace(str, "[^a-z A-Z]", " ");
            str = Regex.Replace(str, "\\s{2,}", " ");
            string[] s = str.Split(' ');
            words = new List<aWord>();
            foreach (string sf in s)
            {
                if (words.Exists(a => a.word == sf))
                {
                    int u = words.FindIndex(a => a.word == sf);
                    aWord w = words[u];
                    w.word = sf;
                    w.Cisu += 1;
                    words[u] = w;
                }
                else
                {
                    aWord word = new aWord();
                    word.word = sf;
                    word.Cisu = 1;
                    words.Add(word);
                }
            }

            List<aWord> NewWords = words.OrderByDescending(a => a.Cisu).ToList();

            WordCount = 1;
            foreach (aWord t in NewWords)
            {
                if (t.word.Trim() == "") continue;
                WordCount++;
            }
        }

        //查看词频统计结果
        private void button1_Click(object sender, EventArgs e)
        {
            if (File.Exists(Environment.CurrentDirectory + "\\statistics.xlsx"))
            {
                System.Diagnostics.Process p = new System.Diagnostics.Process();

                p.StartInfo.UseShellExecute = true;

                p.StartInfo.FileName = Environment.CurrentDirectory + "\\statistics.xlsx";

                p.Start();
            }
            else
            {
                MessageBox.Show("文件未生成");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (File.Exists(handledfile))
            {
                System.Diagnostics.Process p = new System.Diagnostics.Process();

                p.StartInfo.UseShellExecute = true;

                p.StartInfo.FileName = handledfile;

                p.Start();
            }
            else
            {
                MessageBox.Show("文件不存在");
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            radioButton2.Checked = true;
        }
    }
}
