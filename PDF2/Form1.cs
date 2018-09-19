using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using iTextSharp.text;
using iTextSharp.text.pdf;


namespace PDF2
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        
        //dataGridViewにファイルをドロップしたときの処理
        private void dataGridView1_DragDrop(object sender, DragEventArgs e)
        {
            int cnt = dataGridView1.Rows.Count;//行追加前の全行数
            string[] filename = (string[])e.Data.GetData(DataFormats.FileDrop, false);                          //ドロップしたファイルのパスを配列にする
            foreach (string str in filename)
            {                
                if (System.IO.Path.GetExtension(str) == ".pdf" || System.IO.Path.GetExtension(str) == ".PDF")   //PDFファイルだけdataGridViewに追加する
                {
                    dataGridView1.Rows.Add();
                    dataGridView1.Rows[cnt].Cells[0].Value = System.IO.Path.GetFileName(str);                   //ファイル名
                    dataGridView1.Rows[cnt].Cells[1].Value = str;                                               //フルパス
                    cnt++;
                }
            }
        }

        private void dataGridView1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }                
        }

        //「上へ」ボタン
        private void button1_Click(object sender, EventArgs e)
        {
            int rowind = dataGridView1.SelectedCells[0].RowIndex;       //選択した行のIndex
            int rowcnt = dataGridView1.Rows.Count;                      //全行数

            if (rowind == 0) return;                                    //一番上の行のときは何もしない

            DataGridViewRow dgvr = dataGridView1.Rows[rowind];
            dataGridView1.Rows.RemoveAt(rowind);

            if(rowcnt == rowind)
            {
                dataGridView1.Rows.Insert(rowind, dgvr);                //一番下の行だった時だけズレるので
            }
            else
            {
                dataGridView1.Rows.Insert(rowind - 1, dgvr);
            }
            dataGridView1.Rows[rowind - 1].Cells[0].Selected = true;    //挿入した行を選択する
        }


        //「下へ」ボタン
        private void button2_Click(object sender, EventArgs e)
        {
            int rowind = dataGridView1.SelectedCells[0].RowIndex;   //選択した行のIndex
            int rowcnt = dataGridView1.Rows.Count;                  //全行数

            if (rowind+1 == rowcnt) return;                         //一番下の行のときは何もしない

            DataGridViewRow dgvr = dataGridView1.Rows[rowind];      //選択した行のコピーをとっておく
            dataGridView1.Rows.RemoveAt(rowind);                    //選択した行を消す
            dataGridView1.Rows.Insert(rowind + 1, dgvr);            //コピーしておいた行を挿入
            dataGridView1.Rows[rowind + 1].Cells[0].Selected = true;//挿入した行を選択する（これがないと連続で移動するとき不便）
        }


        //「結合」ボタン
        private void button3_Click(object sender, EventArgs e)
        {
            //結合するPDFファイルのパスのリストを作成
            List<string> path_list = new List<string>();
            foreach (object row in dataGridView1.Rows)
            {
                DataGridViewRow hoge = (DataGridViewRow)row;
                path_list.Add(hoge.Cells[1].Value.ToString());
            }

            //保存先を選択して､結合する
            //https://dobon.net/vb/dotnet/form/savefiledialog.html
            //https://zero0nine.com/archives/2866
            SaveFileDialog dialog = new SaveFileDialog();
            dialog.Filter = "PDFファイル(*.PDF)|*.PDF";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                FileStream fs = new FileStream(dialog.FileName, FileMode.Create);   //保存先のFileStreamを取得する
                Document doc = new Document();                                      //Documentのインスタンス
                PdfCopy pdfcpy = new PdfCopy(doc, fs);                              //PdfCopyのインスタンス：pdfcopyの中のdoc(文書本体？)にPDFファイルを継ぎ足していく
                doc.Open();                                                         //OpenメソッドでDocumentを開いて内容編集可能にする
                foreach (string str in path_list)                                   //結合するPDFファイル分だけループを回す
                {
                    PdfReader reader = new PdfReader(str);                          //PDFReader:外部のPDFファイルを読み込む
                    pdfcpy.AddDocument(reader);                                     //読み込んだPDFをpdfcopyに追加する
                    reader.Close();                                                 //readerは都度閉じる
                }
                pdfcpy.Close();                                                     //closeメソッドで､Filestreamがflushされて結合されたPDFファイルが完成
            }
        }   
    }
}
