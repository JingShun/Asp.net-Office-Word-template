using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class tests : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {

    }

    private void DownLoadFile(string parFilePath)
    {
        //將虛擬路徑轉換成實體路徑
        string FilePath = Server.MapPath(parFilePath);

        if (FilePath.Split('\\').Length != 0)
        {
            string FileName = FilePath.Split('\\')[FilePath.Split('\\').Length - 1];

            //中文檔名作轉換
            FileName = HttpUtility.UrlEncode(FileName, Encoding.UTF8);

            FileStream fr = new FileStream(FilePath, FileMode.Open);
            Byte[] buf = new Byte[fr.Length];

            fr.Read(buf, 0, Convert.ToInt32(fr.Length));
            fr.Close();
            fr.Dispose();

            Response.Clear();
            Response.ClearHeaders();
            Response.Buffer = true;
            //轉換文字檔編碼格式用，但本次輸出無文字檔，故註解此段
            //Response.ContentEncoding = parEncoding;
            Response.AddHeader("content-disposition", "attachment; filename=" + FileName);

            Response.BinaryWrite(buf);
            Response.End();
        }
    }

    protected void Button1_Click(object sender, EventArgs e)
    {
        DateTime time = DateTime.Now;

        Dictionary<string, string> dct = new Dictionary<string, string>();
        dct.Add("學年", "105");
        dct.Add("學期", "2");
        dct.Add("社團", "程式開發社");
        dct.Add("地點", "台南市善化區大成路195號");
        dct.Add("日期", (time.Year-1911)+"年"+time.Month+"月"+time.Day+"日");
        dct.Add("老師", "張老師");

        for (int i = 1; i <= 28; i++)
        {
            dct.Add("班級" + i, "201");
            dct.Add("座號" + i, i.ToString());
            dct.Add("學號" + i, "510009");
            dct.Add("姓名" + i, "王亭媗");
        }

        string FilePath = Server.MapPath("/文件/點名單_樣板.docx");

        Class_myDocx.msg = "";

        Byte[] buf = Class_myDocx.MakeDocx(FilePath, dct,Regex:true);

        Trace.Warn(Class_myDocx.msg);
        

        ////輸出
        //Response.Clear();
        //Response.ClearHeaders();
        //Response.Buffer = true;
        ////轉換文字檔編碼格式用，但本次輸出無文字檔，故註解此段
        ////Response.ContentEncoding = parEncoding;
        //Response.AddHeader("content-disposition", "attachment; filename=" + "測試5.docx");

        //Response.BinaryWrite(buf);
        //Response.End();
    }

    protected void Button2_Click(object sender, EventArgs e)
    {
        //綴詞補上Regex的脫逸字元
        System.Text.RegularExpressions.Regex 綴詞Reg = new System.Text.RegularExpressions.Regex("(.)");
        TextBox1.Text = 綴詞Reg.Replace(TextBox1.Text, ".*?\\$1");


    }
}