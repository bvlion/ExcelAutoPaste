using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace AutoPaste
{
    class Program
    {
        static void Main(string[] args)
        {
            // 同ディレクトリにpasted.xlsxを作成する
            String[] fullPaths = System.Reflection.Assembly.GetExecutingAssembly().Location.Split('\\');
            String filePath = null;
            for (int i = 0; i < fullPaths.Length - 1; i++)
            {
                filePath += fullPaths[i] + @"\";
            }
            String fileName = filePath + "pasted.xlsx";

            // 存在した場合の処理
            if (File.Exists(fileName))
            {
                //メッセージボックスを表示する
                DialogResult result = MessageBox.Show("ファイルを上書きしますか？",
                    "既にファイルが存在します!!",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Exclamation,
                    MessageBoxDefaultButton.Button2);

                if (result == DialogResult.Yes)
                {
                    File.Delete(fileName);
                }
                else
                {
                    MessageBox.Show("出力は取り消されました。", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
            }

            // 画像一覧取得
            String[] files = Directory.GetFiles(filePath);
            Array.Sort(files);

            //Excelオブジェクトの初期化
            Excel.Application ExcelApp = null;
            Excel.Workbooks wbs = null;
            Excel.Workbook wb = null;
            Excel.Sheets shs = null;
            Excel.Worksheet ws = null;
            Excel.Shapes shps1 = null;
            Excel.Shape shp1 = null;
            Excel.Range paste1Range = null;
            Excel.Range paste1Cells = null;



            try
            {
                //Excelシートのインスタンスを作る
                ExcelApp = new Excel.Application();
                wbs = ExcelApp.Workbooks;
                wb = wbs.Add();

                int index = 1;
                bool isNextIndex = false;
                foreach (String file in files)
                {
                    if (!(file.EndsWith("png") || file.EndsWith("PNG")))
                    {
                        continue;
                    }
                    String[] pngNames = file.Split('\\');
                    String pngName = pngNames[pngNames.Length - 1].Split('.')[0].Split('_')[0];

                    try
                    {
                        shs = wb.Sheets;
                        try
                        {
                            ws = shs[index];
                        }
                        catch (COMException)
                        {
                            ws = shs.Add();
                            ws.Move(Type.Missing, shs[index]);
                        }
                        ws.Name = pngName;
                        ws.Select(Type.Missing);

                        ExcelApp.Visible = false;

                        shps1 = ws.Shapes;
                        if (!isNextIndex)
                        {
                            shp1 = shps1.AddPicture(file, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue, 1, 14.5f, 0, 0);
                        }
                        else
                        {
                            shp1 = shps1.AddPicture(file, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue, 51.4f * 20, 14.5f, 0, 0);
                        }
                        shp1.ScaleHeight(1, Microsoft.Office.Core.MsoTriState.msoTrue);
                        shp1.ScaleWidth(1, Microsoft.Office.Core.MsoTriState.msoTrue);
                        // セルへ貼り付け
                        paste1Cells = ws.Cells;
                        if (!isNextIndex)
                        {
                            paste1Range = paste1Cells[1, 1] as Microsoft.Office.Interop.Excel.Range;
                            paste1Range.Value = "入力状態";
                        }
                        else
                        {
                            paste1Range = paste1Cells[1, 20] as Microsoft.Office.Interop.Excel.Range;
                            paste1Range.Value = "確認ボタン押下後";
                        }
                        paste1Range.Font.Bold = true;
                    }
                    finally
                    {
                        // Excelのオブジェクトはループごとに開放
                        Marshal.ReleaseComObject(shs);
                        Marshal.ReleaseComObject(shps1);
                        Marshal.ReleaseComObject(shp1);
                        Marshal.ReleaseComObject(paste1Cells);
            	        Marshal.ReleaseComObject(paste1Range);
                        shs = null;
                        shps1 = null;
                        shp1 = null;
                        paste1Cells = null;
                        paste1Range = null;
                    }
                    if (!isNextIndex)
                    {
                        isNextIndex = true;
                    } else {
                        isNextIndex = false;
                        index++;
                    }
                }

                //excelファイルの保存
                wb.SaveAs(fileName);
                wb.Close(true);
                ExcelApp.Quit();
            }
            finally
            {
                //Excelのオブジェクトを開放
                if (ws != null)
                {
                    Marshal.ReleaseComObject(ws);
                }
                Marshal.ReleaseComObject(wb);
                Marshal.ReleaseComObject(wbs);
                Marshal.ReleaseComObject(ExcelApp);
                ws = null;
                shs = null;
                wb = null;
                wbs = null;
                ExcelApp = null;

                GC.Collect();
            }
            MessageBox.Show("出力処理が完了致しました！", "完了！", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
