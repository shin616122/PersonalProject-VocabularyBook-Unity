using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;

public static class LoadData
{
    /// <summary>
    /// エクセルファイルからデータを読み込み
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    public static List<RowData> LoadDataFromExcel(string filePath, bool isReviewOnlyChcked, bool IsQuestionHidden)
    {
        try
        {
            // パスのエクセルを読み込む
            var workbook = new XLWorkbook(filePath);
            // 1行目はヘッダーの想定で、それをスキップして2列目かつデータを入ってるセルのみ読み込む。
            var rows = workbook.Worksheet(1).RangeUsed().RowsUsed().Skip(1);

            // エラーチェック
            if (!rows.Any())
            {
                //　Nullチェック
                // MessageBox.Show("エラー: エクセルファイルにデータがありません。",
                //                 "アドミンに連絡してください！ ",
                //                 MessageBoxButtons.OK,
                //                 MessageBoxIcon.Exclamation);

                // Application.Exit();
            }
            else
            {

                // フォマードチェック
                //MessageBox.Show("アドミンに連絡してください！ エクセルのフォマードが対応していません。",
                //                "残念でした",
                //                MessageBoxButtons.OK,
                //                MessageBoxIcon.Exclamation);

                //Environment.Exit(1);
            }

            // RowDataのリストを作る
            var rowDataList = new List<RowData>();

            // エラーリスt
            var errorList = new List<int>();

            // rowsのデータをRowDataに突っ込む
            foreach (var row in rows)
            {
                var rowData = new RowData();

                // TODO 数式も読めるように書き直す
                int rowNumber = row.RowNumber();
                string id = row.Cell("A").GetString();
                string question = row.Cell("B").GetString();
                string answer = row.Cell("C").GetString();
                string shouldReview = row.Cell("D").GetString();

                // 空セルチェック
                if (string.IsNullOrEmpty(id) ||
                   string.IsNullOrEmpty(question) ||
                   string.IsNullOrEmpty(answer) ||
                   string.IsNullOrEmpty(shouldReview))
                {
                    errorList.Add(rowNumber);
                }

                rowData.RowNumber = rowNumber;
                if (Int32.TryParse(id, out int idInt))
                {
                    rowData.Id = idInt;
                }
                else
                {
                    // MessageBox.Show($"通番は数字のみ　行: {rowNumber}",
                    //                 "アドミンに連絡してください！",
                    //                 MessageBoxButtons.OK,
                    //                 MessageBoxIcon.Exclamation);
                    // 強制終了
                    // Application.Exit();
                }

                // TODO 綺麗に書く。　無理やりフラグによってデータを入れ替えて返してる
                if (IsQuestionHidden)
                {
                    rowData.Question = answer;
                    rowData.Answer = question;
                }
                else
                {
                    rowData.Question = question;
                    rowData.Answer = answer;
                }

                if (Int32.TryParse(shouldReview, out int shouldReviewInt))
                {
                    rowData.ShouldReview = shouldReviewInt;
                }
                else
                {
                    // MessageBox.Show($"フラグ（印）は0か1のみ　行: {rowNumber}",
                    //                 "アドミンに連絡してください！",
                    //                 MessageBoxButtons.OK,
                    //                 MessageBoxIcon.Exclamation);
                    // 強制終了
                    // Application.Exit();
                }

                // 突っ込んだRowDataをリストに突っ込む
                rowDataList.Add(rowData);
            }

            // エラーリストにエラーがあったら、メッセージ出して強制終了
            if (errorList.Any())
            {
                // MessageBox.Show("空セル - 行: " + String.Join(", ", errorList),
                //                 "アドミンに連絡してください！",
                //                 MessageBoxButtons.OK,
                //                 MessageBoxIcon.Exclamation);
                // 強制終了
                // Application.Exit();
            }

            // 印の問題のみにフィルタリング
            if (isReviewOnlyChcked)
            {
                rowDataList = rowDataList.Where(row => row.ShouldReview == 1).ToList();
            }


            return rowDataList;
        }
        catch (Exception ex)
        {
            // MessageBox.Show("エラー: " + ex.Message,
            //                 "アドミンに連絡してください！",
            //                 MessageBoxButtons.OK,
            //                 MessageBoxIcon.Exclamation);
            // 強制終了
            // Application.Exit();

            throw ex;
        }
    }
}

/// <summary>
/// エクセルから読み込んだモデル
/// </summary>
public class RowData
{
    public int RowNumber { get; set; }
    public int Id { get; set; }
    public string Question { get; set; }
    public string Answer { get; set; }
    public int ShouldReview { get; set; }
}
