using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using UnityEngine;
using UnityEngine.UI;

public class CanvasManager : MonoBehaviour
{
    public Text btnText;


    // Unity Object
    public GameObject MainMenu;
    public GameObject QuestionPanel;
    public Text filePathText;
    public Text questionNumberText;
    public Toggle hideQuestionToggle;
    public Toggle shuffleToggle;
    public Toggle reviewOnlyToggle;


    private List<RowData> RowDataList { get; set; }
    private int CurrentRowNumber { get; set; }
    public string FilePath { get; set; }
    public bool IsShuffleChecked { get; set; }
    public bool IsReviewOnlyChcked { get; set; }
    public bool IsQuestionHidden { get; set; }
    public int QuestionNumber { get; set; } = 1;


    /// <summary>
    /// 回答ボタン
    /// </summary>
    public void OnAnswerButtonClick()
    {
        try
        {
            // 答えをtbxAnswerにロードする
            // tbxAnswer.Text = RowDataList.Where(row => row.RowNumber == CurrentRowNumber).Select(row => row.Answer).FirstOrDefault();
        }
        catch (Exception ex)
        {
            // MessageBox.Show("エラー: " + ex.Message,
            //                 "姉御に連絡して！",
            //                 MessageBoxButtons.OK,
            //                 MessageBoxIcon.Exclamation);
            throw ex;
        }
    }

    /// <summary>
    /// 次へボタン
    /// </summary>
    public void OnNextButtonClick()
    {
        try
        {
            // TODO もっときれいな書き方があるはず
            // もしCurrentRowNumberが最後なら、リセットする
            if (CurrentRowNumber != RowDataList.Last().RowNumber)
            {
                // 次のRowNumberを取得
                CurrentRowNumber = RowDataList.SkipWhile(row => row.RowNumber != CurrentRowNumber).Skip(1).First().RowNumber;

                // 次の問題をロードする
                // tbxQuestion.Text = RowDataList.Where(row => row.RowNumber == CurrentRowNumber).Select(row => row.Question).FirstOrDefault();

                // チェックボックスを更新する
                // chkShouldReview.Checked = Convert.ToBoolean(RowDataList.Where(row => row.RowNumber == CurrentRowNumber).Select(row => row.ShouldReview).FirstOrDefault());
            }
            else
            {
                // 最初の問題のIDを更新する
                CurrentRowNumber = RowDataList.First().RowNumber;

                // 最初の問題をロードする
                // tbxQuestion.Text = RowDataList.First().Question; ;

                // 最初の印をロードする
                // chkShouldReview.Checked = Convert.ToBoolean(RowDataList.First().ShouldReview);
            }

            // tbxAnswerを初期化
            // tbxAnswer.Text = String.Empty;

            // 問題番号を更新
            QuestionNumber++;
            QuestionNumber = QuestionNumber > RowDataList.Count() ? 1 : QuestionNumber;
            // lbQuestionNumber.Text = $"問題:{QuestionNumber}/{RowDataList.Count()}";
        }
        catch (Exception ex)
        {
            // MessageBox.Show("エラー: " + ex.Message,
            //                 "姉御に連絡して！",
            //                 MessageBoxButtons.OK,
            //                 MessageBoxIcon.Exclamation);
            throw ex;
        }
    }

    /// <summary>
    /// 戻るボタン
    /// </summary>
    public void OnBackButtonClick()
    {
        try
        {
            // TODO もっときれいな書き方があるはず
            // もしCurrentRowNumberが最後なら、リセットする
            if (CurrentRowNumber != RowDataList.First().RowNumber)
            {
                // ロードした問題のIDを更新する
                CurrentRowNumber = RowDataList.TakeWhile(row => row.RowNumber != CurrentRowNumber).Last().RowNumber;

                // 前の問題をロードする
                // tbxQuestion.Text = RowDataList.Where(row => row.RowNumber == CurrentRowNumber).Select(row => row.Question).FirstOrDefault();

                // チェックボックスを更新する
                // chkShouldReview.Checked = Convert.ToBoolean(RowDataList.Where(row => row.RowNumber == CurrentRowNumber).Select(row => row.ShouldReview).FirstOrDefault());
            }
            else
            {
                // 最後の問題のIDを更新する
                CurrentRowNumber = RowDataList.Last().RowNumber;

                // 最後の問題をロードする
                // tbxQuestion.Text = RowDataList.Last().Question;

                // 最後の印をロードする
                // chkShouldReview.Checked = Convert.ToBoolean(RowDataList.Last().ShouldReview);
            }

            // tbxAnswerを初期化
            // tbxAnswer.Text = String.Empty;

            // 問題番号を更新
            QuestionNumber--;
            QuestionNumber = QuestionNumber <= 0 ? RowDataList.Count() : QuestionNumber;
            // lbQuestionNumber.Text = $"問題:{QuestionNumber}/{RowDataList.Count()}";
        }
        catch (Exception ex)
        {
            // MessageBox.Show("エラー: " + ex.Message,
            //                 "姉御に連絡して！",
            //                 MessageBoxButtons.OK,
            //                 MessageBoxIcon.Exclamation);
            throw ex;
        }
    }

    /// <summary>
    /// 印チェック
    /// </summary>
    public void OnShouldReviewChanged()
    {
        try
        {
            // ListのshouldViewを更新する
            var currentRow = RowDataList.Where(row => row.RowNumber == CurrentRowNumber).First();
            // currentRow.ShouldReview = Convert.ToInt32(chkShouldReview.Checked);

            // エクセルを更新する
            // UpdateData.UpdateShouldReview(CurrentRowNumber, chkShouldReview.Checked, FilePath);
        }
        catch (Exception ex)
        {
            // MessageBox.Show("エラー: " + ex.Message,
            //                "姉御に連絡して！",
            //                MessageBoxButtons.OK,
            //                MessageBoxIcon.Exclamation);
            throw ex;
        }
    }

    /// <summary>
    /// 開始ボタン
    /// </summary>
    public void OnStartButtonClick()
    {
        // 開始パネルを非表示
        MainMenu.active = false;

        // 指定ディレクトリ（EXEと同じPath）にある.xlsxを取得
        string[] filePaths = Directory.GetFiles(Application.dataPath, "*.xlsx");

        if (filePaths.Length == 1)
        {
            // xlsxパスを変数に保存する
            FilePath = filePaths[0];

            // パスラベルを更新
            filePathText.text = $"ファイルパス: {FilePath}";
            btnText.text = FilePath;

            // 問題一覧の取得
            RowDataList = LoadData.LoadDataFromExcel(FilePath, false, false);

            // 問題番号リセットする
            QuestionNumber = 1;
            questionNumberText.text = $"問題:{QuestionNumber}/{RowDataList.Count()}";
        }
        else if (filePaths.Length > 1)
        {
            //　複数エクセを見つかった場合
            // MessageBox.Show("エラー: 複数エクセルが見つかりました。",
            //                 "姉御に連絡して！",
            //                 MessageBoxButtons.OK,
            //                 MessageBoxIcon.Exclamation);

            // 強制終了
            // Application.Exit();
        }
        else if (filePaths.Length <= 0)
        {
            //　エクセを見つからなかった場合
            // MessageBox.Show("エラー: エクセルがありません。",
            //                 "姉御に連絡して！",
            //                 MessageBoxButtons.OK,
            //                 MessageBoxIcon.Exclamation);

            // 強制終了
            // Application.Exit();
        }

        // chkShuffleにチェック入ったら、問題一覧をシャッフルする
        if (IsShuffleChecked)
        {
            RowDataList = RowDataList.OrderBy(row => Guid.NewGuid()).ToList();
        }
        else
        {
            RowDataList = RowDataList.OrderBy(row => row.RowNumber).ToList();
        }

        // ロードした問題のID
        CurrentRowNumber = RowDataList.First().RowNumber;

        // 一番目の問題をtbxQuestionにロードする
        // tbxQuestion.Text = RowDataList.First().Question;

        // 一番目の印をchkShouldReviewにロードする
        // chkShouldReview.Checked = Convert.ToBoolean(RowDataList.First().ShouldReview);

        // tbxAnswerを初期化
        // tbxAnswer.Text = String.Empty;
    }

    /// <summary>
    /// ホームボタン
    /// </summary>
    public void OnHomeButtonClick()
    {
        // 開始パネルを表示
        MainMenu.active = true;
    }

    /// <summary>
    /// 印の問題のみ表示する
    /// </summary>
    public void OnReviewOnlyChanged()
    {
        // IsReviewOnlyChcked = chkReviewOnly.Checked;
    }

    /// <summary>
    /// 問題をシャルルかどうか
    /// </summary>
    public void OnShuffleChanged()
    {
        // IsShuffleChecked = chkShuffle.Checked;
    }

    /// <summary>
    /// 問題を非表示するかどうか
    /// </summary>
    public void OnHideQuestionChanged()
    {
        // IsQuestionHidden = chkHideQuestion.Checked;
    }
}
