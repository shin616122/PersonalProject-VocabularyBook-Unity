using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using UnityEngine;
using UnityEngine.UI;

public class CanvasManager : MonoBehaviour
{
    // Unity Object
    public GameObject MainMenu;
    public GameObject QuestionPanel;
    public Text filePathText;
    public Text questionNumberText;
    public Text tbxQuestion;
    public Text tbxAnswer;
    public Toggle hideQuestionToggle;
    public Toggle shuffleToggle;
    public Toggle reviewOnlyToggle;
    public Toggle shouldReviewToggle;

    private List<RowData> RowDataList { get; set; }
    private int CurrentRowNumber { get; set; }
    public string FilePath { get; set; }
    public bool IsShuffleChecked { get; set; }
    public bool IsReviewOnlyChcked { get; set; }
    public bool IsQuestionHidden { get; set; }
    public int QuestionNumber { get; set; } = 1;
    private bool firstTime = true;

    /// <summary>
    /// 回答ボタン
    /// </summary>
    public void OnAnswerButtonClick()
    {
        try
        {
            // 答えをtbxAnswerにロードする
            tbxAnswer.text = RowDataList.Where(row => row.RowNumber == CurrentRowNumber).Select(row => row.Answer).FirstOrDefault();
        }
        catch (Exception ex)
        {
            // MessageBox.Show("エラー: " + ex.Message,
            //                 "アドミンに連絡してください！",
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
                tbxQuestion.text = RowDataList.Where(row => row.RowNumber == CurrentRowNumber).Select(row => row.Question).FirstOrDefault();

                // チェックボックスを更新する
                shouldReviewToggle.isOn = Convert.ToBoolean(RowDataList.Where(row => row.RowNumber == CurrentRowNumber).Select(row => row.ShouldReview).FirstOrDefault());
            }
            else
            {
                // 最初の問題のIDを更新する
                CurrentRowNumber = RowDataList.First().RowNumber;

                // 最初の問題をロードする
                tbxQuestion.text = RowDataList.First().Question; ;

                // 最初の印をロードする
                shouldReviewToggle.isOn = Convert.ToBoolean(RowDataList.First().ShouldReview);
            }

            // tbxAnswerを初期化
            tbxAnswer.text = String.Empty;

            // 問題番号を更新
            QuestionNumber++;
            QuestionNumber = QuestionNumber > RowDataList.Count() ? 1 : QuestionNumber;
            questionNumberText.text = $"問題:{QuestionNumber}/{RowDataList.Count()}";
        }
        catch (Exception ex)
        {
            // MessageBox.Show("エラー: " + ex.Message,
            //                 "アドミンに連絡してください！",
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
                tbxQuestion.text = RowDataList.Where(row => row.RowNumber == CurrentRowNumber).Select(row => row.Question).FirstOrDefault();

                // チェックボックスを更新する
                shouldReviewToggle.isOn = Convert.ToBoolean(RowDataList.Where(row => row.RowNumber == CurrentRowNumber).Select(row => row.ShouldReview).FirstOrDefault());
            }
            else
            {
                // 最後の問題のIDを更新する
                CurrentRowNumber = RowDataList.Last().RowNumber;

                // 最後の問題をロードする
                tbxQuestion.text = RowDataList.Last().Question;

                // 最後の印をロードする
                shouldReviewToggle.isOn = Convert.ToBoolean(RowDataList.Last().ShouldReview);
            }

            // tbxAnswerを初期化
            tbxAnswer.text = String.Empty;

            // 問題番号を更新
            QuestionNumber--;
            QuestionNumber = QuestionNumber <= 0 ? RowDataList.Count() : QuestionNumber;
            questionNumberText.text = $"問題:{QuestionNumber}/{RowDataList.Count()}";
        }
        catch (Exception ex)
        {
            // MessageBox.Show("エラー: " + ex.Message,
            //                 "アドミンに連絡してください！",
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
            currentRow.ShouldReview = Convert.ToInt32(shouldReviewToggle.isOn);

            // エクセルを更新する
            UpdateData.UpdateShouldReview(CurrentRowNumber, shouldReviewToggle.isOn, FilePath);
        }
        catch (Exception ex)
        {
            // MessageBox.Show("エラー: " + ex.Message,
            //                "アドミンに連絡してください！",
            //                MessageBoxButtons.OK,
            //                MessageBoxIcon.Exclamation);
            throw ex;
        }
    }

    /// <summary>
    /// 開始ボタン
    /// </summary>
    public void OnStartButtonClick(int flag)
    {
        // 開始パネルを非表示
        MainMenu.active = false;

        if (FilePath != "")
        {
            if (flag == 0 && firstTime)
            {
                // Load filepath from data path - only when start button is clicked
                GetFileFromDataPath();
                firstTime = false;
            }
            LoadFileFromPath();
        }
        else
        {
            //　複数エクセを見つかった場合
            // MessageBox.Show("エラー: ファイルパスを確認してください",
            //                 "アドミンに連絡してください！",
            //                 MessageBoxButtons.OK,
            //                 MessageBoxIcon.Exclamation);

            // 強制終了
            // Application.Exit();
        }
    }

    private void GetFileFromDataPath()
    {
        // 指定ディレクトリ（EXEと同じPath）にある.xlsxを取得
        string[] filePaths = Directory.GetFiles(Application.dataPath, "*.xlsx");
        if (filePaths.Length == 1)
        {
            // xlsxパスを変数に保存する
            FilePath = filePaths[0];
        }
        else if (filePaths.Length > 1)
        {
            //　複数エクセを見つかった場合
            // MessageBox.Show("エラー: 複数エクセルが見つかりました。",
            //                 "アドミンに連絡してください！",
            //                 MessageBoxButtons.OK,
            //                 MessageBoxIcon.Exclamation);

            // 強制終了
            // Application.Exit();
        }
        else if (filePaths.Length <= 0)
        {
            //　エクセを見つからなかった場合
            // MessageBox.Show("エラー: エクセルがありません。",
            //                 "アドミンに連絡してください！",
            //                 MessageBoxButtons.OK,
            //                 MessageBoxIcon.Exclamation);

            // 強制終了
            // Application.Exit();
        }
    }

    public void LoadFileFromPath()
    {
        // パスラベルを更新
        filePathText.text = $"ファイルパス: {FilePath}";

        // 問題一覧の取得
        RowDataList = LoadData.LoadDataFromExcel(FilePath, IsReviewOnlyChcked, IsQuestionHidden);

        // 問題番号リセットする
        QuestionNumber = 1;
        questionNumberText.text = $"問題:{QuestionNumber}/{RowDataList.Count()}";

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
        tbxQuestion.text = RowDataList.First().Question;

        // 一番目の印をchkShouldReviewにロードする
        shouldReviewToggle.isOn = Convert.ToBoolean(RowDataList.First().ShouldReview);

        // tbxAnswerを初期化
        tbxAnswer.text = String.Empty;
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
        IsReviewOnlyChcked = reviewOnlyToggle.isOn;
    }

    /// <summary>
    /// 問題をシャルルかどうか
    /// </summary>
    public void OnShuffleChanged()
    {
        IsShuffleChecked = shuffleToggle.isOn;
    }

    /// <summary>
    /// 問題を非表示するかどうか
    /// </summary>
    public void OnHideQuestionChanged()
    {
        IsQuestionHidden = hideQuestionToggle.isOn;
    }

    /// <summary>
    ///アプリを閉じる
    /// </summary>
    public void OnCloseButtonClick()
    {
        Application.Quit();
    }
}
