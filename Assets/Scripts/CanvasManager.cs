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


    // Functions


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
}
