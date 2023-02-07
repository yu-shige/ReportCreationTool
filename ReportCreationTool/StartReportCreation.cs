using ReportCreationTool;

public class StartReportCreation
{
    /// <summary>
    /// 開始メソッド
    /// </summary>
    /// <param name="args"></param>
    static void Main(string[] args)
    {
        try
        {
            // Json作成
            Common.Report001JsonCreate();

            // 開始ログ
            Common.LogOutput(Define.MESSEGE_ID_001);

            // Jsonファイル確認
            string[] jsonFiles = Directory.GetFiles(Define.JSON_FILE_PATH, "*.json");

            if (jsonFiles.Length > 0)
            {
                //Jsonファイルがある場合、帳票作成開始
                StartReport();
            }

            // 終了ログ
            Common.LogOutput(Define.MESSEGE_ID_002);
        }
        catch (Exception ex)
        {
            Common.LogOutput(Define.MESSEGE_ID_003);
            Common.LogOutput(ex.Message);
            Common.LogOutput(ex.StackTrace);
        }
    }


    /// <summary>
    /// 帳票作成開始
    /// </summary>
    private static void StartReport()
    {
        Report001 report001 = new Report001();

        var result = report001.Report001Output();

        if (result)
        {
            // 成功
            Common.LogOutput(Define.MESSEGE_ID_021);
        }
        else
        {
            // 失敗
            Common.LogOutput(Define.MESSEGE_ID_022);

        }

    }
}