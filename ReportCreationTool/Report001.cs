using GrapeCity.Documents.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ReportCreationTool
{
    public class Report001
    {
        /// <summary>
        /// 履歴書の作成を行います。
        /// </summary>
        /// <returns>結果</returns>
        public Boolean Report001Output()
        {
            try
            {
                // Jsonファイルの読み込み
                string jsonFilePath = Define.JSON_FILE_PATH + Define.REPORT_001_JSON_FILE;

                if (!(System.IO.File.Exists(jsonFilePath)))
                {
                    // ファイル存在しない場合エラー
                    Common.LogOutput(Define.MESSEGE_ID_006 + jsonFilePath);
                    return false;
                }

                Report001JsonModel report001JsonModel = Common.Report001JsonRead(jsonFilePath);
                if (report001JsonModel == null)
                {
                    // エラー
                    Common.LogOutput(Define.MESSEGE_ID_004 + jsonFilePath);
                    return false;
                }

                // ワークブック作成
                Workbook workbook = new Workbook();
                string designFilePath = Define.DESIGN_FILE_PATH + Define.REPORT_001_DESIGN_FILE;

                if (!(System.IO.File.Exists(designFilePath)))
                {
                    // ファイル存在しない場合エラー
                    Common.LogOutput(Define.MESSEGE_ID_007 + designFilePath);
                    return false;
                }

                //　デザインファイルの読み込み
                workbook.Open(designFilePath, OpenFileFormat.Xlsx);

                IWorksheet settiongsSheet = workbook.Worksheets["Settings"];

                // Settingsシートからデータの割り当て位置を確認
                Report001SettingsModel report001SettingsModel = Common.DesignSettingsRead(settiongsSheet);
                if (report001SettingsModel == null)
                {
                    // エラー
                    Common.LogOutput(Define.MESSEGE_ID_005 + designFilePath);
                    return false;
                }

                // Jsonファイルチェック処理
                if (!(Common.JsonFileCheak(report001JsonModel, report001SettingsModel)))
                {
                    return false;
                }

                // ワークブック作成
                Workbook report001Workbook = new Workbook();

                // 履歴書シートのコピー
                IWorksheet copyRirekiSheet = workbook.Worksheets["履歴書"].CopyBefore(report001Workbook.Worksheets[0]);
                copyRirekiSheet.Name = "履歴書";
                copyRirekiSheet.Activate();

                // データ割り当て
                if(!(dataAllocation(ref copyRirekiSheet, report001JsonModel, report001SettingsModel)))
                {
                    Common.LogOutput(Define.MESSEGE_ID_008);
                    return false;
                }

                // PDF出力
                string pdfFilePath = Define.PDF_FILE_PATH + "report001" + DateTime.Now.ToString("yyyyMMddhhmmss") + ".pdf";
                copyRirekiSheet.Save(pdfFilePath, SaveFileFormat.Pdf);

                // PDF出力ログ
                Common.LogOutput(Define.MESSEGE_ID_020 + pdfFilePath);


                // JsonファイルをENDフォルダに移動させる。
                string moveFilePath = Define.MOVE_JSON_FILE_PATH + Define.REPORT_001_JSON_FILE;
                File.Move(jsonFilePath, moveFilePath, true);

                return true;
            }
            catch (Exception ex)
            {
                Common.LogOutput(Define.MESSEGE_ID_003);
                Common.LogOutput(ex.Message);
                Common.LogOutput(ex.StackTrace);
                return false;
            }
        }


        /// <summary>
        /// 履歴書シートにデータの割り当てを行う
        /// </summary>
        /// <param name="rirekiSheet">履歴書シート</param>
        /// <param name="report001JsonModel">JSONデータ</param>
        /// <param name="report001SettingsModel">Settingsシートのデータ</param>
        /// <returns>結果</returns>
        private Boolean dataAllocation(ref IWorksheet rirekiSheet, Report001JsonModel report001JsonModel, Report001SettingsModel report001SettingsModel)
        {
            try
            {
                // 名前（ふりがな）
                rirekiSheet.Range[report001SettingsModel.nameHiragana].Value = report001JsonModel.nameHiragana;

                // 名前（漢字）
                rirekiSheet.Range[report001SettingsModel.nameKanji].Value = report001JsonModel.nameKanji;

                // 性別
                rirekiSheet.Range[report001SettingsModel.gender].Value = report001JsonModel.gender;

                // 生年月日
                rirekiSheet.Range[report001SettingsModel.birthday].Value = report001JsonModel.birthday;

                // 年齢
                rirekiSheet.Range[report001SettingsModel.age].Value = report001JsonModel.age;

                // 郵便番号
                rirekiSheet.Range[report001SettingsModel.postCode].Value = report001JsonModel.postCode;

                // 住所
                rirekiSheet.Range[report001SettingsModel.address].Value = report001JsonModel.address;

                // 住所（ひらがな）
                rirekiSheet.Range[report001SettingsModel.addressHiragana].Value = report001JsonModel.addressHiragana;

                // 電話番号
                rirekiSheet.Range[report001SettingsModel.telephoneNumber].Value = report001JsonModel.telephoneNumber;

                //携帯電話
                rirekiSheet.Range[report001SettingsModel.mobilePhoneNumber].Value = report001JsonModel.mobilePhoneNumber;

                // メールアドレス
                rirekiSheet.Range[report001SettingsModel.email].Value = report001JsonModel.email;

                // 配偶者有無
                rirekiSheet.Range[report001SettingsModel.spouse].Value = report001JsonModel.spouse;

                // 扶養家族人数
                rirekiSheet.Range[report001SettingsModel.dependents].Value = report001JsonModel.dependents;

                // 学歴
                // 行を取得
                int educationalBackgrounRow = int.Parse(Regex.Replace(report001SettingsModel.educationalBackgroun, @"[^0-9]", ""));

                for (int i = 0; i < report001JsonModel.educationalBackgrounList.Count; i++)
                {

                    // 学歴（年）
                    rirekiSheet.Range[report001SettingsModel.educationalBackgrounYear].Value = report001JsonModel.educationalBackgrounList[i]["年"];

                    // 学歴（月）
                    rirekiSheet.Range[report001SettingsModel.educationalBackgrounMonth].Value = report001JsonModel.educationalBackgrounList[i]["月"];

                    // 学歴
                    rirekiSheet.Range[report001SettingsModel.educationalBackgroun].Value = report001JsonModel.educationalBackgrounList[i]["学歴"];

                    // 次の行のデータが存在する場合、割り当て位置を変更する
                    if (i < report001JsonModel.educationalBackgrounList.Count - 1)
                    {
                        string educationalBackgrounColumn = string.Empty;
                        educationalBackgrounRow++;

                        //----------------
                        // 学歴（年）
                        //----------------
                        // 列を取得
                        educationalBackgrounColumn = Regex.Replace(report001SettingsModel.educationalBackgrounYear, @"[^a-zA-Z]", "");
                        // 位置を変更
                        report001SettingsModel.educationalBackgrounYear = educationalBackgrounColumn + educationalBackgrounRow.ToString();

                        //---------------
                        // 学歴（月）
                        //---------------
                        // 列を取得
                        educationalBackgrounColumn = Regex.Replace(report001SettingsModel.educationalBackgrounMonth, @"[^a-zA-Z]", "");
                        // 位置を変更
                        report001SettingsModel.educationalBackgrounMonth = educationalBackgrounColumn + educationalBackgrounRow.ToString();

                        //--------------
                        // 学歴（月）
                        //--------------
                        // 列を取得
                        educationalBackgrounColumn = Regex.Replace(report001SettingsModel.educationalBackgroun, @"[^a-zA-Z]", "");
                        // 位置を変更
                        report001SettingsModel.educationalBackgroun = educationalBackgrounColumn + educationalBackgrounRow.ToString();
                    }
                }

                // 職歴
                // 行を取得
                int workHistoryRow = int.Parse(Regex.Replace(report001SettingsModel.workHistory, @"[^0-9]", ""));

                for (int i = 0; i < report001JsonModel.workHistoryList.Count; i++)
                {

                    // 職歴（年）
                    rirekiSheet.Range[report001SettingsModel.workHistoryYear].Value = report001JsonModel.workHistoryList[i]["年"];

                    // 職歴（月）
                    rirekiSheet.Range[report001SettingsModel.workHistoryMonth].Value = report001JsonModel.workHistoryList[i]["月"];

                    // 職歴
                    rirekiSheet.Range[report001SettingsModel.workHistory].Value = report001JsonModel.workHistoryList[i]["職歴"];

                    // 次の行のデータが存在する場合、割り当て位置を変更する
                    if (i < report001JsonModel.workHistoryList.Count - 1)
                    {
                        string workHistoryColumn = string.Empty;
                        workHistoryRow++;

                        //----------------
                        // 職歴（年）
                        //----------------
                        // 列を取得
                        workHistoryColumn = Regex.Replace(report001SettingsModel.workHistoryYear, @"[^a-zA-Z]", "");
                        // 位置を変更
                        report001SettingsModel.workHistoryYear = workHistoryColumn + workHistoryRow.ToString();

                        //---------------
                        // 職歴（月）
                        //---------------
                        // 列を取得
                        workHistoryColumn = Regex.Replace(report001SettingsModel.workHistoryMonth, @"[^a-zA-Z]", "");
                        // 位置を変更
                        report001SettingsModel.workHistoryMonth = workHistoryColumn + workHistoryRow.ToString();

                        //--------------
                        // 職歴（月）
                        //--------------
                        // 列を取得
                        workHistoryColumn = Regex.Replace(report001SettingsModel.workHistory, @"[^a-zA-Z]", "");
                        // 位置を変更
                        report001SettingsModel.workHistory = workHistoryColumn + workHistoryRow.ToString();
                    }
                }

                // 資格
                // 行を取得
                int qualificationRow = int.Parse(Regex.Replace(report001SettingsModel.qualification, @"[^0-9]", ""));

                for (int i = 0; i < report001JsonModel.qualificationList.Count; i++)
                {

                    // 職歴（年）
                    rirekiSheet.Range[report001SettingsModel.qualificationYear].Value = report001JsonModel.qualificationList[i]["年"];

                    // 職歴（月）
                    rirekiSheet.Range[report001SettingsModel.qualificationMonth].Value = report001JsonModel.qualificationList[i]["月"];

                    // 職歴
                    rirekiSheet.Range[report001SettingsModel.qualification].Value = report001JsonModel.qualificationList[i]["資格"];

                    // 次の行のデータが存在する場合、割り当て位置を変更する
                    if (i < report001JsonModel.qualificationList.Count - 1)
                    {
                        string qualificationColumn = string.Empty;
                        qualificationRow++;

                        //----------------
                        // 職歴（年）
                        //----------------
                        // 列を取得
                        qualificationColumn = Regex.Replace(report001SettingsModel.qualificationYear, @"[^a-zA-Z]", "");
                        // 位置を変更
                        report001SettingsModel.qualificationYear = qualificationColumn + qualificationRow.ToString();

                        //---------------
                        // 職歴（月）
                        //---------------
                        // 列を取得
                        qualificationColumn = Regex.Replace(report001SettingsModel.qualificationMonth, @"[^a-zA-Z]", "");
                        // 位置を変更
                        report001SettingsModel.qualificationMonth = qualificationColumn + qualificationRow.ToString();

                        //--------------
                        // 職歴（月）
                        //--------------
                        // 列を取得
                        qualificationColumn = Regex.Replace(report001SettingsModel.qualification, @"[^a-zA-Z]", "");
                        // 位置を変更
                        report001SettingsModel.qualification = qualificationColumn + qualificationRow.ToString();
                    }
                }

                // 備考
                rirekiSheet.Range[report001SettingsModel.remarks].Value = report001JsonModel.remarks;

                // 現在日
                rirekiSheet.Range[report001SettingsModel.currentDate].Value = report001JsonModel.currentDate;

                return true;
            }
            catch (Exception ex)
            {
                Common.LogOutput(Define.MESSEGE_ID_003);
                Common.LogOutput(ex.Message);
                Common.LogOutput(ex.StackTrace);
                return false;
            }

            
        }




    }
}
