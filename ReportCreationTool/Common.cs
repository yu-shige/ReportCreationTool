using GrapeCity.DataVisualization.TypeScript;
using GrapeCity.Documents.Excel;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ReportCreationTool
{
    /// <summary>
    /// 共通クラス
    /// </summary>
    public static class Common
    {

        /// <summary>
        /// ログ出力を行う
        /// </summary>
        /// <param name="messege">メッセージ</param>
        public static void LogOutput(string messege)
        {
            string fileName = Define.LOG_FILE_PATH + DateTime.Now.ToString("yyyyMMdd") + ".txt";

            // ファイル存在チェック
            if (System.IO.File.Exists(fileName))
            {
                // メッセージ追加
                File.AppendAllText(fileName, System.Environment.NewLine + messege);
            }
            else
            {
                // ファイル作成してメッセージ追加
                File.WriteAllText(fileName, messege);
            }
        }

        /// <summary>
        /// Jsonファイル作成(Report001)
        /// </summary>
        public static void Report001JsonCreate()
        {
            // 履歴書のJSONデータ作成
            Report001JsonModel report001JsonModel = new Report001JsonModel();

            report001JsonModel.nameKanji = "〇〇　優";
            report001JsonModel.nameHiragana = "〇〇　ゆう";
            report001JsonModel.gender = "男";
            report001JsonModel.birthday = "199x年1x月4日";
            report001JsonModel.age = "31歳";
            report001JsonModel.postCode = "121-0836";
            report001JsonModel.address = "東京都足立区入谷";
            report001JsonModel.addressHiragana = "とうきょうとあだちくいりや";

            report001JsonModel.telephoneNumber = "0263-55-5555";
            report001JsonModel.mobilePhoneNumber = "090-2222-1111";
            report001JsonModel.email = "xxx@gmail.co.jp";
            report001JsonModel.spouse = "無";
            report001JsonModel.dependents = "0人";


            // 学歴
            for (int i = 0; i < 5; i++)
            {

                Dictionary<string, string> educationalBackgroun = new Dictionary<string, string>();

                switch (i)
                {
                    case 0:
                        educationalBackgroun.Add("年", "2007") ;
                        educationalBackgroun.Add("月", "3");
                        educationalBackgroun.Add("学歴", "〇〇中学校　卒業");
                        report001JsonModel.educationalBackgrounList.Add(educationalBackgroun);
                        break;

                    case 1:
                        educationalBackgroun.Add("年", "2007");
                        educationalBackgroun.Add("月", "4");
                        educationalBackgroun.Add("学歴", "〇〇三高校　普通科　入学");
                        report001JsonModel.educationalBackgrounList.Add(educationalBackgroun);
                        break;

                    case 2:
                        educationalBackgroun.Add("年", "2010");
                        educationalBackgroun.Add("月", "3");
                        educationalBackgroun.Add("学歴", "〇〇高校　普通科　卒業");
                        report001JsonModel.educationalBackgrounList.Add(educationalBackgroun);
                        break;

                    case 3:
                        educationalBackgroun.Add("年", "2010");
                        educationalBackgroun.Add("月", "4");
                        educationalBackgroun.Add("学歴", "〇〇大学情報通信学部組込みソフトウェア工学科　入学");
                        report001JsonModel.educationalBackgrounList.Add(educationalBackgroun);
                        break;
                    case 4:
                        educationalBackgroun.Add("年", "2014");
                        educationalBackgroun.Add("月", "3");
                        educationalBackgroun.Add("学歴", "〇〇大学情報通信学部組込みソフトウェア工学科　卒業");
                        report001JsonModel.educationalBackgrounList.Add(educationalBackgroun);
                        break;
                    default: 
                        break;
                }


            }

            // 職歴
            for (int i = 0; i < 7; i++)
            {

                Dictionary<string, string> workHistory = new Dictionary<string, string>();

                switch (i)
                {
                    case 0:
                        workHistory.Add("年", "2016");
                        workHistory.Add("月", "2");
                        workHistory.Add("職歴", "医療法人　〇〇会　〇〇病院　入社");
                        report001JsonModel.workHistoryList.Add(workHistory);
                        break;

                    case 1:
                        workHistory.Add("年", "2017");
                        workHistory.Add("月", "7");
                        workHistory.Add("職歴", "医療法人　〇〇会　〇〇病院　退社");
                        report001JsonModel.workHistoryList.Add(workHistory);
                        break;

                    case 2:
                        workHistory.Add("年", "2018");
                        workHistory.Add("月", "3");
                        workHistory.Add("職歴", "株式会社〇〇デザイン　入社");
                        report001JsonModel.workHistoryList.Add(workHistory);
                        break;

                    case 3:
                        workHistory.Add("年", "2020");
                        workHistory.Add("月", "1");
                        workHistory.Add("職歴", "株式会社〇〇デザイン　退社");
                        report001JsonModel.workHistoryList.Add(workHistory);
                        break;
                    case 4:
                        workHistory.Add("年", "2020");
                        workHistory.Add("月", "2");
                        workHistory.Add("職歴", "〇〇システム株式会社　入社");
                        report001JsonModel.workHistoryList.Add(workHistory);
                        break;
                    case 5:
                        workHistory.Add("年", "2022");
                        workHistory.Add("月", "12");
                        workHistory.Add("職歴", "〇〇システム株式会社　退社");
                        report001JsonModel.workHistoryList.Add(workHistory);
                        break;
                    case 6:
                        workHistory.Add("年", "2023");
                        workHistory.Add("月", "1");
                        workHistory.Add("職歴", "個人事業主（フリーランスエンジニア）として活動");
                        report001JsonModel.workHistoryList.Add(workHistory);
                        break;
                    default:
                        break;
                }


            }

            // 資格
            for (int i = 0; i < 5; i++)
            {

                Dictionary<string, string> qualification = new Dictionary<string, string>();

                switch (i)
                {
                    case 0:
                        qualification.Add("年", "2010");
                        qualification.Add("月", "8");
                        qualification.Add("資格", "普通自動車免許");
                        report001JsonModel.qualificationList.Add(qualification);
                        break;

                    case 1:
                        qualification.Add("年", "2012");
                        qualification.Add("月", "8");
                        qualification.Add("資格", "普通自動二輪免許");
                        report001JsonModel.qualificationList.Add(qualification);
                        break;

                    case 2:
                        qualification.Add("年", "2014");
                        qualification.Add("月", "6");
                        qualification.Add("資格", "日商簿記検定 2級");
                        report001JsonModel.qualificationList.Add(qualification);
                        break;

                    case 3:
                        qualification.Add("年", "2019");
                        qualification.Add("月", "5");
                        qualification.Add("資格", "Oracle Certified Java Programmer Silver SE 8");
                        report001JsonModel.qualificationList.Add(qualification);
                        break;
                    case 4:
                        qualification.Add("年", "2019");
                        qualification.Add("月", "10");
                        qualification.Add("資格", "ORACLE MASTER Bronze 12c");
                        report001JsonModel.qualificationList.Add(qualification);
                        break;
                    default:
                        break;
                }
            }

            // 現在日
            report001JsonModel.currentDate = DateTime.Now.ToString("yyyy年MM月dd日");

            // 備考
            report001JsonModel.remarks = "エンジニア職を希望します。";

            string jsonFilepath = Define.JSON_FILE_PATH + "report001.json";
            try
            {
                using (var sw = new StreamWriter(jsonFilepath, false, System.Text.Encoding.UTF8))
                {
                    // JSON データにシリアライズ
                    var jsonData = JsonConvert.SerializeObject(report001JsonModel, Formatting.Indented);

                    // JSON データをファイルに書き込み
                    sw.Write(jsonData);
                }
            }
            catch (Exception ex)
            {
                Common.LogOutput(Define.MESSEGE_ID_003);
                Common.LogOutput(ex.Message);
            }
        }

        /// <summary>
        /// Jsonファイル読み込み
        /// </summary>
        /// <param name="jsonFilePath">Jsonファイルパス</param>
        /// <returns>Jsonデータ格納モデル</returns>
        public static Report001JsonModel Report001JsonRead(string jsonFilePath)
        {
            Report001JsonModel report001JsonModel = new Report001JsonModel();

            try
            {
                Encoding enc = Encoding.UTF8;
                using (var reader = new System.IO.StreamReader(jsonFilePath, enc))
                {
                    string jsonStr = reader.ReadToEnd();
                    report001JsonModel = JsonConvert.DeserializeObject<Report001JsonModel>(jsonStr);
                }

                return report001JsonModel;
            }
            catch (Exception ex)
            {
                Common.LogOutput(Define.MESSEGE_ID_003);
                Common.LogOutput(ex.Message);
                Common.LogOutput(ex.StackTrace);
                return null;
            }

            
        }

        /// <summary>
        /// Settingsシートから各項目のデータを取得する
        /// </summary>
        /// <param name="settingsSheet"></param>
        /// <returns>成功：Settingsシート格納モデル、失敗：null</returns>
        public static Report001SettingsModel DesignSettingsRead(IWorksheet settingsSheet)
        {
            Report001SettingsModel report001SettingsModel = new Report001SettingsModel();

            Dictionary<string, string> settingsData= new Dictionary<string, string>();
            Dictionary<string, string> otherData = new Dictionary<string, string>();
            int settingsCount = 0;

            const string SECTION_FLG_SETTINGS = "0";
            const string SECTION_FLG_OTHER = "1";

            // 0:settings、1:other
            string sectionFlg = SECTION_FLG_SETTINGS;

            try
            {
                // Settingsシートから項目を取得
                do
                {
                    settingsCount++;
                    string settingsARange = "A" + settingsCount.toString();
                    string settingsBRange = "B" + settingsCount.toString();

                    // A列から値を取得
                    object aValue = settingsSheet.Range[settingsARange].Value;

                    if (aValue == null)
                    {
                        break;
                    }

                    // []が含まれている場合はセクション
                    if (Regex.IsMatch(aValue.toString(), @"^\[.+\]$"))
                    {
                        if(aValue.ToString() == Define.SECTION_SETTINGS)
                        {
                            sectionFlg = SECTION_FLG_SETTINGS;
                        }
                        else if (aValue.ToString() == Define.SECTION_OTHER)
                        {
                            sectionFlg = SECTION_FLG_OTHER;
                        }
                        else
                        {
                            continue;
                        }

                        continue;
                    }

                    // B列から値を取得
                    object bValue = settingsSheet.Range[settingsBRange].Value;

                    // 値のチェック
                    if (bValue == null)
                    {
                        // エラーメッセージ
                    }

                    switch (sectionFlg) 
                    { 
                        case SECTION_FLG_SETTINGS:
                        
                            if (!(Regex.IsMatch(bValue.toString(), @"^[a-zA-Z]+[0-9]+$")))
                            {
                                // エラーメッセージ
                                Common.LogOutput(Define.MESSEGE_ID_012 + "項目：" + aValue.ToString() + "値：" + bValue.ToString());
                                return null;
                            }
                            break;
                        case SECTION_FLG_OTHER:

                            if (!(Regex.IsMatch(bValue.toString(), @"^[0-9]+$")))
                            {
                                // エラーメッセージ
                                Common.LogOutput(Define.MESSEGE_ID_012 + "項目：" + aValue.ToString() + "値：" + bValue.ToString());
                                return null;
                            }
                            break;
                        default:
                            break;
                    }

                    // モデルに取得したデータを格納する
                    var property = typeof(Report001SettingsModel).GetProperty(aValue.ToString());
                    property.SetValue(report001SettingsModel, bValue.toString());

                } while(true);


                return report001SettingsModel;

            }
            catch (Exception ex)
            {
                Common.LogOutput(Define.MESSEGE_ID_003);
                Common.LogOutput(ex.Message);
                Common.LogOutput(ex.StackTrace);
                return null;
            }

            

        }

        /// <summary>
        /// Jsonデータのチェック処理
        /// </summary>
        /// <param name="report001JsonModel">Jsonデータ</param>
        /// <param name="report001SettingsModel">Settingsシートのデータ</param>
        /// <returns>結果</returns>
        public static Boolean JsonFileCheak(Report001JsonModel report001JsonModel, Report001SettingsModel report001SettingsModel)
        {
            // 学歴最大行のチェック
            if (report001JsonModel.educationalBackgrounList.Count > int.Parse(report001SettingsModel.educationalBackgrounMax))
            {
                Common.LogOutput(Define.MESSEGE_ID_009);
                return false;
            }

            // 職歴最大行のチェック
            if (report001JsonModel.workHistoryList.Count > int.Parse(report001SettingsModel.workHistoryMax))
            {
                Common.LogOutput(Define.MESSEGE_ID_010);
                return false;
            }

            // 資格最大行のチェック
            if (report001JsonModel.qualificationList.Count > int.Parse(report001SettingsModel.qualificationMax))
            {
                Common.LogOutput(Define.MESSEGE_ID_011);
                return false;
            }

            return true;
        }

    }
}
