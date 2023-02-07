using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportCreationTool
{
    /// <summary>
    /// Settingsシートのデータ格納モデル
    /// </summary>
    public class Report001SettingsModel
    {
        // 名前（漢字）
        public string nameKanji { get; set; }

        // 名前（ひらがな）
        public string nameHiragana { get; set; }

        // 性別
        public string gender { get; set; }

        // 誕生日
        public string birthday { get; set; }

        // 年齢
        public string age { get; set; }

        // 郵便番号
        public string postCode { get; set; }

        // 住所
        public string address { get; set; }

        // 住所（ひらがな）
        public string addressHiragana { get; set; }

        // 電話番号
        public string telephoneNumber { get; set; }

        // 携帯電話
        public string mobilePhoneNumber { get; set; }

        // メールアドレス
        public string email { get; set; }

        // 配偶者有無
        public string spouse { get; set; }

        // 扶養家族人数
        public string dependents { get; set; }

        // 学歴
        public string educationalBackgroun { get; set; }

        // 学歴（年）
        public string educationalBackgrounYear { get; set; }

        // 学歴（月）
        public string educationalBackgrounMonth { get; set; }

        // 職歴
        public string workHistory { get; set; }

        // 職歴（年）
        public string workHistoryYear { get; set; }

        // 職歴（月）
        public string workHistoryMonth { get; set; }

        // 資格
        public string qualification { get; set; }

        // 資格（年）
        public string qualificationYear { get; set; }

        // 資格（月）
        public string qualificationMonth { get; set; }

        // 備考
        public string remarks { get; set; }

        // 現在日
        public string currentDate { get; set; }

        // 学歴最大行
        public string educationalBackgrounMax { get; set; }

        // 職歴最大行
        public string workHistoryMax { get; set; }

        // 資格最大行
        public string qualificationMax { get; set; }
    }
}
