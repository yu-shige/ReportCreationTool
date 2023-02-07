using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportCreationTool
{

    public static class Define
    {
        // メッセージ
        public static readonly string MESSEGE_ID_001 = "帳票作成ツールを開始します。";
        public static readonly string MESSEGE_ID_002 = "帳票作成ツールを終了します。";
        public static readonly string MESSEGE_ID_003 = "例外エラーが発生しました。";
        public static readonly string MESSEGE_ID_004 = "Jsonファイルの読み込みに失敗しました。";
        public static readonly string MESSEGE_ID_005 = "Settingsシートからデータ取得中にエラー発生しました。";
        public static readonly string MESSEGE_ID_006 = "JSONファイルが存在しません。パス：";
        public static readonly string MESSEGE_ID_007 = "デザインファイルが存在しません。パス：";
        public static readonly string MESSEGE_ID_008 = "履歴書シートにデータ割り当て中にエラーが発生しました。";
        public static readonly string MESSEGE_ID_009 = "Jsonデータの学歴の行数が履歴書シートの最大行より多くなっています。";
        public static readonly string MESSEGE_ID_010 = "Jsonデータの職歴の行数が履歴書シートの最大行より多くなっています。";
        public static readonly string MESSEGE_ID_011 = "Jsonデータの資格の行数が履歴書シートの最大行より多くなっています。";
        public static readonly string MESSEGE_ID_012 = "Settingsシートの値に誤りがあります。";

        public static readonly string MESSEGE_ID_020 = "履歴書が作成されました。パス：";
        public static readonly string MESSEGE_ID_021 = "履歴書のPDF作成に成功しました。";
        public static readonly string MESSEGE_ID_022 = "履歴書のPDF作成に失敗しました。";

        // ファイルパス
        public static readonly string LOG_FILE_PATH = @"..\\..\\..\\..\\LOG\\";

        public static readonly string DESIGN_FILE_PATH = @"..\\..\\..\\..\\DESIGNFILE\\";

        public static readonly string JSON_FILE_PATH = @"..\\..\\..\\..\\JSONFILE\\";

        public static readonly string MOVE_JSON_FILE_PATH = @"..\\..\\..\\..\\JSONFILE\\END\\";

        public static readonly string PDF_FILE_PATH = @"..\\..\\..\\..\\PDF\\";

        public static readonly string REPORT_001_JSON_FILE = "report001.json";

        public static readonly string REPORT_001_DESIGN_FILE = "履歴書.xlsx";

        // セクション
        public static readonly string SECTION_SETTINGS = "[settings]";

        // セクション
        public static readonly string SECTION_OTHER = "[other]";

    }
}
