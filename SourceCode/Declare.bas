Attribute VB_Name = "Declare"
Option Compare Database
Option Explicit


' ******************************************************
' グローバル変数
' ******************************************************
Public flg最適化 As String  '自動最適化処理の実行許可




' ******************************************************
' グローバル定数
' ******************************************************
Public Const PATH_BIN As String = "_Bin\"
Public Const PATH_ZIP As String = "_Zip\"
Public Const PATH_DATA As String = "Data\"
Public Const PATH_LOG As String = "Log\"

Public Const SYS拡張子 As String = ".sldt"


Public Const F01_OptimizeForm As String = "F01_最適化"
Public Const F02_AppInfo_Display As String = "F02_アプリ情報_参照"
Public Const F02_AppInfo_Input As String = "F02_アプリ情報_入力"
Public Const F03_Toast As String = "F03_トースト"

Public Const F11_MainMenu As String = "F11_メインメニュー"
Public Const F12_InfomationDialog As String = "F12_伝達ダイアログ"
Public Const F13_ConfirmDialog As String = "F13_確認ダイアログ"
Public Const F14_HistoryMemo As String = "F14_履歴メモ"
Public Const F15_ExpQuery As String = "F15_クエリ出力"
Public Const F16_LinkManager As String = "F16_データ管理"

Public Const FA1_ImportData As String = "FA1_取り込み用"
Public Const FA2_ImportDataCSV As String = "FA2_CSV一括取り込み"
Public Const FA3_DataDelete As String = "FA3_データ削除"

Public Const FB0_MstStore As String = "FB0_店舗マスタ"
Public Const FB1_MstMachine As String = "FB1_機種マスタ"
Public Const FB2_MstSamaisu As String = "FB2_差枚数係数算出"
Public Const FB3_MstMachineNum As String = "FB3_台番号マスタ"
Public Const FB4_MstMachineNumType As String = "FB4_台タイプマスタ"
Public Const FB5_MstEvent As String = "FB5_イベントマスタ"
Public Const FB5_MstEvent_Bulk As String = "FB5_イベントマスタ_一括編集"
Public Const FB6_MstEventType As String = "FB6_イベントタイプマスタ"
Public Const FB7_IslandNum As String = "FB7_島番号_一括編集"

Public Const FC1_DataView As String = "FC1_データ一覧用"
Public Const FC2_DataSummary As String = "FC2_データ集計用"
Public Const FC3_DataTimeline As String = "FC3_データ日付推移"
Public Const FC4_DataClossSummary As String = "FC4_クロス集計用"
Public Const FC9_DataDateInfo As String = "FC9_データ単票"

Public Const FD1_GraphTrends As String = "FD1_グラフ_傾向"
Public Const FD2_GraphCompare As String = "FD2_グラフ_比較"
Public Const FD3_GraphCustom As String = "FD3_グラフ_カスタム"
Public Const FD9_GraphCoinTrend As String = "FD9_グラフ_出玉推移"



' ******************************************************
' 列挙体
' ******************************************************
Public Const E_Color_Black As Integer = 0
Public Const E_Color_Blue As Integer = 1
Public Const E_Color_Red As Integer = 2
Public Const E_Color_Green As Integer = 3
Public Const E_Color_DeepBlue As Integer = 4
Public Const E_Color_Purple As Integer = 5






