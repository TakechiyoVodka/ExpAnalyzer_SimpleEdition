using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = ClosedXML.Excel;
using EPPExcel = OfficeOpenXml;
using EPPExcelChart = OfficeOpenXml.Drawing.Chart;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using static ExpAnalyzer_SimpleEdition.FormMain.ClassUnitData;
using static ExpAnalyzer_SimpleEdition.FormMain.ClassShapeUnitData;
using static ExpAnalyzer_SimpleEdition.FormMain.ClassGraphData;
using DocumentFormat.OpenXml.Vml;
using OfficeOpenXml.Drawing.Chart;
using DocumentFormat.OpenXml.Wordprocessing;

namespace ExpAnalyzer_SimpleEdition
{
    public partial class FormMain : Form
    {
        public FormMain()
        {
            InitializeComponent();
        }
        #region 静的フィールド変数
        public static string workbookPath;
        public static List<ClassModelInfo> ModelInfoList;
        public static List<ClassUnitData> UnitDataList;
        public static List<ClassShapeUnitData> ShapeUnitDataList;
        public static List<ClassShapeUnitData> ReShapeUnitDataList;
        public static List<ClassGraphData> GraphDataList;
        #endregion
        /// <summary>
        /// 機種情報クラス
        /// </summary>
        public class ClassModelInfo
        {
            public string ModelName;
            public string FirstHitProb;
            public string ProbVarHitRashRate;
            public string ProbVarHitPersisRate;
        }
        /// <summary>
        /// 台データクラス
        /// </summary>
        public class ClassUnitData
        {
            public string ModelName;
            public int InstallNum = 0;
            public List<ClassUnitList> UnitList = new List<ClassUnitList>();

            /// <summary>
            /// 台リストクラス
            /// </summary>
            public class ClassUnitList
            {
                public string UnitNum;
                public List<ClassDailyData> DailyData = new List<ClassDailyData>();
            }
            /// <summary>
            /// デイリーデータクラス
            /// </summary>
            public class ClassDailyData
            {
                public DateTime DateTime;
                public List<ClassHistoryData> HistoryData = new List<ClassHistoryData>();
            }
            /// <summary>
            /// 履歴データクラス
            /// </summary>
            public class ClassHistoryData
            {
                public int RotateCount = 0;
                public int HitStatus = 0;
            }
        }
        /// <summary>
        /// 台データクラス(整形後)
        /// </summary>
        public class ClassShapeUnitData
        {
            public string ModelName;
            public List<ClassShapeUnitList> UnitList = new List<ClassShapeUnitList>();

            /// <summary>
            /// 台リストクラス
            /// </summary>
            public class ClassShapeUnitList
            {
                public string UnitNum;
                public int RemainRotateCount;
                public List<ClassHistoryData> HistoryData = new List<ClassHistoryData>();
            }
        }
        /// <summary>
        /// 台データクラス(振り分け用)
        /// </summary>
        public class ClassTempUnitData
        {
            public int FirstHitIndex = 0;
            public int FirstHitCount = 0;
            public int ProbVarHitCount = 0;
            public int ProbVarFirstHitCount = 0;
            public int AllRotateCount = 0;
        }
        /// <summary>
        /// グラフデータクラス
        /// </summary>
        public class ClassGraphData
        {
            public string UnitNum;
            public List<ClassGraphAreaData> GraphAreaDataList = new List<ClassGraphAreaData>();

            /// <summary>
            /// グラフ範囲データクラス
            /// </summary>
            public class ClassGraphAreaData
            {
                public int GraphKind;
                public ClassCellRangeData FirstRange = new ClassCellRangeData();
                public ClassCellRangeData LastRange = new ClassCellRangeData();
            }
            /// <summary>
            /// セル位置データクラス
            /// </summary>
            public class ClassCellRangeData
            {
                public int row;
                public int column;
            }
        }
        /// <summary>
        /// フォームロードイベント
        /// </summary>
        private void FormMain_Load(object sender, EventArgs e)
        {
            try
            {
                //設定ファイルから機種情報取得
                ModelInfoList = GetModelInfo();

                //コンボボックス設定
                ComboBoxModelName.DropDownStyle = ComboBoxStyle.DropDownList;

                //DataGridView設定
                DataGridViewUnitData.EnableHeadersVisualStyles = false;
                DataGridViewUnitData.ColumnHeadersDefaultCellStyle.BackColor = System.Drawing.Color.White;
                DataGridViewUnitData.ColumnHeadersDefaultCellStyle.BackColor = System.Drawing.Color.MidnightBlue;
                DataGridViewUnitData.ReadOnly = true;
            }
            catch (Exception ex)
            {
                WinModuleLibrary.ErrorModule.ShowErrorLog(ex);
                return;
            }
        }
        /// <summary>
        /// 参照ボタンクリックイベント
        /// </summary>
        private void ButtonReference_Click(object sender, EventArgs e)
        {
            try
            {
                //OpenFileダイアログ表示
                OpenFileDialog Ofd = new OpenFileDialog();

                Ofd.FileName = "";
                Ofd.InitialDirectory = @"%UserProfile%\Documents";
                Ofd.Filter = "Excelブック(*.xlsx)|*.xlsx|Excelマクロ有効ブック(*.xlsm)|*.xlsm|すべてのファイル(*.*)|*.*";
                Ofd.FilterIndex = 1;
                Ofd.Title = "読込むファイルを選択してください";
                Ofd.RestoreDirectory = true;
                Ofd.CheckFileExists = true;
                Ofd.CheckPathExists = true;

                if (Ofd.ShowDialog() == DialogResult.OK)
                {
                    workbookPath = Ofd.FileName;
                    TextBoxReadDataPath.Text = Ofd.FileName;
                }
                return;
            }
            catch (Exception ex)
            {
                WinModuleLibrary.ErrorModule.ShowErrorLog(ex);
                return;
            }
        }
        /// <summary>
        /// Excelデータ読み込みボタンクリックイベント
        /// </summary>
        private void ButtonReadData_Click(object sender, EventArgs e)
        {
            try
            {
                if (File.Exists(workbookPath) == true)
                {
                    //エクセルから台データの取得
                    UnitDataList = ReadUnitDataFromExcel(workbookPath);

                    //台データの整形
                    ShapeUnitDataList = ShapingUnitData(UnitDataList);

                    //テキストボックスへ店舗名を追加
                    string workbookFileName = System.IO.Path.GetFileNameWithoutExtension(workbookPath);

                    TextBoxStoreName.Text = workbookFileName.Substring(workbookFileName.IndexOf("パチンコデータ_") + 8);

                    //コンボボックスへ機種名を追加
                    for (int i = 0; i < ShapeUnitDataList.Count; i++)
                    {
                        ComboBoxModelName.Items.Add(ShapeUnitDataList[i].ModelName);
                    }
                    ComboBoxModelName.SelectedIndex = 0;

                    //台スペック表示
                    for (int i = 0; i < ModelInfoList.Count; i++)
                    {
                        if(ModelInfoList[i].ModelName == ComboBoxModelName.Text)
                        {
                            TextBoxFirstHitProb.Text = String.Concat("1/", ModelInfoList[i].FirstHitProb);
                            TextBoxProbVarHitRashRate.Text = String.Concat(ModelInfoList[i].ProbVarHitRashRate, @"%");
                            TextBoxProbVarHitPersisRate.Text = String.Concat(ModelInfoList[i].ProbVarHitPersisRate, @"%");
                        }
                    }
                    if (TextBoxFirstHitProb.Text == "")
                    {
                        throw new Exception("台スペックの情報取得に失敗しました。");
                    }
                    //DataGridViewへ台データを表示
                    DispUnitDataInDataGridView(ShapeUnitDataList[0].UnitList);
                }
                else
                {
                    throw new Exception("Excelデータファイルの読み込みに失敗しました。");
                }
            }
            catch (Exception ex)
            {
                WinModuleLibrary.ErrorModule.ShowErrorLog(ex);
                return;
            }
        }
        /// <summary>
        /// データ解析/レポート出力ボタンクリックイベント
        /// </summary>
        private void ButtonAnalysData_Click(object sender, EventArgs e)
        {
            try
            {
                string analysisUnitDataDirPath = String.Concat(System.IO.Path.GetDirectoryName(workbookPath), @"\解析データ");

                Directory.CreateDirectory(analysisUnitDataDirPath);

                //台データ整形(直近初当たり回数100回の台データのみ抽出)
                ReShapeUnitDataList = ReShapingUnitData(ShapeUnitDataList);

                //Excelへデータ出力
                for (int i = 0; i < ReShapeUnitDataList.Count; i++)
                {
                    Excel.XLWorkbook Workbook = new Excel.XLWorkbook();

                    //Excelへ解析データ出力(Summaryシート)
                    Workbook = ExportUnitDataToExcel_Summary(ReShapeUnitDataList[i], Workbook);

                    //Excelへ解析データ出力(GraphDataシート)
                    Workbook = ExportUnitDataToExcel_GraqhData(ReShapeUnitDataList[i], Workbook);

                    //Excelを規程の配置場所へ保存
                    Workbook.SaveAs(string.Concat(analysisUnitDataDirPath, @"\パチンコ解析データ_", ReShapeUnitDataList[i].ModelName, ".xlsx"));

                    //EPPulsを使用
                    FileInfo WorkbookInfo = new FileInfo(string.Concat(analysisUnitDataDirPath, @"\パチンコ解析データ_", ReShapeUnitDataList[i].ModelName, ".xlsx"));

                    using (EPPExcel.ExcelPackage ExcelPackage = new EPPExcel.ExcelPackage(WorkbookInfo))
                    {
                        //デバッグモードでの例外回避
                        EPPExcel.ExcelPackage.LicenseContext = EPPExcel.LicenseContext.NonCommercial;

                        EPPExcel.ExcelWorkbook EPPWorkbook = ExcelPackage.Workbook;

                        //Excelへ解析データ出力(Graphシート)
                        EPPWorkbook = ExportUnitDataToExcel_Graqh(ReShapeUnitDataList[i], ExcelPackage.Workbook);

                        //Excelを規程の配置場所へ保存
                        ExcelPackage.Save();
                    }
                }
                MessageBox.Show("パチンコ解析データ出力完了", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                WinModuleLibrary.ErrorModule.ShowErrorLog(ex);
                return;
            }
        }
        /// <summary>
        /// コンボボックス選択アイテム変更イベント
        /// </summary>
        private void ComboBoxModelName_SelectionChangeCommitted(object sender, EventArgs e)
        {
            if (ShapeUnitDataList != null)
            {
                for (int i = 0; i < ShapeUnitDataList.Count; i++)
                {
                    if (i == ComboBoxModelName.SelectedIndex)
                    {
                        //DataGridViewへ台データを表示
                        DispUnitDataInDataGridView(ShapeUnitDataList[i].UnitList);
                    }
                }
                //台スペック表示
                for (int i = 0; i < ModelInfoList.Count; i++)
                {
                    if (ModelInfoList[i].ModelName == ComboBoxModelName.Text)
                    {
                        TextBoxFirstHitProb.Text = String.Concat("1/", ModelInfoList[i].FirstHitProb);
                        TextBoxProbVarHitRashRate.Text = String.Concat(ModelInfoList[i].ProbVarHitRashRate, @"%");
                        TextBoxProbVarHitPersisRate.Text = String.Concat(ModelInfoList[i].ProbVarHitPersisRate, @"%");
                    }
                }
                if (TextBoxFirstHitProb.Text == "")
                {
                    throw new Exception("台スペックの情報取得に失敗しました。");
                }
                return;
            }
            else
            {
                throw new Exception("台データが読込まれていません。");
            }
        }
        /// <summary>
        /// 設定ファイルから機種情報取得
        /// </summary>
        private static List<ClassModelInfo> GetModelInfo()
        {
            List<ClassModelInfo> ModelInfoList = new List<ClassModelInfo>();
            string modelInfoSettingFilePath = String.Concat(Directory.GetCurrentDirectory(), @"\ModelInfoSetting.ini");

            if (File.Exists(modelInfoSettingFilePath) == false)
            {
                throw new Exception("機種情報設定ファイルの取得に失敗しました。");
            }
            using (StreamReader sr = new StreamReader(modelInfoSettingFilePath, System.Text.Encoding.UTF8))
            {
                ClassModelInfo ModelInfo = new ClassModelInfo();

                while (sr.EndOfStream == false)
                {
                    string[] SplitReadLine = sr.ReadLine().Split(':');

                    //機種名
                    if (SplitReadLine[0] == "ModelName")
                    {
                        ModelInfo.ModelName = SplitReadLine[1].ToString();
                    }
                    //初当たり確率
                    else if (SplitReadLine[0] == "FirstHitProb")
                    {
                        ModelInfo.FirstHitProb = SplitReadLine[1].ToString();
                    }
                    //確変突入率
                    else if (SplitReadLine[0] == "ProbVarHitRashRate")
                    {
                        ModelInfo.ProbVarHitRashRate = SplitReadLine[1].ToString();
                    }
                    //確変継続率
                    else if (SplitReadLine[0] == "ProbVarHitPersisRate")
                    {
                        ModelInfo.ProbVarHitPersisRate = SplitReadLine[1].ToString();
                        ModelInfoList.Add(ModelInfo);
                        ModelInfo = new ClassModelInfo();
                    }
                    else
                    {
                        continue;
                    }
                }
            }
            return ModelInfoList;
        }
        /// <summary>
        /// エクセルから台データの取得
        /// </summary>
        private static List<ClassUnitData> ReadUnitDataFromExcel(string workbookPath)
        {
            Excel.XLWorkbook Workbook = new Excel.XLWorkbook(workbookPath);
            List<ClassUnitData> UnitDataList = new List<ClassUnitData>();

            for (int i = 1; i <= Workbook.Worksheets.Count; i++)
            {
                //デバッグ用
                if (i == Workbook.Worksheets.Count)
                {
                    continue;
                }
                Excel.IXLWorksheet Worksheet = Workbook.Worksheet(i);
                ClassUnitData UnitData = new ClassUnitData();
                int unitCount = 0;

                for (int j = 1; j <= Worksheet.RowCount(); j++)
                {
                    ClassUnitList UnitList = new ClassUnitList();
                    bool endFlg = false;

                    for (int k = 1; k <= Worksheet.ColumnCount(); k++)
                    {
                        if (k > 1)
                        {
                            if (k == 2 && Worksheet.Cell(j, k).Value.ToString() != "")
                            {
                                //機種名
                                if (Worksheet.Cell(j, k).Value.ToString() == "機種名")
                                {
                                    if (Worksheet.Cell(j, k + 1).Value.ToString() != null
                                        || Worksheet.Cell(j, k + 1).Value.ToString() != "")
                                    {
                                        UnitData.ModelName = Worksheet.Cell(j, k + 1).Value.ToString();
                                        break;
                                    }
                                    else
                                    {
                                        throw new Exception(string.Concat("セル(", ConvNumToAlphabet(k + 1), j, ")の機種名の取得に失敗しました。"));
                                    }
                                }
                                //設置台数
                                else if (Worksheet.Cell(j, k).Value.ToString() == "設置台数")
                                {
                                    if (int.TryParse(Worksheet.Cell(j, k + 1).Value.ToString(), out int installNum) == true)
                                    {
                                        UnitData.InstallNum = installNum;
                                        j++;
                                        break;
                                    }
                                    else
                                    {
                                        throw new Exception(string.Concat("セル(", ConvNumToAlphabet(k + 1), j, ")の設置台数の取得に失敗しました。"));
                                    }
                                }
                                else
                                {
                                    //台番号
                                    if (Regex.IsMatch(Worksheet.Cell(j, k).Value.ToString(), "^\\d+$") == true)
                                    {
                                        UnitList.UnitNum = Worksheet.Cell(j, k).Value.ToString();
                                        unitCount++;
                                        continue;
                                    }
                                    else
                                    {
                                        throw new Exception(string.Concat("セル(", ConvNumToAlphabet(k), j, ")の台番号の取得に失敗しました。"));
                                    }
                                }
                            }
                            else
                            {
                                if (DateTime.TryParse(Worksheet.Cell(j, k).Value.ToString(), out DateTime dateTime) == true)
                                {
                                    //回転数と大当りステータスの取得
                                    ClassDailyData DailyData = GetRotateCountAndHitStatus(Worksheet, j, k, dateTime);

                                    //日ごとの履歴データをリストへ追加
                                    UnitList.DailyData.Add(DailyData);
                                    k++;
                                    continue;
                                }
                                else
                                {
                                    if (Worksheet.Cell(j + 1, k).Value.ToString() == "回転数"
                                        || Worksheet.Cell(j + 1, k + 1).Value.ToString() == "ステータス")
                                    {
                                        throw new Exception(string.Concat("セル(", ConvNumToAlphabet(k), j, ")の日付情報の取得に失敗しました。"));
                                    }
                                    else
                                    {
                                        if (UnitList.UnitNum != null && UnitList.DailyData.Count != 0)
                                        {
                                            //1台毎の台データをリストへ追加
                                            UnitData.UnitList.Add(UnitList);

                                            //最終台のデータ取得が完了するまで探索
                                            if (UnitData.InstallNum - unitCount != 0)
                                            {
                                                j += 101;
                                            }
                                            else
                                            {
                                                endFlg = true;
                                            }
                                        }
                                        break;
                                    }
                                }
                            }
                        }
                    }
                    if (endFlg == true)
                    {
                        break;
                    }
                }
                UnitDataList.Add(UnitData);
            }
            return UnitDataList;
        }
        /// <summary>
        /// 回転数と大当りステータスの取得
        /// </summary>
        private static ClassDailyData GetRotateCountAndHitStatus(Excel.IXLWorksheet Worksheet, int rowCount, int columnCount, DateTime dateTime)
        {
            ClassDailyData DailyData = new ClassDailyData();

            //日付
            DailyData.DateTime = dateTime;

            if (Worksheet.Cell(rowCount + 1, columnCount).Value.ToString() == "回転数"
                && Worksheet.Cell(rowCount + 1, columnCount + 1).Value.ToString() == "ステータス")
            {
                ClassHistoryData HistoryData;
                bool convRotCntResult = false;
                bool convStatusResult = false;
                int rotateCount = 0;
                int status = 0;

                rowCount += 2;

                //ステータス(0)でない場合は探索
                while (int.Parse(Worksheet.Cell(rowCount, columnCount + 1).Value.ToString()) != 0)
                {
                    HistoryData = new ClassHistoryData();
                    convRotCntResult = int.TryParse(Worksheet.Cell(rowCount, columnCount).Value.ToString(), out rotateCount);
                    convStatusResult = int.TryParse(Worksheet.Cell(rowCount, columnCount + 1).Value.ToString(), out status);

                    //回転数
                    if (convRotCntResult == true)
                    {
                        HistoryData.RotateCount = rotateCount;
                    }
                    else
                    {
                        throw new Exception(string.Concat("セル(", ConvNumToAlphabet(columnCount), rowCount, ")の回転数情報の取得に失敗しました。"));
                    }
                    //ステータス
                    if (convStatusResult == true)
                    {
                        HistoryData.HitStatus = status;
                    }
                    else
                    {
                        throw new Exception(string.Concat("セル(", ConvNumToAlphabet(columnCount + 1), rowCount, ")のステータス情報の取得に失敗しました。"));
                    }
                    //リストに追加
                    DailyData.HistoryData.Add(HistoryData);
                    rowCount++;
                }
                HistoryData = new ClassHistoryData();

                //ステータス(0)を1行だけ取得
                convRotCntResult = int.TryParse(Worksheet.Cell(rowCount, columnCount).Value.ToString(), out rotateCount);
                convStatusResult = int.TryParse(Worksheet.Cell(rowCount, columnCount + 1).Value.ToString(), out status);

                //回転数
                if (convRotCntResult == true)
                {
                    HistoryData.RotateCount = rotateCount;
                }
                else
                {
                    throw new Exception(string.Concat("セル(", ConvNumToAlphabet(columnCount), rowCount, ")の回転数情報の取得に失敗しました。"));
                }
                //ステータス
                if (convStatusResult == true)
                {
                    HistoryData.HitStatus = status;
                }
                else
                {
                    throw new Exception(string.Concat("セル(", ConvNumToAlphabet(columnCount + 1), rowCount, ")のステータス情報の取得に失敗しました。"));
                }
                //リストに追加
                DailyData.HistoryData.Add(HistoryData);
            }
            else
            {
                throw new Exception(string.Concat("セル(", ConvNumToAlphabet(columnCount), rowCount, ")の値が不正です。"));
            }
            return DailyData;
        }
        /// <summary>
        /// Excelのカラム番号をアルファベットへ変換
        /// </summary>
        public static string ConvNumToAlphabet(int columnNum)
        {
            string alphabet = string.Empty;

            if (columnNum > 0)
            {
                while (columnNum > 0)
                {
                    columnNum--;
                    alphabet = Convert.ToChar(columnNum % 26 + 65) + alphabet;
                    columnNum = columnNum / 26;
                }
            }
            return alphabet;
        }
        /// <summary>
        /// 台データの整形
        /// </summary>
        private static List<ClassShapeUnitData> ShapingUnitData(List<ClassUnitData> UnitDataList)
        {
            List<ClassShapeUnitData> ShapeUnitDataList = new List<ClassShapeUnitData>();

            for (int i = 0; i < UnitDataList.Count; i++)
            {
                ClassShapeUnitData ShapeUnitData = new ClassShapeUnitData();

                ShapeUnitData.ModelName = UnitDataList[i].ModelName;

                for (int j = 0; j < UnitDataList[i].UnitList.Count; j++)
                {
                    ClassShapeUnitList ShapeUnitList = new ClassShapeUnitList();
                    bool TakeoverPrevDataFlg = false;
                    int FinalRotateCount = 0;

                    ShapeUnitList.UnitNum = UnitDataList[i].UnitList[j].UnitNum;

                    for (int k = 0; k < UnitDataList[i].UnitList[j].DailyData.Count; k++)
                    {
                        for (int l = 0; l < UnitDataList[i].UnitList[j].DailyData[k].HistoryData.Count; l++)
                        {
                            //残スタート数
                            if (UnitDataList[i].UnitList[j].DailyData[k].HistoryData[l].HitStatus == 0)
                            {
                                //残スタートが連続する場合は加算
                                if (TakeoverPrevDataFlg == true)
                                {
                                    FinalRotateCount += UnitDataList[i].UnitList[j].DailyData[k].HistoryData[l].RotateCount;
                                }
                                else
                                {
                                    FinalRotateCount = UnitDataList[i].UnitList[j].DailyData[k].HistoryData[l].RotateCount;
                                }
                                TakeoverPrevDataFlg = true;

                                //最終日の最後尾データは残スタート数とする
                                if (k == UnitDataList[i].UnitList[j].DailyData.Count - 1
                                    && l == UnitDataList[i].UnitList[j].DailyData[k].HistoryData.Count - 1)
                                {
                                    ShapeUnitList.RemainRotateCount = FinalRotateCount;
                                }
                            }
                            //初当たりと確変
                            else
                            {
                                ClassHistoryData HistoryData = new ClassHistoryData();

                                //残スタート数を次の日の回転数へ加算
                                if (TakeoverPrevDataFlg == true
                                    && UnitDataList[i].UnitList[j].DailyData[k].HistoryData[l].HitStatus == 1)
                                {
                                    HistoryData.RotateCount = UnitDataList[i].UnitList[j].DailyData[k].HistoryData[l].RotateCount + FinalRotateCount;
                                    HistoryData.HitStatus = UnitDataList[i].UnitList[j].DailyData[k].HistoryData[l].HitStatus;
                                    TakeoverPrevDataFlg = false;
                                }
                                else
                                {
                                    HistoryData.RotateCount = UnitDataList[i].UnitList[j].DailyData[k].HistoryData[l].RotateCount;
                                    HistoryData.HitStatus = UnitDataList[i].UnitList[j].DailyData[k].HistoryData[l].HitStatus;
                                }
                                ShapeUnitList.HistoryData.Add(HistoryData);
                            }
                        }
                    }
                    ShapeUnitData.UnitList.Add(ShapeUnitList);
                }
                ShapeUnitDataList.Add(ShapeUnitData);
            }
            return ShapeUnitDataList;
        }
        /// <summary>
        /// 台データ整形(直近初当たり回数100回の台データのみ抽出)
        /// </summary>
        private static List<ClassShapeUnitData> ReShapingUnitData(List<ClassShapeUnitData> ShapeUnitDataList)
        {
            List<ClassShapeUnitData> ReShapeUnitDataList = new List<ClassShapeUnitData>();

            for (int i = 0; i < ShapeUnitDataList.Count; i++)
            {
                ClassShapeUnitData ReShapeUnitData = new ClassShapeUnitData();

                //機種名
                ReShapeUnitData.ModelName = ShapeUnitDataList[i].ModelName;

                for (int j = 0; j < ShapeUnitDataList[i].UnitList.Count; j++)
                {
                    ClassShapeUnitList ReShapeUnitList = new ClassShapeUnitList();

                    //台番号
                    ReShapeUnitList.UnitNum = ShapeUnitDataList[i].UnitList[j].UnitNum;

                    //残スタート回転数
                    ReShapeUnitList.RemainRotateCount = ShapeUnitDataList[i].UnitList[j].RemainRotateCount;

                    for (int k = 0; k < ShapeUnitDataList[i].UnitList[j].HistoryData.Count; k++)
                    {
                        ClassHistoryData HistoryData = new ClassHistoryData();

                        //回転数と大当りステータス
                        HistoryData.RotateCount = ShapeUnitDataList[i].UnitList[j].HistoryData[k].RotateCount;
                        HistoryData.HitStatus = ShapeUnitDataList[i].UnitList[j].HistoryData[k].HitStatus;

                        ReShapeUnitList.HistoryData.Add(HistoryData);
                    }
                    ReShapeUnitData.UnitList.Add(ReShapeUnitList);
                }
                ReShapeUnitDataList.Add(ReShapeUnitData);
            }
            return ReShapeUnitDataList;
        }
        /// <summary>
        /// DataGridViewへ台データを表示
        /// </summary>
        private void DispUnitDataInDataGridView(List<ClassShapeUnitList> ShapeUnitList)
        {
            //DataGridView内データを初期化
            DataGridViewUnitData.Rows.Clear();

            for (int i = 0; i < ShapeUnitList.Count; i++)
            {
                DataGridViewUnitData.Rows.Add();
                DataGridViewUnitData.Rows[i].Cells[0].Value = ShapeUnitList[i].UnitNum;
                DataGridViewUnitData.Rows[i].Cells[4].Value = ShapeUnitList[i].RemainRotateCount;

                int firstHitCount = 0;
                int probVarHitCount = 0;
                int allRotateCount = 0;

                for (int j = 0; j < ShapeUnitList[i].HistoryData.Count; j++)
                {
                    switch (ShapeUnitList[i].HistoryData[j].HitStatus)
                    {
                        case 1:
                            firstHitCount++;
                            break;
                        case 2:
                            probVarHitCount++;
                            break;
                    }
                    allRotateCount += ShapeUnitList[i].HistoryData[j].RotateCount;
                }
                DataGridViewUnitData.Rows[i].Cells[1].Value = firstHitCount.ToString();
                DataGridViewUnitData.Rows[i].Cells[2].Value = probVarHitCount.ToString();
                DataGridViewUnitData.Rows[i].Cells[3].Value = allRotateCount.ToString();
            }
            return;
        }
        /// <summary>
        /// Excelへ解析データ出力(Summaryシート)
        /// </summary>
        private static Excel.XLWorkbook ExportUnitDataToExcel_Summary(ClassShapeUnitData ShapeUnitData, Excel.XLWorkbook Workbook)
        {
            Excel.IXLWorksheet Worksheet = Workbook.Worksheets.Add(string.Concat("Summary"));
            int endRow = 0;

            //列幅調整
            Worksheet.Column(2).Width = 8.43;

            for (int i = 3; i <= 8; i++)
            {
                Worksheet.Column(i).Width = 16;
            }
            Worksheet.Cell(2, 2).Value = "機種名";
            Worksheet.Cell(4, 2).Value = "台番号";
            Worksheet.Cell(4, 3).Value = "初当り回数";
            Worksheet.Cell(4, 4).Value = "確変回数";
            Worksheet.Cell(4, 5).Value = "総回転数";
            Worksheet.Cell(4, 6).Value = "残り回転数";
            Worksheet.Cell(4, 7).Value = "初当り合成確率";
            Worksheet.Cell(4, 8).Value = "確変突入率";

            Worksheet.Cell(2, 2).Style.Alignment.Horizontal = Excel.XLAlignmentHorizontalValues.Center;
            Worksheet.Cell(2, 2).Style.Alignment.Vertical = Excel.XLAlignmentVerticalValues.Center;
            Worksheet.Range(string.Concat(ConvNumToAlphabet(2), 2)).Style.Border.OutsideBorder = Excel.XLBorderStyleValues.Thin;

            for (int i = 2; i <= 8; i++)
            {
                Worksheet.Cell(4, i).Style.Alignment.Horizontal = Excel.XLAlignmentHorizontalValues.Center;
                Worksheet.Cell(4, i).Style.Alignment.Vertical = Excel.XLAlignmentVerticalValues.Center;
            }
            //機種名
            Worksheet.Cell(2, 3).Value = ShapeUnitData.ModelName;
            Worksheet.Range(string.Concat(ConvNumToAlphabet(3), 2, ":", ConvNumToAlphabet(7), 2)).Merge();
            Worksheet.Cell(2, 3).Style.Alignment.Horizontal = Excel.XLAlignmentHorizontalValues.Center;
            Worksheet.Cell(2, 3).Style.Alignment.Vertical = Excel.XLAlignmentVerticalValues.Center;
            Worksheet.Range(string.Concat(ConvNumToAlphabet(3), 2, ":", ConvNumToAlphabet(7), 2)).Style.Border.OutsideBorder = Excel.XLBorderStyleValues.Thin;

            for (int i = 0; i < ShapeUnitData.UnitList.Count; i++)
            {
                int firstHitCount = 0;
                int probVarHitCount = 0;
                int probVarFirstHitCount = 0;
                int allRotateCount = 0;
                int totalRotateCount = 0;

                //台番号
                Worksheet.Cell(i + 5, 2).Value = ShapeUnitData.UnitList[i].UnitNum.ToString();
                Worksheet.Cell(i + 5, 2).Style.Alignment.Horizontal = Excel.XLAlignmentHorizontalValues.Center;
                Worksheet.Cell(i + 5, 2).Style.Alignment.Vertical = Excel.XLAlignmentVerticalValues.Center;

                for (int j = 0; j < ShapeUnitData.UnitList[i].HistoryData.Count; j++)
                {
                    switch (ShapeUnitData.UnitList[i].HistoryData[j].HitStatus)
                    {
                        case 1:
                            firstHitCount++;

                            if (j < ShapeUnitData.UnitList[i].HistoryData.Count - 1
                                && ShapeUnitData.UnitList[i].HistoryData[j + 1].HitStatus == 2)
                            {
                                probVarFirstHitCount++;
                            }
                            allRotateCount += ShapeUnitData.UnitList[i].HistoryData[j].RotateCount;

                            break;
                        case 2:
                            probVarHitCount++;
                            break;
                    }
                    totalRotateCount += ShapeUnitData.UnitList[i].HistoryData[j].RotateCount;
                }
                //初当たり回数
                Worksheet.Cell(i + 5, 3).Value = firstHitCount;
                Worksheet.Cell(i + 5, 3).Style.Alignment.Horizontal = Excel.XLAlignmentHorizontalValues.Right;
                Worksheet.Cell(i + 5, 3).Style.Alignment.Vertical = Excel.XLAlignmentVerticalValues.Center;

                //確変回数
                Worksheet.Cell(i + 5, 4).Value = probVarHitCount;
                Worksheet.Cell(i + 5, 4).Style.Alignment.Horizontal = Excel.XLAlignmentHorizontalValues.Right;
                Worksheet.Cell(i + 5, 4).Style.Alignment.Vertical = Excel.XLAlignmentVerticalValues.Center;

                //総回転数
                Worksheet.Cell(i + 5, 5).Value = totalRotateCount;
                Worksheet.Cell(i + 5, 5).Style.Alignment.Horizontal = Excel.XLAlignmentHorizontalValues.Right;
                Worksheet.Cell(i + 5, 5).Style.Alignment.Vertical = Excel.XLAlignmentVerticalValues.Center;

                //残りスタート数
                Worksheet.Cell(i + 5, 6).Value = ShapeUnitData.UnitList[i].RemainRotateCount;
                Worksheet.Cell(i + 5, 6).Style.Alignment.Horizontal = Excel.XLAlignmentHorizontalValues.Right;
                Worksheet.Cell(i + 5, 6).Style.Alignment.Vertical = Excel.XLAlignmentVerticalValues.Center;

                //0除算防止
                if (firstHitCount != 0)
                {
                    //初当り合成確率
                    Worksheet.Cell(i + 5, 7).Value = allRotateCount / firstHitCount;
                    Worksheet.Cell(i + 5, 7).Style.Alignment.Horizontal = Excel.XLAlignmentHorizontalValues.Right;
                    Worksheet.Cell(i + 5, 7).Style.Alignment.Vertical = Excel.XLAlignmentVerticalValues.Center;

                    //確変突入率
                    Worksheet.Cell(i + 5, 8).Value = probVarFirstHitCount * 100 / firstHitCount;
                    Worksheet.Cell(i + 5, 8).Style.Alignment.Horizontal = Excel.XLAlignmentHorizontalValues.Right;
                    Worksheet.Cell(i + 5, 8).Style.Alignment.Vertical = Excel.XLAlignmentVerticalValues.Center;
                }
                else
                {
                    Worksheet.Cell(i + 5, 7).Value = allRotateCount;
                    Worksheet.Cell(i + 5, 7).Style.Alignment.Horizontal = Excel.XLAlignmentHorizontalValues.Right;
                    Worksheet.Cell(i + 5, 7).Style.Alignment.Vertical = Excel.XLAlignmentVerticalValues.Center;

                    Worksheet.Cell(i + 5, 8).Value = firstHitCount;
                    Worksheet.Cell(i + 5, 8).Style.Alignment.Horizontal = Excel.XLAlignmentHorizontalValues.Right;
                    Worksheet.Cell(i + 5, 8).Style.Alignment.Vertical = Excel.XLAlignmentVerticalValues.Center;
                }
                if (i == ShapeUnitData.UnitList.Count - 1)
                {
                    endRow = i + 5;
                }
            }
            //罫線を描画
            string range = string.Concat(ConvNumToAlphabet(2), 4, ":", ConvNumToAlphabet(8), endRow);
            
            Worksheet.Range(range).Style.Border.OutsideBorder = Excel.XLBorderStyleValues.Thin;
            Worksheet.Range(range).Style.Border.InsideBorder = Excel.XLBorderStyleValues.Thin;

            range = string.Concat(ConvNumToAlphabet(2), 4, ":", ConvNumToAlphabet(8), 4);
            Worksheet.Range(range).Style.Border.BottomBorder = Excel.XLBorderStyleValues.Double;

            return Workbook;
        }
        /// <summary>
        /// Excelへ解析データ出力(GraphDataシート)
        /// </summary>
        private static Excel.XLWorkbook ExportUnitDataToExcel_GraqhData(ClassShapeUnitData ShapeUnitData, Excel.XLWorkbook Workbook)
        {
            #region 修正後
            Excel.IXLWorksheet WorksheetGraqh = Workbook.Worksheets.Add(string.Concat("Graqh"));
            Excel.IXLWorksheet WorksheetGraqhData = Workbook.Worksheets.Add(string.Concat("GraqhData"));
            GraphDataList = new List<ClassGraphData>();

            int row = 2;

            //列幅調整
            WorksheetGraqhData.Column(2).Width = 8.43;
            WorksheetGraqhData.Column(6).Width = 8.43;

            for (int i = 3; i <= 4; i++)
            {
                WorksheetGraqhData.Column(i).Width = 16;
            }
            for (int i = 7; i <= 9; i++)
            {
                WorksheetGraqhData.Column(i).Width = 16;
            }

            for (int i = 0; i < ShapeUnitData.UnitList.Count; i++)
            {
                ClassGraphData GraphAreaData = new ClassGraphData();
                ClassGraphAreaData GraphAreaData_FirstHitProb = new ClassGraphAreaData();
                ClassGraphAreaData GraphAreaData_ProbVarHitRashRate = new ClassGraphAreaData();
                List<int> FirstHitIndexList = new List<int>();

                int firstHitCount = 0;
                int ResetHitCount = 0;
                int probVarHitCount = 0;
                int probVarFirstHitCount = 0;
                int allRotateCount = 0;

                for (int j = 0; j < ShapeUnitData.UnitList[i].HistoryData.Count; j++)
                {
                    if (ShapeUnitData.UnitList[i].HistoryData[j].HitStatus == 1)
                    {
                        FirstHitIndexList.Add(j);
                    }
                }
                //グラフ種別
                GraphAreaData_FirstHitProb.GraphKind = 1;
                GraphAreaData_ProbVarHitRashRate.GraphKind = 2;

                //台番号
                WorksheetGraqhData.Cell(row, 1).Value = ShapeUnitData.UnitList[i].UnitNum;
                GraphAreaData.UnitNum = ShapeUnitData.UnitList[i].UnitNum;

                GraphAreaData_FirstHitProb.FirstRange.row = row;
                GraphAreaData_FirstHitProb.FirstRange.column = 2;
                GraphAreaData_ProbVarHitRashRate.FirstRange.row = row;
                GraphAreaData_ProbVarHitRashRate.FirstRange.column = 6;

                WorksheetGraqhData.Cell(row, 3).Value = "基準値";
                WorksheetGraqhData.Cell(row, 4).Value = "合成確率";
                WorksheetGraqhData.Cell(row, 7).Value = "基準値";
                WorksheetGraqhData.Cell(row, 8).Value = "確変突入率";
                WorksheetGraqhData.Cell(row, 9).Value = "確変回数";
                row++;

                //初当たり0回目のデータ表示
                WorksheetGraqhData.Cell(row, 2).Value = 0;
                WorksheetGraqhData.Cell(row, 6).Value = 0;
                WorksheetGraqhData.Cell(row, 9).Value = 0;

                for (int k = 0; k < ModelInfoList.Count; k++)
                {
                    if (ModelInfoList[k].ModelName == ShapeUnitData.ModelName)
                    {
                        WorksheetGraqhData.Cell(row, 3).Value = int.Parse(ModelInfoList[k].FirstHitProb);
                        WorksheetGraqhData.Cell(row, 4).Value = int.Parse(ModelInfoList[k].FirstHitProb);
                        WorksheetGraqhData.Cell(row, 7).Value = int.Parse(ModelInfoList[k].ProbVarHitRashRate);
                        WorksheetGraqhData.Cell(row, 8).Value = int.Parse(ModelInfoList[k].ProbVarHitRashRate);
                    }
                }
                row++;

                //初当たり回数が5回未満
                if (FirstHitIndexList.Count < 5)
                {
                    WorksheetGraqhData.Cell(row, 2).Value = FirstHitIndexList.Count;
                    WorksheetGraqhData.Cell(row, 6).Value = FirstHitIndexList.Count;

                    for (int k = 0; k < ModelInfoList.Count; k++)
                    {
                        if (ModelInfoList[k].ModelName == ShapeUnitData.ModelName)
                        {
                            WorksheetGraqhData.Cell(row, 3).Value = int.Parse(ModelInfoList[k].FirstHitProb);
                            WorksheetGraqhData.Cell(row, 7).Value = int.Parse(ModelInfoList[k].ProbVarHitRashRate);
                        }
                    }
                    for (int j = 0; j < ShapeUnitData.UnitList[i].HistoryData.Count; j++)
                    {
                        switch (ShapeUnitData.UnitList[i].HistoryData[j].HitStatus)
                        {
                            case 1:
                                firstHitCount++;

                                if (j < ShapeUnitData.UnitList[i].HistoryData.Count - 1
                                    && ShapeUnitData.UnitList[i].HistoryData[j + 1].HitStatus == 2)
                                {
                                    probVarFirstHitCount++;
                                }
                                allRotateCount += ShapeUnitData.UnitList[i].HistoryData[j].RotateCount;

                                break;
                            case 2:
                                probVarHitCount++;
                                break;
                        }
                    }
                    //確変回数
                    WorksheetGraqhData.Cell(row, 9).Value = probVarHitCount;

                    //0除算防止
                    if (firstHitCount != 0)
                    {
                        //初当り合成確率
                        WorksheetGraqhData.Cell(row, 4).Value = allRotateCount / firstHitCount;

                        //確変突入率
                        WorksheetGraqhData.Cell(row, 8).Value = probVarFirstHitCount * 100 / firstHitCount;
                    }
                    else
                    {
                        WorksheetGraqhData.Cell(row, 4).Value = allRotateCount;
                        WorksheetGraqhData.Cell(row, 8).Value = firstHitCount;
                    }
                    row++;
                }
                //初当たり回数が5回以上
                else
                {
                    List<ClassTempUnitData> TempUnitDataList = new List<ClassTempUnitData>();

                    int remainDivFirstHitCount = FirstHitIndexList.Count % 5; //ここ
                    int divFirstHitCount = 5; //ここ

                    //初当たり10回未満は先頭で計算
                    if (remainDivFirstHitCount != 0)
                    {
                        divFirstHitCount = remainDivFirstHitCount;
                    }
                    for (int j = 0; j < ShapeUnitData.UnitList[i].HistoryData.Count; j++)
                    {
                        switch (ShapeUnitData.UnitList[i].HistoryData[j].HitStatus)
                        {
                            case 1:
                                ResetHitCount++;
                                firstHitCount++;

                                if (j < ShapeUnitData.UnitList[i].HistoryData.Count - 1
                                    && ShapeUnitData.UnitList[i].HistoryData[j + 1].HitStatus == 2)
                                {
                                    probVarFirstHitCount++;
                                }
                                allRotateCount += ShapeUnitData.UnitList[i].HistoryData[j].RotateCount;

                                break;
                            case 2:
                                probVarHitCount++;
                                break;
                        }
                        //初当たり10回ごとの最終データ取得時
                        if (ResetHitCount > divFirstHitCount || j == ShapeUnitData.UnitList[i].HistoryData.Count - 1)
                        {
                            ClassTempUnitData TempUnitData = new ClassTempUnitData();

                            TempUnitData.FirstHitIndex = divFirstHitCount;
                            TempUnitData.FirstHitCount = firstHitCount - 1;
                            TempUnitData.ProbVarHitCount = probVarHitCount;

                            if (j < ShapeUnitData.UnitList[i].HistoryData.Count - 1
                                && ShapeUnitData.UnitList[i].HistoryData[j + 1].HitStatus == 2)
                            {
                                TempUnitData.ProbVarFirstHitCount = probVarFirstHitCount - 1;
                            }
                            else
                            {
                                TempUnitData.ProbVarFirstHitCount = probVarFirstHitCount;
                            }
                            TempUnitData.AllRotateCount = allRotateCount;
                            TempUnitDataList.Add(TempUnitData);

                            //リセット
                            firstHitCount = 1;
                            probVarHitCount = 0;

                            if (j < ShapeUnitData.UnitList[i].HistoryData.Count - 1
                                && ShapeUnitData.UnitList[i].HistoryData[j + 1].HitStatus == 2)
                            {
                                probVarFirstHitCount = 1;
                            }
                            else
                            {
                                probVarFirstHitCount = 0;
                            }
                            allRotateCount = ShapeUnitData.UnitList[i].HistoryData[j].RotateCount;
                            divFirstHitCount += 5; //ここ
                        }
                    }
                    for (int j = 0; j < TempUnitDataList.Count; j++)
                    {
                        WorksheetGraqhData.Cell(row, 2).Value = TempUnitDataList[j].FirstHitIndex;
                        WorksheetGraqhData.Cell(row, 6).Value = TempUnitDataList[j].FirstHitIndex;

                        for (int k = 0; k < ModelInfoList.Count; k++)
                        {
                            if (ModelInfoList[k].ModelName == ShapeUnitData.ModelName)
                            {
                                WorksheetGraqhData.Cell(row, 3).Value = int.Parse(ModelInfoList[k].FirstHitProb);
                                WorksheetGraqhData.Cell(row, 7).Value = int.Parse(ModelInfoList[k].ProbVarHitRashRate);
                            }
                        }
                        //確変回数
                        WorksheetGraqhData.Cell(row, 9).Value = TempUnitDataList[j].ProbVarHitCount;

                        //0除算防止
                        if (TempUnitDataList[j].FirstHitCount != 0)
                        {
                            //初当り合成確率
                            WorksheetGraqhData.Cell(row, 4).Value = TempUnitDataList[j].AllRotateCount / TempUnitDataList[j].FirstHitCount;

                            //確変突入率
                            WorksheetGraqhData.Cell(row, 8).Value = TempUnitDataList[j].ProbVarFirstHitCount * 100 / TempUnitDataList[j].FirstHitCount;
                        }
                        else
                        {
                            WorksheetGraqhData.Cell(row, 4).Value = TempUnitDataList[j].AllRotateCount;
                            WorksheetGraqhData.Cell(row, 8).Value = TempUnitDataList[j].FirstHitCount;
                        }
                        row++;
                    }
                }
                GraphAreaData_FirstHitProb.LastRange.row = row - 1;
                GraphAreaData_FirstHitProb.LastRange.column = 4;
                GraphAreaData_ProbVarHitRashRate.LastRange.row = row - 1;
                GraphAreaData_ProbVarHitRashRate.LastRange.column = 9;

                //罫線を描画
                WorksheetGraqhData.Range(string.Concat(ConvNumToAlphabet(GraphAreaData_FirstHitProb.FirstRange.column), GraphAreaData_FirstHitProb.FirstRange.row)).Style.Border.DiagonalBorder = Excel.XLBorderStyleValues.Thin;
                WorksheetGraqhData.Range(string.Concat(ConvNumToAlphabet(GraphAreaData_FirstHitProb.FirstRange.column), GraphAreaData_FirstHitProb.FirstRange.row)).Style.Border.DiagonalDown = true;
                WorksheetGraqhData.Range(string.Concat(ConvNumToAlphabet(GraphAreaData_ProbVarHitRashRate.FirstRange.column), GraphAreaData_ProbVarHitRashRate.FirstRange.row)).Style.Border.DiagonalBorder = Excel.XLBorderStyleValues.Thin;
                WorksheetGraqhData.Range(string.Concat(ConvNumToAlphabet(GraphAreaData_ProbVarHitRashRate.FirstRange.column), GraphAreaData_ProbVarHitRashRate.FirstRange.row)).Style.Border.DiagonalDown = true;
                WorksheetGraqhData.Range(string.Concat(ConvNumToAlphabet(GraphAreaData_FirstHitProb.FirstRange.column), GraphAreaData_FirstHitProb.FirstRange.row, ":", ConvNumToAlphabet(GraphAreaData_FirstHitProb.LastRange.column), GraphAreaData_FirstHitProb.LastRange.row)).Style.Border.InsideBorder = Excel.XLBorderStyleValues.Thin;
                WorksheetGraqhData.Range(string.Concat(ConvNumToAlphabet(GraphAreaData_FirstHitProb.FirstRange.column), GraphAreaData_FirstHitProb.FirstRange.row, ":", ConvNumToAlphabet(GraphAreaData_FirstHitProb.LastRange.column), GraphAreaData_FirstHitProb.LastRange.row)).Style.Border.OutsideBorder = Excel.XLBorderStyleValues.Thin;
                WorksheetGraqhData.Range(string.Concat(ConvNumToAlphabet(GraphAreaData_ProbVarHitRashRate.FirstRange.column), GraphAreaData_ProbVarHitRashRate.FirstRange.row, ":", ConvNumToAlphabet(GraphAreaData_ProbVarHitRashRate.LastRange.column), GraphAreaData_ProbVarHitRashRate.LastRange.row)).Style.Border.InsideBorder = Excel.XLBorderStyleValues.Thin;
                WorksheetGraqhData.Range(string.Concat(ConvNumToAlphabet(GraphAreaData_ProbVarHitRashRate.FirstRange.column), GraphAreaData_ProbVarHitRashRate.FirstRange.row, ":", ConvNumToAlphabet(GraphAreaData_ProbVarHitRashRate.LastRange.column), GraphAreaData_ProbVarHitRashRate.LastRange.row)).Style.Border.OutsideBorder = Excel.XLBorderStyleValues.Thin;

                GraphAreaData.GraphAreaDataList.Add(GraphAreaData_FirstHitProb);
                GraphAreaData.GraphAreaDataList.Add(GraphAreaData_ProbVarHitRashRate);
                GraphDataList.Add(GraphAreaData);
                row++;
            }
            #endregion
            #region 修正前
            //Excel.IXLWorksheet WorksheetGraqh = Workbook.Worksheets.Add(string.Concat("Graqh"));
            //Excel.IXLWorksheet WorksheetGraqhData = Workbook.Worksheets.Add(string.Concat("GraqhData"));
            //GraphDataList = new List<ClassGraphData>();

            //int row = 2;

            ////列幅調整
            //WorksheetGraqhData.Column(2).Width = 8.43;
            //WorksheetGraqhData.Column(6).Width = 8.43;

            //for (int i = 3; i <= 4; i++)
            //{
            //    WorksheetGraqhData.Column(i).Width = 16;
            //}
            //for (int i = 7; i <= 9; i++)
            //{
            //    WorksheetGraqhData.Column(i).Width = 16;
            //}

            //for (int i = 0; i < ShapeUnitData.UnitList.Count; i++)
            //{
            //    ClassGraphData GraphAreaData = new ClassGraphData();
            //    ClassGraphAreaData GraphAreaData_FirstHitProb = new ClassGraphAreaData();
            //    ClassGraphAreaData GraphAreaData_ProbVarHitRashRate = new ClassGraphAreaData();
            //    List<int> FirstHitIndexList = new List<int>();

            //    int firstHitCount = 0;
            //    int ResetHitCount = 0;
            //    int probVarHitCount = 0;
            //    int probVarFirstHitCount = 0;
            //    int allRotateCount = 0;

            //    for (int j = 0; j < ShapeUnitData.UnitList[i].HistoryData.Count; j++)
            //    {
            //        if (ShapeUnitData.UnitList[i].HistoryData[j].HitStatus == 1)
            //        {
            //            FirstHitIndexList.Add(j);
            //        }
            //    }
            //    //グラフ種別
            //    GraphAreaData_FirstHitProb.GraphKind = 1;
            //    GraphAreaData_ProbVarHitRashRate.GraphKind = 2;

            //    //台番号
            //    WorksheetGraqhData.Cell(row, 1).Value = ShapeUnitData.UnitList[i].UnitNum;
            //    GraphAreaData.UnitNum = ShapeUnitData.UnitList[i].UnitNum;

            //    GraphAreaData_FirstHitProb.FirstRange.row = row;
            //    GraphAreaData_FirstHitProb.FirstRange.column = 2;
            //    GraphAreaData_ProbVarHitRashRate.FirstRange.row = row;
            //    GraphAreaData_ProbVarHitRashRate.FirstRange.column = 6;

            //    WorksheetGraqhData.Cell(row, 3).Value = "基準値";
            //    WorksheetGraqhData.Cell(row, 4).Value = "合成確率";
            //    WorksheetGraqhData.Cell(row, 7).Value = "基準値";
            //    WorksheetGraqhData.Cell(row, 8).Value = "確変突入率";
            //    WorksheetGraqhData.Cell(row, 9).Value = "確変回数";
            //    row++;

            //    //初当たり0回目のデータ表示
            //    WorksheetGraqhData.Cell(row, 2).Value = 0;
            //    WorksheetGraqhData.Cell(row, 6).Value = 0;
            //    WorksheetGraqhData.Cell(row, 9).Value = 0;

            //    for (int k = 0; k < ModelInfoList.Count; k++)
            //    {
            //        if (ModelInfoList[k].ModelName == ShapeUnitData.ModelName)
            //        {
            //            WorksheetGraqhData.Cell(row, 3).Value = int.Parse(ModelInfoList[k].FirstHitProb);
            //            WorksheetGraqhData.Cell(row, 4).Value = int.Parse(ModelInfoList[k].FirstHitProb);
            //            WorksheetGraqhData.Cell(row, 7).Value = int.Parse(ModelInfoList[k].ProbVarHitRashRate);
            //            WorksheetGraqhData.Cell(row, 8).Value = int.Parse(ModelInfoList[k].ProbVarHitRashRate);
            //        }
            //    }
            //    row++;

            //    //初当たり回数が10回未満
            //    if (FirstHitIndexList.Count < 10)
            //    {
            //        WorksheetGraqhData.Cell(row, 2).Value = FirstHitIndexList.Count;
            //        WorksheetGraqhData.Cell(row, 6).Value = FirstHitIndexList.Count;

            //        for (int k = 0; k < ModelInfoList.Count; k++)
            //        {
            //            if (ModelInfoList[k].ModelName == ShapeUnitData.ModelName)
            //            {
            //                WorksheetGraqhData.Cell(row, 3).Value = int.Parse(ModelInfoList[k].FirstHitProb);
            //                WorksheetGraqhData.Cell(row, 7).Value = int.Parse(ModelInfoList[k].ProbVarHitRashRate);
            //            }
            //        }
            //        for (int j = 0; j < ShapeUnitData.UnitList[i].HistoryData.Count; j++)
            //        {
            //            switch (ShapeUnitData.UnitList[i].HistoryData[j].HitStatus)
            //            {
            //                case 1:
            //                    firstHitCount++;

            //                    if (j < ShapeUnitData.UnitList[i].HistoryData.Count - 1
            //                        && ShapeUnitData.UnitList[i].HistoryData[j + 1].HitStatus == 2)
            //                    {
            //                        probVarFirstHitCount++;
            //                    }
            //                    allRotateCount += ShapeUnitData.UnitList[i].HistoryData[j].RotateCount;

            //                    break;
            //                case 2:
            //                    probVarHitCount++;
            //                    break;
            //            }
            //        }
            //        //確変回数
            //        WorksheetGraqhData.Cell(row, 9).Value = probVarHitCount;

            //        //0除算防止
            //        if (firstHitCount != 0)
            //        {
            //            //初当り合成確率
            //            WorksheetGraqhData.Cell(row, 4).Value = allRotateCount / firstHitCount;

            //            //確変突入率
            //            WorksheetGraqhData.Cell(row, 8).Value = probVarFirstHitCount * 100 / firstHitCount;
            //        }
            //        else
            //        {
            //            WorksheetGraqhData.Cell(row, 4).Value = allRotateCount;
            //            WorksheetGraqhData.Cell(row, 8).Value = firstHitCount;
            //        }
            //        row++;
            //    }
            //    //初当たり回数が10回以上
            //    else
            //    {
            //        List<ClassTempUnitData> TempUnitDataList = new List<ClassTempUnitData>();

            //        int remainDivFirstHitCount = FirstHitIndexList.Count % 10;
            //        int divFirstHitCount = 10;

            //        //初当たり10回未満は先頭で計算
            //        if (remainDivFirstHitCount != 0)
            //        {
            //            divFirstHitCount = remainDivFirstHitCount;
            //        }
            //        for (int j = 0; j < ShapeUnitData.UnitList[i].HistoryData.Count; j++)
            //        {
            //            switch (ShapeUnitData.UnitList[i].HistoryData[j].HitStatus)
            //            {
            //                case 1:
            //                    ResetHitCount++;
            //                    firstHitCount++;

            //                    if (j < ShapeUnitData.UnitList[i].HistoryData.Count - 1
            //                        && ShapeUnitData.UnitList[i].HistoryData[j + 1].HitStatus == 2)
            //                    {
            //                        probVarFirstHitCount++;
            //                    }
            //                    allRotateCount += ShapeUnitData.UnitList[i].HistoryData[j].RotateCount;

            //                    break;
            //                case 2:
            //                    probVarHitCount++;
            //                    break;
            //            }
            //            //初当たり10回ごとの最終データ取得時
            //            if (ResetHitCount > divFirstHitCount || j == ShapeUnitData.UnitList[i].HistoryData.Count - 1)
            //            {
            //                ClassTempUnitData TempUnitData = new ClassTempUnitData();

            //                TempUnitData.FirstHitIndex = divFirstHitCount;
            //                TempUnitData.FirstHitCount = firstHitCount - 1;
            //                TempUnitData.ProbVarHitCount = probVarHitCount;

            //                if (j < ShapeUnitData.UnitList[i].HistoryData.Count - 1
            //                    && ShapeUnitData.UnitList[i].HistoryData[j + 1].HitStatus == 2)
            //                {
            //                    TempUnitData.ProbVarFirstHitCount = probVarFirstHitCount - 1;
            //                }
            //                else
            //                {
            //                    TempUnitData.ProbVarFirstHitCount = probVarFirstHitCount;
            //                }
            //                TempUnitData.AllRotateCount = allRotateCount;
            //                TempUnitDataList.Add(TempUnitData);

            //                //リセット
            //                firstHitCount = 1;
            //                probVarHitCount = 0;

            //                if (j < ShapeUnitData.UnitList[i].HistoryData.Count - 1
            //                    && ShapeUnitData.UnitList[i].HistoryData[j + 1].HitStatus == 2)
            //                {
            //                    probVarFirstHitCount = 1;
            //                }
            //                else
            //                {
            //                    probVarFirstHitCount = 0;
            //                }
            //                allRotateCount = ShapeUnitData.UnitList[i].HistoryData[j].RotateCount;
            //                divFirstHitCount += 10;
            //            }
            //        }
            //        for (int j = 0; j < TempUnitDataList.Count; j++)
            //        {
            //            WorksheetGraqhData.Cell(row, 2).Value = TempUnitDataList[j].FirstHitIndex;
            //            WorksheetGraqhData.Cell(row, 6).Value = TempUnitDataList[j].FirstHitIndex;

            //            for (int k = 0; k < ModelInfoList.Count; k++)
            //            {
            //                if (ModelInfoList[k].ModelName == ShapeUnitData.ModelName)
            //                {
            //                    WorksheetGraqhData.Cell(row, 3).Value = int.Parse(ModelInfoList[k].FirstHitProb);
            //                    WorksheetGraqhData.Cell(row, 7).Value = int.Parse(ModelInfoList[k].ProbVarHitRashRate);
            //                }
            //            }
            //            //確変回数
            //            WorksheetGraqhData.Cell(row, 9).Value = TempUnitDataList[j].ProbVarHitCount;

            //            //0除算防止
            //            if (TempUnitDataList[j].FirstHitCount != 0)
            //            {
            //                //初当り合成確率
            //                WorksheetGraqhData.Cell(row, 4).Value = TempUnitDataList[j].AllRotateCount / TempUnitDataList[j].FirstHitCount;

            //                //確変突入率
            //                WorksheetGraqhData.Cell(row, 8).Value = TempUnitDataList[j].ProbVarFirstHitCount * 100 / TempUnitDataList[j].FirstHitCount;
            //            }
            //            else
            //            {
            //                WorksheetGraqhData.Cell(row, 4).Value = TempUnitDataList[j].AllRotateCount;
            //                WorksheetGraqhData.Cell(row, 8).Value = TempUnitDataList[j].FirstHitCount;
            //            }
            //            row++;
            //        }
            //    }
            //    GraphAreaData_FirstHitProb.LastRange.row = row - 1;
            //    GraphAreaData_FirstHitProb.LastRange.column = 4;
            //    GraphAreaData_ProbVarHitRashRate.LastRange.row = row - 1;
            //    GraphAreaData_ProbVarHitRashRate.LastRange.column = 9;

            //    //罫線を描画
            //    WorksheetGraqhData.Range(string.Concat(ConvNumToAlphabet(GraphAreaData_FirstHitProb.FirstRange.column), GraphAreaData_FirstHitProb.FirstRange.row)).Style.Border.DiagonalBorder = Excel.XLBorderStyleValues.Thin;
            //    WorksheetGraqhData.Range(string.Concat(ConvNumToAlphabet(GraphAreaData_FirstHitProb.FirstRange.column), GraphAreaData_FirstHitProb.FirstRange.row)).Style.Border.DiagonalDown = true;
            //    WorksheetGraqhData.Range(string.Concat(ConvNumToAlphabet(GraphAreaData_ProbVarHitRashRate.FirstRange.column), GraphAreaData_ProbVarHitRashRate.FirstRange.row)).Style.Border.DiagonalBorder = Excel.XLBorderStyleValues.Thin;
            //    WorksheetGraqhData.Range(string.Concat(ConvNumToAlphabet(GraphAreaData_ProbVarHitRashRate.FirstRange.column), GraphAreaData_ProbVarHitRashRate.FirstRange.row)).Style.Border.DiagonalDown = true;
            //    WorksheetGraqhData.Range(string.Concat(ConvNumToAlphabet(GraphAreaData_FirstHitProb.FirstRange.column), GraphAreaData_FirstHitProb.FirstRange.row, ":", ConvNumToAlphabet(GraphAreaData_FirstHitProb.LastRange.column), GraphAreaData_FirstHitProb.LastRange.row)).Style.Border.InsideBorder = Excel.XLBorderStyleValues.Thin;
            //    WorksheetGraqhData.Range(string.Concat(ConvNumToAlphabet(GraphAreaData_FirstHitProb.FirstRange.column), GraphAreaData_FirstHitProb.FirstRange.row, ":", ConvNumToAlphabet(GraphAreaData_FirstHitProb.LastRange.column), GraphAreaData_FirstHitProb.LastRange.row)).Style.Border.OutsideBorder = Excel.XLBorderStyleValues.Thin;
            //    WorksheetGraqhData.Range(string.Concat(ConvNumToAlphabet(GraphAreaData_ProbVarHitRashRate.FirstRange.column), GraphAreaData_ProbVarHitRashRate.FirstRange.row, ":", ConvNumToAlphabet(GraphAreaData_ProbVarHitRashRate.LastRange.column), GraphAreaData_ProbVarHitRashRate.LastRange.row)).Style.Border.InsideBorder = Excel.XLBorderStyleValues.Thin;
            //    WorksheetGraqhData.Range(string.Concat(ConvNumToAlphabet(GraphAreaData_ProbVarHitRashRate.FirstRange.column), GraphAreaData_ProbVarHitRashRate.FirstRange.row, ":", ConvNumToAlphabet(GraphAreaData_ProbVarHitRashRate.LastRange.column), GraphAreaData_ProbVarHitRashRate.LastRange.row)).Style.Border.OutsideBorder = Excel.XLBorderStyleValues.Thin;

            //    GraphAreaData.GraphAreaDataList.Add(GraphAreaData_FirstHitProb);
            //    GraphAreaData.GraphAreaDataList.Add(GraphAreaData_ProbVarHitRashRate);
            //    GraphDataList.Add(GraphAreaData);
            //    row++;
            #endregion
            return Workbook;
        }
        /// <summary>
        /// Excelへ解析データ出力(Graphシート)
        /// </summary>
        private static EPPExcel.ExcelWorkbook ExportUnitDataToExcel_Graqh(ClassShapeUnitData ShapeUnitData, EPPExcel.ExcelWorkbook Workbook)
        {
            //Graphシート
            EPPExcel.ExcelWorksheet WorksheetGraph = Workbook.Worksheets[1];

            //GraphDataシート
            EPPExcel.ExcelWorksheet WorksheetGraphData = Workbook.Worksheets[2];

            int graphNum = 1;
            int row = 2;

            for (int i = 0; i < GraphDataList.Count; i++)
            {
                //台番号
                WorksheetGraph.Cells[row, 1].Value = GraphDataList[i].UnitNum;

                for (int j = 0; j < GraphDataList[i].GraphAreaDataList.Count; j++)
                {
                    switch (GraphDataList[i].GraphAreaDataList[j].GraphKind)
                    {
                        case 1:
                            //グラフの追加
                            //EPPExcelChart.ExcelChart Chart1 = WorksheetGraph.Drawings.AddChart(string.Concat("Graph", graphNum), EPPExcelChart.eChartType.XYScatterSmoothNoMarkers);
                            EPPExcelChart.ExcelChart Chart1 = WorksheetGraph.Drawings.AddChart(string.Concat("Graph", graphNum), EPPExcelChart.eChartType.XYScatterSmoothNoMarkers);

                            //グラフの位置とサイズ
                            Chart1.SetPosition(row - 1, 0, 1, 0);
                            Chart1.SetSize(384, 240);

                            //グラフデータの設定
                            using (EPPExcel.ExcelRange RangeBaseX = WorksheetGraphData.Cells[
                                GraphDataList[i].GraphAreaDataList[j].FirstRange.row + 1, GraphDataList[i].GraphAreaDataList[j].FirstRange.column,
                                GraphDataList[i].GraphAreaDataList[j].LastRange.row, GraphDataList[i].GraphAreaDataList[j].FirstRange.column])
                            {
                                using (EPPExcel.ExcelRange RangeBaseY = WorksheetGraphData.Cells[
                                    GraphDataList[i].GraphAreaDataList[j].FirstRange.row + 1, GraphDataList[i].GraphAreaDataList[j].FirstRange.column + 1,
                                    GraphDataList[i].GraphAreaDataList[j].LastRange.row, GraphDataList[i].GraphAreaDataList[j].FirstRange.column + 1])
                                {
                                    EPPExcelChart.ExcelChartSerie ChartSerie = Chart1.Series.Add(RangeBaseY, RangeBaseX);
                                    ChartSerie.Border.Fill.Color = System.Drawing.Color.Blue;
                                    ChartSerie.Border.Width = 1.5;
                                }
                                using (EPPExcel.ExcelRange RangeBaseY = WorksheetGraphData.Cells[
                                    GraphDataList[i].GraphAreaDataList[j].FirstRange.row + 1, GraphDataList[i].GraphAreaDataList[j].FirstRange.column + 2,
                                    GraphDataList[i].GraphAreaDataList[j].LastRange.row, GraphDataList[i].GraphAreaDataList[j].FirstRange.column + 2])
                                {
                                    EPPExcelChart.ExcelChart ChartType2 = Chart1.PlotArea.ChartTypes.Add(EPPExcelChart.eChartType.ColumnClustered);
                                    EPPExcelChart.ExcelChartSerie ChartSerie = ChartType2.Series.Add(RangeBaseY, RangeBaseX);
                                    ChartSerie.Border.Fill.Color = System.Drawing.Color.Red;
                                    ChartSerie.Border.Width = 1.5;
                                    //EPPExcelChart.ExcelChartSerie ChartSerie = Chart1.Series.Add(RangeBaseY, RangeBaseX);
                                    //ChartSerie.Border.Fill.Color = System.Drawing.Color.Red;
                                    //ChartSerie.Border.Width = 1.5;
                                }
                            }
                            //X軸のラベルとメモリ設定
                            EPPExcelChart.ExcelChartAxis Axis1X = Chart1.Axis[0];
                            Axis1X.MajorUnit = 10D;
                            Axis1X.MinorUnit = 10D;
                            Axis1X.MajorTickMark = eAxisTickMark.In;
                            Axis1X.MinorTickMark = eAxisTickMark.None;

                            //Y軸のラベルとメモリ設定
                            EPPExcelChart.ExcelChartAxis Axis1Y = Chart1.Axis[1];
                            Axis1Y.MaxValue = 2000D;
                            Axis1Y.MinValue = 0D;
                            Axis1Y.MajorUnit = 100D;
                            Axis1Y.MinorUnit = 50D;
                            Axis1Y.MajorTickMark = eAxisTickMark.In;
                            Axis1Y.MinorTickMark = eAxisTickMark.In;

                            //凡例の表示
                            //Chart1.Legend.Position = eLegendPosition.Bottom;
                            //Chart1.Series[0].Header = "基準値";
                            //Chart1.Series[1].Header = "合成確率";

                            break;
                        case 2:
                            //グラフの追加
                            EPPExcelChart.ExcelChart Chart2 = WorksheetGraph.Drawings.AddChart(string.Concat("Graph", graphNum), EPPExcelChart.eChartType.XYScatterSmoothNoMarkers);

                            //グラフの位置とサイズ
                            Chart2.SetPosition(row - 1, 0, 8, 0);
                            Chart2.SetSize(384, 240);

                            //グラフデータの設定
                            using (EPPExcel.ExcelRange RangeBaseX = WorksheetGraphData.Cells[
                                GraphDataList[i].GraphAreaDataList[j].FirstRange.row + 1, GraphDataList[i].GraphAreaDataList[j].FirstRange.column,
                                GraphDataList[i].GraphAreaDataList[j].LastRange.row, GraphDataList[i].GraphAreaDataList[j].FirstRange.column])
                            {
                                using (EPPExcel.ExcelRange RangeBaseY = WorksheetGraphData.Cells[
                                    GraphDataList[i].GraphAreaDataList[j].FirstRange.row + 1, GraphDataList[i].GraphAreaDataList[j].FirstRange.column + 1,
                                    GraphDataList[i].GraphAreaDataList[j].LastRange.row, GraphDataList[i].GraphAreaDataList[j].FirstRange.column + 1])
                                {
                                    EPPExcelChart.ExcelChartSerie ChartSerie = Chart2.Series.Add(RangeBaseY, RangeBaseX);
                                    ChartSerie.Border.Fill.Color = System.Drawing.Color.Blue;
                                    ChartSerie.Border.Width = 1.5;
                                }
                                using (EPPExcel.ExcelRange RangeBaseY = WorksheetGraphData.Cells[
                                    GraphDataList[i].GraphAreaDataList[j].FirstRange.row + 1, GraphDataList[i].GraphAreaDataList[j].FirstRange.column + 2,
                                    GraphDataList[i].GraphAreaDataList[j].LastRange.row, GraphDataList[i].GraphAreaDataList[j].FirstRange.column + 2])
                                {
                                    EPPExcelChart.ExcelChartSerie ChartSerie = Chart2.Series.Add(RangeBaseY, RangeBaseX);
                                    ChartSerie.Border.Fill.Color = System.Drawing.Color.DarkOrange;
                                    ChartSerie.Border.Width = 1.5;
                                }
                            }
                            //X軸のラベルとメモリ設定
                            EPPExcelChart.ExcelChartAxis Axis2X = Chart2.Axis[0];
                            Axis2X.MajorUnit = 10D;
                            Axis2X.MinorUnit = 10D;
                            Axis2X.MajorTickMark = eAxisTickMark.In;
                            Axis2X.MinorTickMark = eAxisTickMark.None;

                            //Y軸のラベルとメモリ設定
                            EPPExcelChart.ExcelChartAxis Axis2Y = Chart2.Axis[1];
                            Axis2Y.MajorUnit = 10D;
                            Axis2Y.MinorUnit = 5D;
                            Axis2Y.MajorTickMark = eAxisTickMark.In;
                            Axis2Y.MinorTickMark = eAxisTickMark.In;

                            //凡例の表示
                            Chart2.Legend.Position = eLegendPosition.Bottom;
                            Chart2.Series[0].Header = "基準値";
                            Chart2.Series[1].Header = "確変突入率";
                            graphNum++;

                            //グラフの追加
                            EPPExcelChart.ExcelChart Chart3 = WorksheetGraph.Drawings.AddChart(string.Concat("Graph", graphNum), EPPExcelChart.eChartType.XYScatterSmoothNoMarkers);

                            //グラフの位置とサイズ
                            Chart3.SetPosition(row - 1, 0, 15, 0);
                            Chart3.SetSize(384, 240);

                            //グラフデータの設定
                            using (EPPExcel.ExcelRange RangeBaseX = WorksheetGraphData.Cells[
                                GraphDataList[i].GraphAreaDataList[j].FirstRange.row + 1, GraphDataList[i].GraphAreaDataList[j].FirstRange.column,
                                GraphDataList[i].GraphAreaDataList[j].LastRange.row, GraphDataList[i].GraphAreaDataList[j].FirstRange.column])
                            {
                                using (EPPExcel.ExcelRange RangeBaseY = WorksheetGraphData.Cells[
                                    GraphDataList[i].GraphAreaDataList[j].FirstRange.row + 1, GraphDataList[i].GraphAreaDataList[j].FirstRange.column + 3,
                                    GraphDataList[i].GraphAreaDataList[j].LastRange.row, GraphDataList[i].GraphAreaDataList[j].FirstRange.column + 3])
                                {
                                    EPPExcelChart.ExcelChartSerie ChartSerie = Chart3.Series.Add(RangeBaseY, RangeBaseX);
                                    ChartSerie.Border.Fill.Color = System.Drawing.Color.Green;
                                    ChartSerie.Border.Width = 1.5;
                                }
                            }
                            //X軸のラベルとメモリ設定
                            EPPExcelChart.ExcelChartAxis Axis3X = Chart3.Axis[0];
                            Axis3X.MajorUnit = 10D;
                            Axis3X.MinorUnit = 10D;
                            Axis3X.MajorTickMark = eAxisTickMark.In;
                            Axis3X.MinorTickMark = eAxisTickMark.None;

                            //Y軸のラベルとメモリ設定
                            EPPExcelChart.ExcelChartAxis Axis3Y = Chart3.Axis[1];
                            Axis3Y.MajorUnit = 10D;
                            Axis3Y.MinorUnit = 5D;
                            Axis3Y.MajorTickMark = eAxisTickMark.In;
                            Axis3Y.MinorTickMark = eAxisTickMark.In;

                            //凡例の表示
                            Chart3.Legend.Position = eLegendPosition.Bottom;
                            Chart3.Series[0].Header = "確変回数";

                            break;
                    }
                    graphNum++;
                }
                row += 13;
            }
            return Workbook;
        }
    }
}