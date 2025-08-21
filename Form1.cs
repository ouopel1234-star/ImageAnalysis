using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using AForge.Video;
using AForge.Video.DirectShow;
using System.IO;

namespace ImageAnalysis
{
    // MTF 方法列舉
    public enum MTFMethod
    {
        PixelDifference, // 像素灰度差異
        MTF50           // 標準 MTF50
    }

    // MTF 設定資料結構
    public class MTFSettingData
    {
        public string ResultFilesDirectory { get; set; } = "";
        public string SettingFile { get; set; } = "";
        public MTFMethod MTFMethod { get; set; } = MTFMethod.PixelDifference;
        public List<MTFRegionData> MTFRegions { get; set; } = new List<MTFRegionData>();
        public int Criterion { get; set; } = 50;
    }

    // 修正 MTFRegionData 類別，加入 Criterion 屬性
    public class MTFRegionData
    {
        public string Name { get; set; } = "";
        public int X { get; set; } = 0;
        public int Y { get; set; } = 0;
        public string CurrentValue { get; set; } = "0.00";
        public string MaxValue { get; set; } = "0.00";
        public bool IsSelected { get; set; } = true;
        public int Criterion { get; set; } = 50; // 新增 Criterion 屬性
    }


    public partial class Form1 : Form
    {
        private FilterInfoCollection videoDevices;
        private VideoCaptureDevice videoDevice;
        private bool isDeviceRunning = false;

        // MTF擷取區域
        private Rectangle[] mtfRegions = new Rectangle[5]; // M0, M1, M2, M3, MC
        private Rectangle[] colorRegions = new Rectangle[3]; // C1, C2, C3
        private bool[] mtfRegionVisible = new bool[5];
        private bool[] colorRegionVisible = new bool[3];

        // 拖曳相關變數
        private bool isDragging = false;
        private int dragRegionIndex = -1;
        private bool isDragMTF = true; // true for MTF, false for Color
        private Point dragStartPoint;

        // MTF 結果顯示
        private GroupBox[] mtfGroups = new GroupBox[5];
        private PictureBox[] mtfPictureBoxes = new PictureBox[5];
        private Label[] mtfCurrentLabels = new Label[5];
        private Label[] mtfMaxLabels = new Label[5];

        // 顏色結果顯示
        private Label[] colorLabels = new Label[3];

        // 當前圖像
        private Bitmap currentImage;
        private float imageScale = 1.0f;

        // MTF(SFR)折線圖表成員變數
        private PictureBox[] mtfChartBoxes = new PictureBox[5];
        private List<double>[] mtfHistoryData = new List<double>[5];
        private List<DateTime>[] mtfTimeStamps = new List<DateTime>[5];
        private Timer chartUpdateTimer;
        private const int MAX_DATA_POINTS = 60; // 顯示最近60秒的資料

        // 影像拖曳成員變數
        private float currentPB1Ratio = 1.0f; // 當前 PB1 比例
        private PointF imageOffset = new PointF(0, 0); // 影像偏移量
        private bool isImageDragging = false; // 影像拖曳狀態
        private Point imageDragStartPoint; // 拖曳開始點

        // 執行緒同步物件
        private readonly object imageLock = new object();

        // 設定檔相關成員變數
        private string currentSettingsFilePath = "";

        // 新增：儲存 Criterion 值的成員變數
        private int[] mtfCriterionValues = new int[5] { 50, 50, 50, 50, 50 };



        // 設定檔參數結構
        // 修正 ImageAnalysisSettings 類別，加入 Criterion 屬性
        public class ImageAnalysisSettings
        {
            public int MTFArea { get; set; } = 128;
            public int OutputRatio { get; set; } = 100;

            // MTF 區域中心點
            public Point M0_Center { get; set; } = new Point(164, 164);
            public Point M1_Center { get; set; } = new Point(164, 464);
            public Point M2_Center { get; set; } = new Point(664, 464);
            public Point M3_Center { get; set; } = new Point(664, 164);
            public Point MC_Center { get; set; } = new Point(414, 314);

            // 彩色區域中心點
            public Point C1_Center { get; set; } = new Point(364, 414);
            public Point C2_Center { get; set; } = new Point(264, 514);
            public Point C3_Center { get; set; } = new Point(564, 514);

            // MTF Criterion 值
            public int M0_Criterion { get; set; } = 50;
            public int M1_Criterion { get; set; } = 50;
            public int M2_Criterion { get; set; } = 50;
            public int M3_Criterion { get; set; } = 50;
            public int MC_Criterion { get; set; } = 50;
        }

        // MTF 方法選擇變數
        private MTFMethod currentMTFMethod = MTFMethod.PixelDifference;

        // MTF Setting Form 參考
        private MTFSettingForm mtfSettingForm = null;
        private bool isFormSyncing = false; // 防止遞迴同步

        public event Action<MTFSettingData> MTFDataChanged;
        public event Action<int, Point> MTFRegionMoved; // regionIndex, newCenter









        // 修正 Form1 建構子
        public Form1()
        {
            // 先建立介面
            Initialize_Component();

            // 再初始化其他元件
            InitializeRegions();
            InitializeMTFGroups();
            InitializeColorLabels();

            // 最後初始化攝影機
            InitializeCamera();

            // 設定檔初始化移到 Form_Load 事件
        }

        // 修正 InitializeComponent 方法，加入 Load 事件註冊
        private void Initialize_Component()
        {
            this.SuspendLayout();

            // Form 設定
            this.Text = "C#ImageAnalysis V0.0";
            this.Size = new Size(1400, 900);
            this.WindowState = FormWindowState.Maximized;
            this.StartPosition = FormStartPosition.CenterScreen;

            // 註冊事件
            this.FormClosing += Form1_FormClosing;
            this.Load += Form1_Load; // 加入 Load 事件

            // 建立控制項
            CreateControls();

            this.ResumeLayout(false);
        }

        // 加入 Form1_Load 事件處理方法
        private void Form1_Load(object sender, EventArgs e)
        {
            // 在 Form 完全載入後初始化設定檔
            try
            {
                InitializeSettingsFile();
            }
            catch (Exception ex)
            {
                MessageBox.Show("載入設定檔時發生錯誤: " + ex.Message);
            }
        }

        // 設定檔初始化方法
        // 簡化 InitializeSettingsFile 方法，移除 BeginInvoke
        private void InitializeSettingsFile()
        {
            try
            {
                string exePath = Application.StartupPath;
                string[] setFiles = Directory.GetFiles(exePath, "*.set");

                if (setFiles.Length == 0)
                {
                    // 沒有設定檔，建立新的
                    CreateNewSettingsFile();
                }
                else
                {
                    // 有設定檔，讓使用者選擇
                    SelectAndLoadSettingsFile();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("設定檔初始化失敗: " + ex.Message);
                CreateNewSettingsFile(); // 失敗時建立新檔
            }
        }

        // 建立新設定檔
        // 修正 CreateNewSettingsFile 方法，確保目錄存在
        private void CreateNewSettingsFile()
        {
            try
            {
                string exePath = Application.StartupPath;

                // 確保目錄存在
                if (!Directory.Exists(exePath))
                {
                    Directory.CreateDirectory(exePath);
                }

                string fileName = $"IASetsup_{DateTime.Now:yyyyMMdd}.set";
                currentSettingsFilePath = Path.Combine(exePath, fileName);

                // 建立預設設定
                ImageAnalysisSettings defaultSettings = new ImageAnalysisSettings();

                // 儲存預設設定
                SaveSettingsToFile(defaultSettings, currentSettingsFilePath);

                // 套用預設設定到介面
                ApplySettingsToUI(defaultSettings);

                MessageBox.Show($"已建立新的參數設定檔: {fileName}");
            }
            catch (Exception ex)
            {
                MessageBox.Show("建立設定檔失敗: " + ex.Message);
            }
        }



        // 選擇並載入設定檔
        private void SelectAndLoadSettingsFile()
        {
            try
            {
                OpenFileDialog ofd = new OpenFileDialog();
                ofd.InitialDirectory = Application.StartupPath;
                ofd.Filter = "Settings Files (*.set)|*.set";
                ofd.Title = "選擇參數設定檔";

                // 如果只有一個設定檔，直接載入
                string[] setFiles = Directory.GetFiles(Application.StartupPath, "*.set");
                if (setFiles.Length == 1)
                {
                    currentSettingsFilePath = setFiles[0];
                    ImageAnalysisSettings settings = LoadSettingsFromFile(currentSettingsFilePath);
                    if (settings != null)
                    {
                        ApplySettingsToUI(settings);
                        MessageBox.Show($"已載入參數設定檔: {Path.GetFileName(currentSettingsFilePath)}");
                    }
                    return;
                }

                // 多個設定檔時讓使用者選擇
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    currentSettingsFilePath = ofd.FileName;

                    // 載入設定
                    ImageAnalysisSettings settings = LoadSettingsFromFile(currentSettingsFilePath);
                    if (settings != null)
                    {
                        ApplySettingsToUI(settings);
                        MessageBox.Show($"已載入參數設定檔: {Path.GetFileName(currentSettingsFilePath)}");
                    }
                }
                else
                {
                    // 使用者取消選擇，使用第一個找到的設定檔
                    if (setFiles.Length > 0)
                    {
                        currentSettingsFilePath = setFiles[0];
                        ImageAnalysisSettings settings = LoadSettingsFromFile(currentSettingsFilePath);
                        if (settings != null)
                        {
                            ApplySettingsToUI(settings);
                            MessageBox.Show($"已載入預設參數設定檔: {Path.GetFileName(currentSettingsFilePath)}");
                        }
                    }
                    else
                    {
                        // 完全沒有設定檔，建立新的
                        CreateNewSettingsFile();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("載入設定檔失敗: " + ex.Message);
                CreateNewSettingsFile();
            }
        }


        // 從檔案載入設定
        // 修正 LoadSettingsFromFile 方法，載入 Criterion 值
        private ImageAnalysisSettings LoadSettingsFromFile(string filePath)
        {
            try
            {
                if (!System.IO.File.Exists(filePath))
                {
                    MessageBox.Show($"設定檔不存在: {filePath}");
                    return new ImageAnalysisSettings(); // 回傳預設設定
                }

                ImageAnalysisSettings settings = new ImageAnalysisSettings();
                string[] lines = System.IO.File.ReadAllLines(filePath);

                foreach (string line in lines)
                {
                    if (string.IsNullOrWhiteSpace(line) || line.StartsWith("//"))
                        continue;

                    string[] parts = line.Split('=');
                    if (parts.Length != 2)
                        continue;

                    string key = parts[0].Trim();
                    string value = parts[1].Trim();

                    switch (key)
                    {
                        case "MTFArea":
                            if (int.TryParse(value, out int mtfArea))
                                settings.MTFArea = mtfArea;
                            break;

                        case "OutputRatio":
                            if (int.TryParse(value, out int outputRatio))
                                settings.OutputRatio = outputRatio;
                            break;

                        // MTF 區域中心點
                        case "M0_Center":
                            settings.M0_Center = ParsePointFromString(value);
                            break;
                        case "M1_Center":
                            settings.M1_Center = ParsePointFromString(value);
                            break;
                        case "M2_Center":
                            settings.M2_Center = ParsePointFromString(value);
                            break;
                        case "M3_Center":
                            settings.M3_Center = ParsePointFromString(value);
                            break;
                        case "MC_Center":
                            settings.MC_Center = ParsePointFromString(value);
                            break;

                        // 彩色區域中心點
                        case "C1_Center":
                            settings.C1_Center = ParsePointFromString(value);
                            break;
                        case "C2_Center":
                            settings.C2_Center = ParsePointFromString(value);
                            break;
                        case "C3_Center":
                            settings.C3_Center = ParsePointFromString(value);
                            break;

                        // MTF Criterion 值
                        case "M0_Criterion":
                            if (int.TryParse(value, out int m0Crit))
                                settings.M0_Criterion = m0Crit;
                            break;
                        case "M1_Criterion":
                            if (int.TryParse(value, out int m1Crit))
                                settings.M1_Criterion = m1Crit;
                            break;
                        case "M2_Criterion":
                            if (int.TryParse(value, out int m2Crit))
                                settings.M2_Criterion = m2Crit;
                            break;
                        case "M3_Criterion":
                            if (int.TryParse(value, out int m3Crit))
                                settings.M3_Criterion = m3Crit;
                            break;
                        case "MC_Criterion":
                            if (int.TryParse(value, out int mcCrit))
                                settings.MC_Criterion = mcCrit;
                            break;
                    }
                }

                return settings;
            }
            catch (Exception ex)
            {
                MessageBox.Show("讀取設定檔內容失敗: " + ex.Message);
                return new ImageAnalysisSettings(); // 回傳預設設定
            }
        }


        // 解析字串為 Point
        private Point ParsePointFromString(string pointStr)
        {
            try
            {
                string[] coords = pointStr.Split(',');
                if (coords.Length == 2)
                {
                    int x = int.Parse(coords[0].Trim());
                    int y = int.Parse(coords[1].Trim());
                    return new Point(x, y);
                }
            }
            catch { }

            return new Point(0, 0);
        }

        // 儲存設定到檔案
        // 修正 SaveSettingsToFile 方法，加入彩色區域儲存
        // 修正 SaveSettingsToFile 方法，儲存 Criterion 值
        private void SaveSettingsToFile(ImageAnalysisSettings settings, string filePath)
        {
            try
            {
                List<string> lines = new List<string>();

                // 加入檔案標頭
                lines.Add("// Image Analysis Settings File");
                lines.Add("// Created: " + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));
                lines.Add("// =====================================");
                lines.Add("");

                // 加入基本參數
                lines.Add("// Basic Parameters");
                lines.Add("MTFArea=" + settings.MTFArea);
                lines.Add("OutputRatio=" + settings.OutputRatio);
                lines.Add("");

                // 加入 MTF 區域參數
                lines.Add("// MTF Region Centers (X,Y)");
                lines.Add("M0_Center=" + settings.M0_Center.X + "," + settings.M0_Center.Y);
                lines.Add("M1_Center=" + settings.M1_Center.X + "," + settings.M1_Center.Y);
                lines.Add("M2_Center=" + settings.M2_Center.X + "," + settings.M2_Center.Y);
                lines.Add("M3_Center=" + settings.M3_Center.X + "," + settings.M3_Center.Y);
                lines.Add("MC_Center=" + settings.MC_Center.X + "," + settings.MC_Center.Y);
                lines.Add("");

                // 加入彩色區域參數
                lines.Add("// Color Region Centers (X,Y)");
                lines.Add("C1_Center=" + settings.C1_Center.X + "," + settings.C1_Center.Y);
                lines.Add("C2_Center=" + settings.C2_Center.X + "," + settings.C2_Center.Y);
                lines.Add("C3_Center=" + settings.C3_Center.X + "," + settings.C3_Center.Y);
                lines.Add("");

                // 加入 MTF Criterion 值
                lines.Add("// MTF Criterion Values");
                lines.Add("M0_Criterion=" + settings.M0_Criterion);
                lines.Add("M1_Criterion=" + settings.M1_Criterion);
                lines.Add("M2_Criterion=" + settings.M2_Criterion);
                lines.Add("M3_Criterion=" + settings.M3_Criterion);
                lines.Add("MC_Criterion=" + settings.MC_Criterion);
                lines.Add("");

                lines.Add("// End of Settings");

                System.IO.File.WriteAllLines(filePath, lines, Encoding.UTF8);
            }
            catch (Exception ex)
            {
                MessageBox.Show("儲存設定檔失敗: " + ex.Message);
            }
        }

        // 套用設定到介面
        // 修正 ApplySettingsToUI 方法，加入控制項存在檢查
        // 修正 ApplySettingsToUI 方法，加入彩色區域位置套用
        // 修正 ApplySettingsToUI 方法，載入 Criterion 值
        private void ApplySettingsToUI(ImageAnalysisSettings settings)
        {
            try
            {
                // 套用 MTF Area
                ComboBox cmbMTFArea = this.Controls.Find("cmbMTFArea", true).FirstOrDefault() as ComboBox;
                if (cmbMTFArea != null)
                {
                    cmbMTFArea.Text = settings.MTFArea.ToString();
                }

                // 套用 Output Ratio
                ComboBox cmbOutputRatio = this.Controls.Find("cmbOutputRatio", true).FirstOrDefault() as ComboBox;
                if (cmbOutputRatio != null)
                {
                    cmbOutputRatio.Text = settings.OutputRatio.ToString();
                }

                // 確保區域陣列已初始化
                if (mtfRegions != null && mtfRegions.Length >= 5 &&
                    colorRegions != null && colorRegions.Length >= 3)
                {
                    // 套用 MTF 區域位置 (從中心點計算左上角)
                    int halfSize = settings.MTFArea / 2;

                    mtfRegions[0] = new Rectangle(settings.M0_Center.X - halfSize, settings.M0_Center.Y - halfSize, settings.MTFArea, settings.MTFArea);
                    mtfRegions[1] = new Rectangle(settings.M1_Center.X - halfSize, settings.M1_Center.Y - halfSize, settings.MTFArea, settings.MTFArea);
                    mtfRegions[2] = new Rectangle(settings.M2_Center.X - halfSize, settings.M2_Center.Y - halfSize, settings.MTFArea, settings.MTFArea);
                    mtfRegions[3] = new Rectangle(settings.M3_Center.X - halfSize, settings.M3_Center.Y - halfSize, settings.MTFArea, settings.MTFArea);
                    mtfRegions[4] = new Rectangle(settings.MC_Center.X - halfSize, settings.MC_Center.Y - halfSize, settings.MTFArea, settings.MTFArea);

                    // 套用彩色區域位置 (從中心點計算左上角)
                    colorRegions[0] = new Rectangle(settings.C1_Center.X - halfSize, settings.C1_Center.Y - halfSize, settings.MTFArea, settings.MTFArea);
                    colorRegions[1] = new Rectangle(settings.C2_Center.X - halfSize, settings.C2_Center.Y - halfSize, settings.MTFArea, settings.MTFArea);
                    colorRegions[2] = new Rectangle(settings.C3_Center.X - halfSize, settings.C3_Center.Y - halfSize, settings.MTFArea, settings.MTFArea);
                }

                // 儲存 Criterion 值到成員變數（用於 MTF Setting Form）
                StoreCriterionValues(settings);

                // 更新顯示
                PictureBox pb = this.Controls.Find("pictureBox1", true).FirstOrDefault() as PictureBox;
                if (pb != null)
                    pb.Invalidate();
            }
            catch (Exception ex)
            {
                MessageBox.Show("套用設定失敗: " + ex.Message);
            }
        }


        // 從介面取得當前設定
        // 修正 GetCurrentSettingsFromUI 方法，加入控制項存在檢查
        // 修正 GetCurrentSettingsFromUI 方法，加入彩色區域位置取得
        private ImageAnalysisSettings GetCurrentSettingsFromUI()
        {
            try
            {
                ImageAnalysisSettings settings = new ImageAnalysisSettings();

                // 取得 MTF Area
                ComboBox cmbMTFArea = this.Controls.Find("cmbMTFArea", true).FirstOrDefault() as ComboBox;
                if (cmbMTFArea != null && int.TryParse(cmbMTFArea.Text, out int mtfArea))
                    settings.MTFArea = mtfArea;

                // 取得 Output Ratio
                ComboBox cmbOutputRatio = this.Controls.Find("cmbOutputRatio", true).FirstOrDefault() as ComboBox;
                if (cmbOutputRatio != null && int.TryParse(cmbOutputRatio.Text, out int outputRatio))
                    settings.OutputRatio = outputRatio;

                // 確保區域陣列已初始化且有效
                if (mtfRegions != null && mtfRegions.Length >= 5)
                {
                    // 取得 MTF 區域中心點
                    settings.M0_Center = new Point(mtfRegions[0].X + mtfRegions[0].Width / 2, mtfRegions[0].Y + mtfRegions[0].Height / 2);
                    settings.M1_Center = new Point(mtfRegions[1].X + mtfRegions[1].Width / 2, mtfRegions[1].Y + mtfRegions[1].Height / 2);
                    settings.M2_Center = new Point(mtfRegions[2].X + mtfRegions[2].Width / 2, mtfRegions[2].Y + mtfRegions[2].Height / 2);
                    settings.M3_Center = new Point(mtfRegions[3].X + mtfRegions[3].Width / 2, mtfRegions[3].Y + mtfRegions[3].Height / 2);
                    settings.MC_Center = new Point(mtfRegions[4].X + mtfRegions[4].Width / 2, mtfRegions[4].Y + mtfRegions[4].Height / 2);
                }

                if (colorRegions != null && colorRegions.Length >= 3)
                {
                    // 取得彩色區域中心點
                    settings.C1_Center = new Point(colorRegions[0].X + colorRegions[0].Width / 2, colorRegions[0].Y + colorRegions[0].Height / 2);
                    settings.C2_Center = new Point(colorRegions[1].X + colorRegions[1].Width / 2, colorRegions[1].Y + colorRegions[1].Height / 2);
                    settings.C3_Center = new Point(colorRegions[2].X + colorRegions[2].Width / 2, colorRegions[2].Y + colorRegions[2].Height / 2);
                }

                return settings;
            }
            catch (Exception ex)
            {
                MessageBox.Show("取得當前設定失敗: " + ex.Message);
                return new ImageAnalysisSettings();
            }
        }



        private void CreateControls()
        {
            // Device ComboBox
            ComboBox cmbDevices = new ComboBox();
            cmbDevices.Name = "cmbDevices";
            cmbDevices.Location = new Point(15, 57);
            cmbDevices.Size = new Size(150, 25);
            cmbDevices.DropDownStyle = ComboBoxStyle.DropDownList;
            this.Controls.Add(cmbDevices);

            Label lblDevices = new Label();
            lblDevices.Text = "Devices";
            lblDevices.Location = new Point(15, 37);
            lblDevices.Size = new Size(60, 20);
            this.Controls.Add(lblDevices);

            // FrameRate
            ComboBox cmbFrameRate = new ComboBox();
            cmbFrameRate.Name = "cmbFrameRate";
            cmbFrameRate.Location = new Point(180, 57);
            cmbFrameRate.Size = new Size(80, 25);
            cmbFrameRate.Items.AddRange(new string[] { "7.00", "15.00", "30.00" });
            cmbFrameRate.Text = "7.00";
            this.Controls.Add(cmbFrameRate);

            Label lblFrameRate = new Label();
            lblFrameRate.Text = "FrameRate";
            lblFrameRate.Location = new Point(180, 37);
            lblFrameRate.Size = new Size(70, 20);
            this.Controls.Add(lblFrameRate);

            // Resolution
            ComboBox cmbResolution = new ComboBox();
            cmbResolution.Name = "cmbResolution";
            cmbResolution.Location = new Point(15, 100);
            cmbResolution.Size = new Size(100, 25);
            cmbResolution.Items.AddRange(new string[] { "1920, 1080", "1280, 720", "640, 480" });
            cmbResolution.Text = "1920, 1080";
            this.Controls.Add(cmbResolution);

            Label lblResolution = new Label();
            lblResolution.Text = "Resolution";
            lblResolution.Location = new Point(15, 80);
            lblResolution.Size = new Size(70, 20);
            this.Controls.Add(lblResolution);

            // MTF Area
            ComboBox cmbMTFArea = new ComboBox();
            cmbMTFArea.Name = "cmbMTFArea";
            cmbMTFArea.Location = new Point(275, 57);
            cmbMTFArea.Size = new Size(80, 25);
            cmbMTFArea.Items.AddRange(new string[] { "64", "128", "256" });
            cmbMTFArea.Text = "128";
            cmbMTFArea.SelectedIndexChanged += CmbMTFArea_SelectedIndexChanged;
            this.Controls.Add(cmbMTFArea);

            Label lblMTFArea = new Label();
            lblMTFArea.Text = "MTF Area";
            lblMTFArea.Location = new Point(275, 37);
            lblMTFArea.Size = new Size(70, 20);
            this.Controls.Add(lblMTFArea);

            // PB1 Ratio
            ComboBox cmbPB1Ratio = new ComboBox();
            cmbPB1Ratio.Name = "cmbPB1Ratio";
            cmbPB1Ratio.Location = new Point(130, 100);
            cmbPB1Ratio.Size = new Size(80, 25);
            cmbPB1Ratio.Items.AddRange(new string[] { "50 %", "75 %", "100 %" });
            cmbPB1Ratio.Text = "100 %"; // 預設 100%
            cmbPB1Ratio.SelectedIndexChanged += CmbPB1Ratio_SelectedIndexChanged; // 加入事件處理
            this.Controls.Add(cmbPB1Ratio);

            Label lblPB1Ratio = new Label();
            lblPB1Ratio.Text = "PB1 Ratio";
            lblPB1Ratio.Location = new Point(130, 80);
            lblPB1Ratio.Size = new Size(70, 20);
            this.Controls.Add(lblPB1Ratio);

            // Output Ratio
            ComboBox cmbOutputRatio = new ComboBox();
            cmbOutputRatio.Name = "cmbOutputRatio";
            cmbOutputRatio.Location = new Point(275, 100);
            cmbOutputRatio.Size = new Size(80, 25);
            cmbOutputRatio.Items.AddRange(new string[] { "10", "30", "50", "100", "130", "150", "200" });
            cmbOutputRatio.Text = "100";
            this.Controls.Add(cmbOutputRatio);

            Label lblOutputRatio = new Label();
            lblOutputRatio.Text = "Output Ratio";
            lblOutputRatio.Location = new Point(275, 80);
            lblOutputRatio.Size = new Size(80, 20);
            this.Controls.Add(lblOutputRatio);

            // 按鍵區域 (PB)
            CreateButtons();

            // Point 區域
            CreatePointLabels();

            // Denoising_Edge 區域
            CreateDenosingEdgeControls();

            // MTF Result 和 Color Result 區域
            CreateResultControls();

            // 主要 PictureBox - 完全自訂繪製模式
            PictureBox pictureBox1 = new PictureBox();
            pictureBox1.Name = "pictureBox1";
            pictureBox1.Location = new Point(15, 140);
            pictureBox1.Size = new Size(800, 600);
            pictureBox1.BackColor = Color.Black;
            pictureBox1.SizeMode = PictureBoxSizeMode.Normal;
            // 不設定 Image 屬性，完全透過 Paint 事件繪製
            pictureBox1.MouseDown += PictureBox1_MouseDown;
            pictureBox1.MouseMove += PictureBox1_MouseMove;
            pictureBox1.MouseUp += PictureBox1_MouseUp;
            pictureBox1.Paint += PictureBox1_Paint;
            this.Controls.Add(pictureBox1);
        }

        private void CreateButtons()
        {
            string[] buttonTexts = { "Start", "Open", "Stop", "Save Stop", "Save Image", "Video", "Exit" };
            Point startPos = new Point(380, 37);

            for (int i = 0; i < buttonTexts.Length; i++)
            {
                Button btn = new Button();
                btn.Text = buttonTexts[i];
                btn.Name = "btn" + buttonTexts[i].Replace(" ", "");
                btn.Size = new Size(80, 25);
                btn.Location = new Point(startPos.X + (i % 4) * 85, startPos.Y + (i / 4) * 30);
                btn.Click += Button_Click;
                this.Controls.Add(btn);
            }
        }

        private void CreatePointLabels()
        {
            // Point 標籤
            Label lblPointTitle = new Label();
            lblPointTitle.Text = "Point";
            lblPointTitle.Location = new Point(580, 37);
            lblPointTitle.Size = new Size(40, 20);
            this.Controls.Add(lblPointTitle);

            Label lblPoint = new Label();
            lblPoint.Name = "lblPoint";
            lblPoint.Text = "Point: 477, 2, 1.50, 1.35";
            lblPoint.Location = new Point(580, 57);
            lblPoint.Size = new Size(150, 20);
            this.Controls.Add(lblPoint);

            Label lblRealPoint = new Label();
            lblRealPoint.Name = "lblRealPoint";
            lblRealPoint.Text = "Real Point: 715, 2";
            lblRealPoint.Location = new Point(580, 77);
            lblRealPoint.Size = new Size(150, 20);
            this.Controls.Add(lblRealPoint);

            Label lblScreenScale = new Label();
            lblScreenScale.Name = "lblScreenScale";
            lblScreenScale.Text = "Screen Scale: 1.50";
            lblScreenScale.Location = new Point(580, 97);
            lblScreenScale.Size = new Size(150, 20);
            this.Controls.Add(lblScreenScale);
        }

        // CreateDenosingEdgeControls 方法，加入 MTF Setting 選項
        private void CreateDenosingEdgeControls()
        {
            Label lblDenosingEdge = new Label();
            lblDenosingEdge.Text = "Denoising_Edge";
            lblDenosingEdge.Location = new Point(750, 37);
            lblDenosingEdge.Size = new Size(100, 20);
            this.Controls.Add(lblDenosingEdge);

            ComboBox cmbDenoise1 = new ComboBox();
            cmbDenoise1.Name = "cmbDenoise1";
            cmbDenoise1.Location = new Point(750, 57);
            cmbDenoise1.Size = new Size(80, 25);
            cmbDenoise1.Items.AddRange(new string[] { "None", "FastPixel" });
            cmbDenoise1.Text = "None";
            this.Controls.Add(cmbDenoise1);

            ComboBox cmbDenoise2 = new ComboBox();
            cmbDenoise2.Name = "cmbDenoise2";
            cmbDenoise2.Location = new Point(750, 87);
            cmbDenoise2.Size = new Size(80, 25);
            cmbDenoise2.Items.AddRange(new string[] { "None", "FastPixel" });
            cmbDenoise2.Text = "FastPixel";
            this.Controls.Add(cmbDenoise2);

            // 第三個下拉式選單 - 加入 MTF Setting 選項
            ComboBox cmbDenoise3 = new ComboBox();
            cmbDenoise3.Name = "cmbDenoise3";
            cmbDenoise3.Location = new Point(750, 117);
            cmbDenoise3.Size = new Size(80, 25);
            cmbDenoise3.Items.AddRange(new string[] { "None", "MTF Setting" });
            cmbDenoise3.Text = "None";
            cmbDenoise3.SelectedIndexChanged += CmbDenoise3_SelectedIndexChanged;
            this.Controls.Add(cmbDenoise3);
        }


        private void CreateResultControls()
        {
            // MTF Result
            Label lblMTFResult = new Label();
            lblMTFResult.Text = "MTF Result";
            lblMTFResult.Location = new Point(950, 37);
            lblMTFResult.Size = new Size(70, 20);
            this.Controls.Add(lblMTFResult);

            // MTF 區域按鈕
            string[] mtfButtons = { "M0", "M1", "M2", "M3", "MC" };
            for (int i = 0; i < mtfButtons.Length; i++)
            {
                CheckBox chk = new CheckBox();
                chk.Text = mtfButtons[i];
                chk.Name = "chk" + mtfButtons[i];
                chk.Location = new Point(950 + (i % 2) * 40, 57 + (i / 2) * 25);
                chk.Size = new Size(35, 20);
                chk.Checked = true;
                chk.CheckedChanged += MtfCheckBox_CheckedChanged;
                this.Controls.Add(chk);
            }

            // MTF 數值顯示
            Label lblMTFValues = new Label();
            lblMTFValues.Name = "lblMTFValues";
            lblMTFValues.Text = "MTF 0 -- 16.56\nMTF 1 -- 16.83\nMTF 2 -- 2.83\nMTF 3 -- 4.27\nMTF 4 -- 33.55";
            lblMTFValues.Location = new Point(1050, 57);
            lblMTFValues.Size = new Size(100, 100);
            lblMTFValues.ForeColor = Color.Red;
            this.Controls.Add(lblMTFValues);

            // Color Result
            Label lblColorResult = new Label();
            lblColorResult.Text = "Color Result";
            lblColorResult.Location = new Point(1170, 37);
            lblColorResult.Size = new Size(80, 20);
            this.Controls.Add(lblColorResult);

            // Color 區域按鈕
            string[] colorButtons = { "C1", "C2", "C3" };
            for (int i = 0; i < colorButtons.Length; i++)
            {
                CheckBox chk = new CheckBox();
                chk.Text = colorButtons[i];
                chk.Name = "chk" + colorButtons[i];
                chk.Location = new Point(1170, 57 + i * 25);
                chk.Size = new Size(35, 20);
                chk.Checked = true;
                chk.CheckedChanged += ColorCheckBox_CheckedChanged;
                this.Controls.Add(chk);
            }

            // Color 數值顯示
            Label lblColorValues = new Label();
            lblColorValues.Name = "lblColorValues";
            lblColorValues.Text = "C1 -- 98.7, 2.0, -1.5\nC2 -- 52.8, 5.3, -0.5\nC3 -- 67.5, 10.8, -15.4";
            lblColorValues.Location = new Point(1210, 57);
            lblColorValues.Size = new Size(150, 100);
            this.Controls.Add(lblColorValues);
        }

        // 修正 InitializeCamera 方法，加入錯誤處理
        private void InitializeCamera()
        {
            try
            {
                videoDevices = new FilterInfoCollection(FilterCategory.VideoInputDevice);
                ComboBox cmbDevices = this.Controls.Find("cmbDevices", true).FirstOrDefault() as ComboBox;

                if (cmbDevices != null)
                {
                    if (videoDevices.Count > 0)
                    {
                        foreach (FilterInfo device in videoDevices)
                        {
                            cmbDevices.Items.Add(device.Name);
                        }
                        if (cmbDevices.Items.Count > 0)
                            cmbDevices.SelectedIndex = 0;
                    }
                    else
                    {
                        cmbDevices.Items.Add("No devices found");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("初始化攝影機失敗: " + ex.Message);
                // 不要阻止程式繼續執行
            }
        }


        // 修正 InitializeRegions 方法，確保彩色區域使用正確的預設值
        // 修正 InitializeRegions 方法，確保正確初始化
        private void InitializeRegions()
        {
            int mtfSize = 128; // 預設 MTF Area 大小

            // 確保陣列已初始化
            if (mtfRegions == null) mtfRegions = new Rectangle[5];
            if (colorRegions == null) colorRegions = new Rectangle[3];
            if (mtfRegionVisible == null) mtfRegionVisible = new bool[5];
            if (colorRegionVisible == null) colorRegionVisible = new bool[3];

            // 初始化 MTF 區域 (紅框)
            mtfRegions[0] = new Rectangle(100, 100, mtfSize, mtfSize); // M0
            mtfRegions[1] = new Rectangle(100, 400, mtfSize, mtfSize); // M1
            mtfRegions[2] = new Rectangle(600, 400, mtfSize, mtfSize); // M2
            mtfRegions[3] = new Rectangle(600, 100, mtfSize, mtfSize); // M3
            mtfRegions[4] = new Rectangle(350, 250, mtfSize, mtfSize); // MC

            // 初始化 Color 區域 (黑框)
            colorRegions[0] = new Rectangle(300, 350, mtfSize, mtfSize); // C1
            colorRegions[1] = new Rectangle(200, 450, mtfSize, mtfSize); // C2
            colorRegions[2] = new Rectangle(500, 450, mtfSize, mtfSize); // C3

            // 預設全部顯示
            for (int i = 0; i < 5; i++) mtfRegionVisible[i] = true;
            for (int i = 0; i < 3; i++) colorRegionVisible[i] = true;

            System.Diagnostics.Debug.WriteLine("MTF 區域初始化完成");
        }

        private void InitializeMTFGroups()
        {
            string[] groupNames = { "M0", "M1", "M2", "M3", "MC" };

            Point startPos = new Point(850, 200);
            int groupWidth = 180;
            int groupHeight = 130;
            int horizontalSpacing = 190;
            int verticalSpacing = 140;

            // 初始化資料陣列
            for (int i = 0; i < 5; i++)
            {
                mtfHistoryData[i] = new List<double>();
                mtfTimeStamps[i] = new List<DateTime>();
            }

            // 建立並啟動更新計時器
            chartUpdateTimer = new Timer();
            chartUpdateTimer.Interval = 1000; // 1秒更新一次
            chartUpdateTimer.Tick += ChartUpdateTimer_Tick;
            chartUpdateTimer.Start();

            for (int i = 0; i < 5; i++)
            {
                GroupBox group = new GroupBox();
                group.Text = groupNames[i];
                group.Size = new Size(groupWidth, groupHeight);

                int col = i % 3;
                int row = i / 3;

                group.Location = new Point(
                    startPos.X + col * horizontalSpacing,
                    startPos.Y + row * verticalSpacing
                );

                group.Visible = true;

                // 灰階圖像 PictureBox
                PictureBox pb = new PictureBox();
                pb.Size = new Size(70, 50);
                pb.Location = new Point(10, 20);
                pb.BackColor = Color.Gray;
                pb.SizeMode = PictureBoxSizeMode.Zoom;
                group.Controls.Add(pb);
                mtfPictureBoxes[i] = pb;

                // MTF 圖表 PictureBox - 修正事件註冊方式
                PictureBox chart = new PictureBox();
                chart.Size = new Size(90, 50);
                chart.Location = new Point(85, 20);
                chart.BackColor = Color.White;
                chart.BorderStyle = BorderStyle.FixedSingle;
                chart.Tag = i; // 使用 Tag 屬性儲存索引
                chart.Paint += Chart_Paint; // 統一的事件處理
                group.Controls.Add(chart);
                mtfChartBoxes[i] = chart;

                // 當前值標籤
                Label currentLabel = new Label();
                currentLabel.Text = "0.00";
                currentLabel.Location = new Point(10, 80);
                currentLabel.Size = new Size(50, 20);
                currentLabel.ForeColor = Color.Red;
                group.Controls.Add(currentLabel);
                mtfCurrentLabels[i] = currentLabel;

                // 最大值標籤
                Label maxLabel = new Label();
                maxLabel.Text = "0.00";
                maxLabel.Location = new Point(70, 80);
                maxLabel.Size = new Size(50, 20);
                maxLabel.ForeColor = Color.Blue;
                group.Controls.Add(maxLabel);
                mtfMaxLabels[i] = maxLabel;

                this.Controls.Add(group);
                mtfGroups[i] = group;
            }
        }

        // 計時器事件處理
        // 在計時器中加入顏色分析
        private void ChartUpdateTimer_Tick(object sender, EventArgs e)
        {
            try
            {
                // MTF 分析
                for (int i = 0; i < 5; i++)
                {
                    if (mtfRegionVisible[i] && currentImage != null)
                    {
                        PerformSingleMTFAnalysis(i);
                    }
                }

                // 顏色分析 - 每秒更新一次
                if (currentImage != null)
                {
                    PerformColorAnalysis();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("計時器更新失敗: " + ex.Message);
            }
        }

        // 單一 MTF 區域分析
        // 修正 PerformSingleMTFAnalysis 方法，支援兩種 MTF 計算模式
        private void PerformSingleMTFAnalysis(int regionIndex)
        {
            try
            {
                if (regionIndex < 0 || regionIndex >= 5) return;

                Bitmap imageToAnalyze = null;

                // 安全地取得當前影像
                lock (imageLock)
                {
                    if (currentImage == null) return;

                    try
                    {
                        imageToAnalyze = new Bitmap(currentImage);
                    }
                    catch (Exception)
                    {
                        return;
                    }
                }

                if (imageToAnalyze == null) return;

                try
                {
                    Rectangle region = mtfRegions[regionIndex];
                    if (region.X >= 0 && region.Y >= 0 &&
                        region.Right <= imageToAnalyze.Width &&
                        region.Bottom <= imageToAnalyze.Height)
                    {
                        Bitmap regionImage = imageToAnalyze.Clone(region, imageToAnalyze.PixelFormat);
                        Bitmap grayImage = ConvertToGrayscale(regionImage);

                        // 根據選擇的方法計算 MTF 值
                        double mtfValue;
                        switch (currentMTFMethod)
                        {
                            case MTFMethod.MTF50:
                                mtfValue = CalculateMTF50(grayImage);
                                break;
                            case MTFMethod.PixelDifference:
                            default:
                                mtfValue = CalculateMTF(grayImage); // 原來的方法
                                break;
                        }

                        // 更新歷史資料
                        UpdateMTFHistory(regionIndex, mtfValue);

                        // 更新顯示 (在 UI 執行緒中)
                        this.Invoke(new MethodInvoker(delegate ()
                        {
                            if (mtfPictureBoxes[regionIndex].Image != null)
                                mtfPictureBoxes[regionIndex].Image.Dispose();

                            mtfPictureBoxes[regionIndex].Image = grayImage;
                            mtfCurrentLabels[regionIndex].Text = mtfValue.ToString("F2");

                            // 計算最大值
                            double maxValue = mtfHistoryData[regionIndex].Count > 0 ?
                                             mtfHistoryData[regionIndex].Max() : mtfValue;
                            mtfMaxLabels[regionIndex].Text = maxValue.ToString("F2");

                            // 更新圖表顯示
                            mtfChartBoxes[regionIndex].Invalidate();
                        }));

                        regionImage.Dispose();
                    }
                }
                finally
                {
                    imageToAnalyze.Dispose();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("MTF 分析錯誤: " + ex.Message);
            }
        }



        // 更新 MTF 歷史資料
        private void UpdateMTFHistory(int regionIndex, double mtfValue)
        {
            DateTime now = DateTime.Now;

            // 加入新資料
            mtfHistoryData[regionIndex].Add(mtfValue);
            mtfTimeStamps[regionIndex].Add(now);

            // 移除超過時間範圍的舊資料
            while (mtfHistoryData[regionIndex].Count > MAX_DATA_POINTS)
            {
                mtfHistoryData[regionIndex].RemoveAt(0);
                mtfTimeStamps[regionIndex].RemoveAt(0);
            }
        }

        // 圖表繪製方法
        private void Chart_Paint(object sender, PaintEventArgs e)
        {
            PictureBox chart = sender as PictureBox;
            Graphics g = e.Graphics;

            // 從 Tag 取得區域索引，並進行安全檢查
            if (chart.Tag == null) return;

            int regionIndex;
            if (!int.TryParse(chart.Tag.ToString(), out regionIndex)) return;

            // 檢查索引是否在有效範圍內
            if (regionIndex < 0 || regionIndex >= 5) return;

            // 檢查資料陣列是否已初始化
            if (mtfHistoryData == null || mtfHistoryData[regionIndex] == null) return;

            // 清除背景
            g.Clear(Color.White);

            // 檢查是否有足夠的資料點
            if (mtfHistoryData[regionIndex].Count < 2)
            {
                // 繪製 "No Data" 提示
                using (Font font = new Font("Arial", 8))
                {
                    using (Brush textBrush = new SolidBrush(Color.Gray))
                    {
                        string noDataText = "No Data";
                        SizeF textSize = g.MeasureString(noDataText, font);
                        float x = (chart.Width - textSize.Width) / 2;
                        float y = (chart.Height - textSize.Height) / 2;
                        g.DrawString(noDataText, font, textBrush, x, y);
                    }
                }
                return;
            }

            // 設定繪圖參數
            int margin = 5;
            int chartWidth = chart.Width - 2 * margin;
            int chartHeight = chart.Height - 2 * margin;

            // 計算數值範圍
            double minValue = mtfHistoryData[regionIndex].Min();
            double maxValue = mtfHistoryData[regionIndex].Max();
            double valueRange = maxValue - minValue;

            if (valueRange == 0) valueRange = 1; // 避免除以零

            // 準備繪圖點
            List<PointF> points = new List<PointF>();

            for (int i = 0; i < mtfHistoryData[regionIndex].Count; i++)
            {
                float x = margin + (float)i * chartWidth / Math.Max(1, (MAX_DATA_POINTS - 1));
                float y = margin + chartHeight - (float)((mtfHistoryData[regionIndex][i] - minValue) / valueRange) * chartHeight;
                points.Add(new PointF(x, y));
            }

            // 繪製格線
            using (Pen gridPen = new Pen(Color.LightGray, 1))
            {
                // 垂直格線 (時間軸)
                for (int i = 0; i <= 6; i++) // 每10秒一條線
                {
                    float x = margin + (float)i * chartWidth / 6;
                    g.DrawLine(gridPen, x, margin, x, margin + chartHeight);
                }

                // 水平格線 (數值軸)
                for (int i = 0; i <= 4; i++)
                {
                    float y = margin + (float)i * chartHeight / 4;
                    g.DrawLine(gridPen, margin, y, margin + chartWidth, y);
                }
            }

            // 繪製折線
            if (points.Count >= 2)
            {
                using (Pen linePen = new Pen(Color.Blue, 2))
                {
                    try
                    {
                        g.DrawLines(linePen, points.ToArray());
                    }
                    catch (Exception)
                    {
                        // 如果繪製失敗，跳過折線繪製
                    }
                }

                // 繪製數據點
                using (Brush pointBrush = new SolidBrush(Color.Red))
                {
                    foreach (PointF point in points)
                    {
                        if (point.X >= 0 && point.Y >= 0 && point.X <= chart.Width && point.Y <= chart.Height)
                        {
                            g.FillEllipse(pointBrush, point.X - 1, point.Y - 1, 2, 2);
                        }
                    }
                }
            }

            // 繪製數值標籤
            using (Font font = new Font("Arial", 6))
            {
                using (Brush textBrush = new SolidBrush(Color.Black))
                {
                    // Y軸最大值
                    g.DrawString(maxValue.ToString("F1"), font, textBrush, 2, margin - 2);
                    // Y軸最小值
                    g.DrawString(minValue.ToString("F1"), font, textBrush, 2, margin + chartHeight - 8);
                }
            }
        }


        // 修正 InitializeColorLabels 方法，設定初始顯示
        private void InitializeColorLabels()
        {
            for (int i = 0; i < 3; i++)
            {
                Label label = new Label();
                label.Name = "lblColor" + (i + 1);
                label.Text = $"C{i + 1} -- L:0.0, a:0.0, b:0.0";
                label.Location = new Point(850, 100 + i * 30);
                label.Size = new Size(200, 20);
                this.Controls.Add(label);
                colorLabels[i] = label;
            }
        }

        // 修正 Button_Click 方法，加入 Save Stop 功能
        // 修正 Button_Click 方法，Exit 時正確儲存設定
        private void Button_Click(object sender, EventArgs e)
        {
            Button btn = sender as Button;

            switch (btn.Name)
            {
                case "btnStart":
                    StartCamera();
                    break;
                case "btnOpen":
                    OpenImageFile();
                    break;
                case "btnStop":
                    StopCamera();
                    break;
                case "btnSaveStop":
                    StopCamera();
                    SaveCurrentSettings();
                    break;
                case "btnSaveImage":
                    SaveImage();
                    break;
                case "btnVideo":
                    ShowVideoProperties();
                    break;
                case "btnExit":
                    SaveCurrentSettingsWithCriterion(); // 使用新的儲存方法
                    this.Close();
                    break;
            }
        }


        // 儲存當前設定
        // 修正 SaveCurrentSettings 方法，也使用正確的 Criterion 值
        private void SaveCurrentSettings()
        {
            try
            {
                if (string.IsNullOrEmpty(currentSettingsFilePath))
                {
                    // 如果沒有設定檔路徑，建立新檔
                    string exePath = Application.StartupPath;
                    string fileName = $"IASetsup_{DateTime.Now:yyyyMMdd}.set";
                    currentSettingsFilePath = System.IO.Path.Combine(exePath, fileName);
                }

                // 使用完整設定而不是簡化版本
                ImageAnalysisSettings currentSettings = GetCurrentCompleteSettings();
                SaveSettingsToFile(currentSettings, currentSettingsFilePath);

                MessageBox.Show($"設定已儲存到: {System.IO.Path.GetFileName(currentSettingsFilePath)}");

                System.Diagnostics.Debug.WriteLine($"Save Stop 儲存設定完成");
                System.Diagnostics.Debug.WriteLine($"儲存的 Criterion 值: M0={currentSettings.M0_Criterion}, M1={currentSettings.M1_Criterion}, M2={currentSettings.M2_Criterion}, M3={currentSettings.M3_Criterion}, MC={currentSettings.MC_Criterion}");
            }
            catch (Exception ex)
            {
                MessageBox.Show("儲存設定失敗: " + ex.Message);
            }
        }

        // 公開的設定存取方法 (原本的預留功能)
        public void SaveSettings()
        {
            SaveCurrentSettings();
        }

        public void LoadSettings()
        {
            SelectAndLoadSettingsFile();
        }

        private void StartCamera()
        {
            try
            {
                if (videoDevices != null && videoDevices.Count > 0)
                {
                    ComboBox cmbDevices = this.Controls.Find("cmbDevices", true)[0] as ComboBox;
                    int selectedIndex = cmbDevices.SelectedIndex;

                    if (selectedIndex >= 0 && selectedIndex < videoDevices.Count)
                    {
                        videoDevice = new VideoCaptureDevice(videoDevices[selectedIndex].MonikerString);
                        videoDevice.NewFrame += VideoDevice_NewFrame;
                        videoDevice.Start();
                        isDeviceRunning = true;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("啟動攝影機失敗: " + ex.Message);
            }
        }

        private void StopCamera()
        {
            try
            {
                if (videoDevice != null && isDeviceRunning)
                {
                    videoDevice.SignalToStop();
                    videoDevice.WaitForStop();
                    videoDevice = null;
                    isDeviceRunning = false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("停止攝影機失敗: " + ex.Message);
            }
        }

        private void VideoDevice_NewFrame(object sender, NewFrameEventArgs eventArgs)
        {
            try
            {
                // 在背景執行緒中複製影像
                Bitmap newFrame = (Bitmap)eventArgs.Frame.Clone();

                // 在 UI 執行緒中更新顯示
                this.Invoke(new MethodInvoker(delegate ()
                {
                    lock (imageLock)
                    {
                        // 釋放舊的影像
                        if (currentImage != null)
                        {
                            currentImage.Dispose();
                        }

                        currentImage = newFrame;
                    }

                    PictureBox pb = this.Controls.Find("pictureBox1", true)[0] as PictureBox;
                    pb.Invalidate(); // 觸發重繪
                }));
            }
            catch (Exception ex)
            {
                // 處理例外，避免程式崩潰
                System.Diagnostics.Debug.WriteLine("VideoDevice_NewFrame 錯誤: " + ex.Message);
            }
        }

        // 修正開啟圖片檔案方法，加入執行緒安全
        private void OpenImageFile()
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Image Files|*.bmp;*.jpg;*.jpeg;*.png";

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    Bitmap newImage = new Bitmap(ofd.FileName);

                    lock (imageLock)
                    {
                        // 釋放舊的影像
                        if (currentImage != null)
                            currentImage.Dispose();

                        currentImage = newImage;
                    }

                    PictureBox pb = this.Controls.Find("pictureBox1", true)[0] as PictureBox;
                    pb.Invalidate();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("載入圖片失敗: " + ex.Message);
                }
            }
        }


        // 修正 SaveImage 方法，加入執行緒安全
        private void SaveImage()
        {
            try
            {
                Bitmap imageToSave = null;

                // 安全地取得當前影像
                lock (imageLock)
                {
                    if (currentImage != null)
                    {
                        try
                        {
                            imageToSave = new Bitmap(currentImage);
                        }
                        catch (Exception)
                        {
                            MessageBox.Show("無法存取當前影像");
                            return;
                        }
                    }
                }

                if (imageToSave != null)
                {
                    SaveFileDialog sfd = new SaveFileDialog();
                    sfd.Filter = "JPEG Files|*.jpg|BMP Files|*.bmp";
                    sfd.DefaultExt = "jpg";
                    sfd.FileName = "image_" + DateTime.Now.ToString("yyyyMMdd_HHmmss");

                    if (sfd.ShowDialog() == DialogResult.OK)
                    {
                        // 取得 Output Ratio 設定
                        ComboBox cmbOutputRatio = this.Controls.Find("cmbOutputRatio", true)[0] as ComboBox;
                        int outputRatio = int.Parse(cmbOutputRatio.Text);

                        // 計算新的尺寸
                        int newWidth = (int)(imageToSave.Width * outputRatio / 100.0);
                        int newHeight = (int)(imageToSave.Height * outputRatio / 100.0);

                        // 建立縮放後的圖像
                        Bitmap scaledImage = new Bitmap(newWidth, newHeight);
                        using (Graphics g = Graphics.FromImage(scaledImage))
                        {
                            g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                            g.DrawImage(imageToSave, 0, 0, newWidth, newHeight);
                        }

                        // 根據選擇的格式儲存
                        string extension = Path.GetExtension(sfd.FileName).ToLower();
                        switch (extension)
                        {
                            case ".jpg":
                            case ".jpeg":
                                SaveAsJPEG(scaledImage, sfd.FileName);
                                break;
                            case ".bmp":
                                scaledImage.Save(sfd.FileName, System.Drawing.Imaging.ImageFormat.Bmp);
                                break;
                            default:
                                scaledImage.Save(sfd.FileName, System.Drawing.Imaging.ImageFormat.Jpeg);
                                break;
                        }

                        scaledImage.Dispose();
                        imageToSave.Dispose();

                        // 儲存螢幕截圖
                        SaveScreenshot(sfd.FileName);

                        MessageBox.Show($"圖片已儲存: {sfd.FileName}\n" +
                                      $"原始尺寸: {imageToSave.Width}x{imageToSave.Height}\n" +
                                      $"儲存尺寸: {newWidth}x{newHeight} ({outputRatio}%)\n" +
                                      $"螢幕截圖也已一併儲存");
                    }
                }
                else
                {
                    MessageBox.Show("沒有圖像可以儲存");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("儲存圖片失敗: " + ex.Message);
            }
        }

        private void SaveAsJPEG(Bitmap image, string filename)
        {
            // 設定 JPEG 品質參數
            System.Drawing.Imaging.ImageCodecInfo jpegCodec = GetEncoder(System.Drawing.Imaging.ImageFormat.Jpeg);
            System.Drawing.Imaging.Encoder encoder = System.Drawing.Imaging.Encoder.Quality;
            System.Drawing.Imaging.EncoderParameters encoderParams = new System.Drawing.Imaging.EncoderParameters(1);
            System.Drawing.Imaging.EncoderParameter encoderParam = new System.Drawing.Imaging.EncoderParameter(encoder, 90L); // 90% 品質
            encoderParams.Param[0] = encoderParam;

            image.Save(filename, jpegCodec, encoderParams);

            encoderParams.Dispose();
        }

        private System.Drawing.Imaging.ImageCodecInfo GetEncoder(System.Drawing.Imaging.ImageFormat format)
        {
            System.Drawing.Imaging.ImageCodecInfo[] codecs = System.Drawing.Imaging.ImageCodecInfo.GetImageDecoders();
            foreach (System.Drawing.Imaging.ImageCodecInfo codec in codecs)
            {
                if (codec.FormatID == format.Guid)
                {
                    return codec;
                }
            }
            return null;
        }




        private void SaveScreenshot(string originalFilename)
        {
            try
            {
                // 建立螢幕截圖檔名
                string directory = Path.GetDirectoryName(originalFilename);
                string filenameWithoutExt = Path.GetFileNameWithoutExtension(originalFilename);
                string screenshotFilename = Path.Combine(directory, filenameWithoutExt + "_screenshot.png");

                // 方法1: 使用 Form 的 DrawToBitmap 方法 (推薦)
                Rectangle formBounds = this.Bounds;
                Bitmap screenshot = new Bitmap(formBounds.Width, formBounds.Height);
                this.DrawToBitmap(screenshot, new Rectangle(0, 0, formBounds.Width, formBounds.Height));

                // 儲存為 PNG 格式
                screenshot.Save(screenshotFilename, System.Drawing.Imaging.ImageFormat.Png);
                screenshot.Dispose();
            }
            catch (Exception ex)
            {
                // 如果方法1失敗，嘗試方法2
                try
                {
                    SaveScreenshotAlternative(originalFilename);
                }
                catch
                {
                    MessageBox.Show("儲存螢幕截圖失敗: " + ex.Message);
                }
            }
        }

        private void SaveScreenshotAlternative(string originalFilename)
        {
            // 建立螢幕截圖檔名
            string directory = Path.GetDirectoryName(originalFilename);
            string filenameWithoutExt = Path.GetFileNameWithoutExtension(originalFilename);
            string screenshotFilename = Path.Combine(directory, filenameWithoutExt + "_screenshot.png");

            // 方法2: 直接從螢幕擷取，處理最大化狀態
            Rectangle bounds;

            if (this.WindowState == FormWindowState.Maximized)
            {
                // 最大化時使用工作區域
                bounds = Screen.FromControl(this).WorkingArea;
            }
            else
            {
                // 一般狀態使用 Form 邊界
                bounds = this.Bounds;
            }

            Bitmap screenshot = new Bitmap(bounds.Width, bounds.Height);

            using (Graphics g = Graphics.FromImage(screenshot))
            {
                // 確保從正確的位置開始擷取
                Point capturePoint = this.WindowState == FormWindowState.Maximized ?
                                   new Point(0, 0) :
                                   this.Location;

                g.CopyFromScreen(capturePoint, Point.Empty, bounds.Size);
            }

            screenshot.Save(screenshotFilename, System.Drawing.Imaging.ImageFormat.Png);
            screenshot.Dispose();
        }

        // 如果還是有問題，可以使用這個更完整的方法
        private void SaveFullScreenshot(string originalFilename)
        {
            try
            {
                // 建立螢幕截圖檔名
                string directory = Path.GetDirectoryName(originalFilename);
                string filenameWithoutExt = Path.GetFileNameWithoutExtension(originalFilename);
                string screenshotFilename = Path.Combine(directory, filenameWithoutExt + "_screenshot.png");

                // 暫時將 Form 帶到最前面
                this.TopMost = true;
                this.Focus();
                Application.DoEvents();
                System.Threading.Thread.Sleep(100); // 等待畫面更新

                // 取得實際的 Form 大小和位置
                Rectangle formRect = this.RectangleToScreen(this.ClientRectangle);

                // 如果是最大化，調整為整個螢幕
                if (this.WindowState == FormWindowState.Maximized)
                {
                    formRect = this.Bounds;
                }

                Bitmap screenshot = new Bitmap(formRect.Width, formRect.Height);

                using (Graphics g = Graphics.FromImage(screenshot))
                {
                    g.CopyFromScreen(formRect.Location, Point.Empty, formRect.Size);
                }

                // 恢復 TopMost 設定
                this.TopMost = false;

                screenshot.Save(screenshotFilename, System.Drawing.Imaging.ImageFormat.Png);
                screenshot.Dispose();
            }
            catch (Exception ex)
            {
                this.TopMost = false; // 確保恢復設定
                MessageBox.Show("儲存螢幕截圖失敗: " + ex.Message);
            }
        }

        private void ShowVideoProperties()
        {
            try
            {
                if (videoDevice != null)
                {
                    videoDevice.DisplayPropertyPage(IntPtr.Zero);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("顯示攝影機設定失敗: " + ex.Message);
            }
        }

        // 修正 PictureBox1_Paint 方法，加入執行緒安全
        private void PictureBox1_Paint(object sender, PaintEventArgs e)
        {
            PictureBox pb = sender as PictureBox;
            Graphics g = e.Graphics;

            // 清除背景
            g.Clear(pb.BackColor);

            Bitmap imageToUse = null;

            // 安全地複製當前影像參考
            lock (imageLock)
            {
                if (currentImage != null)
                {
                    // 建立當前影像的複本以避免多執行緒問題
                    try
                    {
                        imageToUse = new Bitmap(currentImage);
                    }
                    catch (Exception)
                    {
                        // 如果複製失敗，直接返回
                        return;
                    }
                }
            }

            // 如果沒有影像，只繪製背景
            if (imageToUse == null) return;

            try
            {
                // 計算基礎縮放比例 (將原圖適應到 PictureBox)
                float baseScaleX = (float)pb.Width / imageToUse.Width;
                float baseScaleY = (float)pb.Height / imageToUse.Height;
                float baseScale = Math.Min(baseScaleX, baseScaleY);

                // 套用 PB1 Ratio
                float finalScale = baseScale * currentPB1Ratio;
                imageScale = finalScale;

                // 計算縮放後的圖像尺寸
                int scaledWidth = (int)(imageToUse.Width * finalScale);
                int scaledHeight = (int)(imageToUse.Height * finalScale);

                // 計算基礎偏移 (居中顯示)
                int baseOffsetX = (pb.Width - (int)(imageToUse.Width * baseScale)) / 2;
                int baseOffsetY = (pb.Height - (int)(imageToUse.Height * baseScale)) / 2;

                // 計算最終顯示位置
                int displayX = baseOffsetX + (int)imageOffset.X;
                int displayY = baseOffsetY + (int)imageOffset.Y;

                // 設定高品質繪製
                g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;

                // 繪製影像
                Rectangle destRect = new Rectangle(displayX, displayY, scaledWidth, scaledHeight);
                g.DrawImage(imageToUse, destRect);

                // 繪製 MTF 和顏色區域
                DrawMTFAndColorRegions(g, pb, finalScale, displayX, displayY);
            }
            catch (Exception ex)
            {
                // 繪製錯誤處理
                System.Diagnostics.Debug.WriteLine("Paint 錯誤: " + ex.Message);
            }
            finally
            {
                // 釋放複製的影像
                if (imageToUse != null)
                {
                    imageToUse.Dispose();
                }
            }
        }




        private void PictureBox1_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                PictureBox pb = sender as PictureBox;

                // 檢查是否為影像拖曳模式 (75% 或 50%)
                if (currentPB1Ratio > 1.0f)
                {
                    // 優先檢查是否點擊在區域上
                    Point imagePoint = GetImageCoordinates(e.Location, pb);
                    bool clickedOnRegion = false;

                    // 檢查 MTF 區域
                    for (int i = 0; i < 5; i++)
                    {
                        if (mtfRegionVisible[i] && mtfRegions[i].Contains(imagePoint))
                        {
                            isDragging = true;
                            isDragMTF = true;
                            dragRegionIndex = i;
                            dragStartPoint = imagePoint;
                            clickedOnRegion = true;
                            return;
                        }
                    }

                    // 檢查 Color 區域
                    for (int i = 0; i < 3; i++)
                    {
                        if (colorRegionVisible[i] && colorRegions[i].Contains(imagePoint))
                        {
                            isDragging = true;
                            isDragMTF = false;
                            dragRegionIndex = i;
                            dragStartPoint = imagePoint;
                            clickedOnRegion = true;
                            return;
                        }
                    }

                    // 如果沒點擊在區域上，則開始影像拖曳
                    if (!clickedOnRegion)
                    {
                        isImageDragging = true;
                        imageDragStartPoint = e.Location;
                    }
                }
                else
                {
                    // 100% 模式下的原有區域拖曳邏輯
                    Point imagePoint = GetImageCoordinates(e.Location, pb);

                    // 檢查 MTF 區域
                    for (int i = 0; i < 5; i++)
                    {
                        if (mtfRegionVisible[i] && mtfRegions[i].Contains(imagePoint))
                        {
                            isDragging = true;
                            isDragMTF = true;
                            dragRegionIndex = i;
                            dragStartPoint = imagePoint;
                            return;
                        }
                    }

                    // 檢查 Color 區域
                    for (int i = 0; i < 3; i++)
                    {
                        if (colorRegionVisible[i] && colorRegions[i].Contains(imagePoint))
                        {
                            isDragging = true;
                            isDragMTF = false;
                            dragRegionIndex = i;
                            dragStartPoint = imagePoint;
                            return;
                        }
                    }
                }
            }
        }

        private void PictureBox1_MouseMove(object sender, MouseEventArgs e)
        {
            PictureBox pb = sender as PictureBox;
            Point imagePoint = GetImageCoordinates(e.Location, pb);

            // 更新滑鼠位置顯示
            UpdatePointLabels(e.Location, imagePoint, pb);

            if (isImageDragging)
            {
                // 影像拖曳
                Point offset = new Point(
                    e.Location.X - imageDragStartPoint.X,
                    e.Location.Y - imageDragStartPoint.Y
                );

                imageOffset = new PointF(
                    imageOffset.X + offset.X,
                    imageOffset.Y + offset.Y
                );

                imageDragStartPoint = e.Location;
                pb.Invalidate();
            }
            else if (isDragging)
            {
                // 區域拖曳
                Point offset = new Point(
                    imagePoint.X - dragStartPoint.X,
                    imagePoint.Y - dragStartPoint.Y
                );

                if (isDragMTF)
                {
                    mtfRegions[dragRegionIndex] = new Rectangle(
                        mtfRegions[dragRegionIndex].X + offset.X,
                        mtfRegions[dragRegionIndex].Y + offset.Y,
                        mtfRegions[dragRegionIndex].Width,
                        mtfRegions[dragRegionIndex].Height
                    );
                }
                else
                {
                    colorRegions[dragRegionIndex] = new Rectangle(
                        colorRegions[dragRegionIndex].X + offset.X,
                        colorRegions[dragRegionIndex].Y + offset.Y,
                        colorRegions[dragRegionIndex].Width,
                        colorRegions[dragRegionIndex].Height
                    );
                }

                dragStartPoint = imagePoint;
                pb.Invalidate();
            }
        }


        // 修正 PictureBox1_MouseUp 方法，拖曳結束後立即更新顏色分析
        // 修正 PictureBox1_MouseUp 方法，加入同步通知
        private void PictureBox1_MouseUp(object sender, MouseEventArgs e)
        {
            if (isImageDragging)
            {
                isImageDragging = false;
            }
            else if (isDragging)
            {
                isDragging = false;
                int tempRegionIndex = dragRegionIndex;
                bool tempIsDragMTF = isDragMTF;
                dragRegionIndex = -1;

                // 執行對應的分析
                if (tempIsDragMTF)
                {
                    // MTF 區域被移動，執行 MTF 分析
                    PerformSingleMTFAnalysis(tempRegionIndex);

                    // 通知 MTF Setting Form 區域位置變更
                    if (mtfSettingForm != null && !mtfSettingForm.IsDisposed && !isFormSyncing)
                    {
                        Point newCenter = new Point(
                            mtfRegions[tempRegionIndex].X + mtfRegions[tempRegionIndex].Width / 2,
                            mtfRegions[tempRegionIndex].Y + mtfRegions[tempRegionIndex].Height / 2
                        );
                        MTFRegionMoved?.Invoke(tempRegionIndex, newCenter);
                    }
                }
                else
                {
                    // 顏色區域被移動，執行顏色分析
                    PerformColorAnalysis();
                }
            }
        }

        // 修正 GetImageCoordinates 方法，加入執行緒安全
        private Point GetImageCoordinates(Point pictureBoxPoint, PictureBox pb)
        {
            lock (imageLock)
            {
                if (currentImage == null) return Point.Empty;

                try
                {
                    // 計算基礎縮放
                    float baseScaleX = (float)pb.Width / currentImage.Width;
                    float baseScaleY = (float)pb.Height / currentImage.Height;
                    float baseScale = Math.Min(baseScaleX, baseScaleY);

                    // 套用 PB1 Ratio
                    float finalScale = baseScale * currentPB1Ratio;

                    // 計算偏移
                    int baseOffsetX = (pb.Width - (int)(currentImage.Width * baseScale)) / 2;
                    int baseOffsetY = (pb.Height - (int)(currentImage.Height * baseScale)) / 2;

                    int displayOffsetX = baseOffsetX + (int)imageOffset.X;
                    int displayOffsetY = baseOffsetY + (int)imageOffset.Y;

                    // 轉換座標
                    int imageX = (int)((pictureBoxPoint.X - displayOffsetX) / finalScale);
                    int imageY = (int)((pictureBoxPoint.Y - displayOffsetY) / finalScale);

                    return new Point(imageX, imageY);
                }
                catch (Exception)
                {
                    return Point.Empty;
                }
            }
        }

        private void UpdatePointLabels(Point pictureBoxPoint, Point imagePoint, PictureBox pb)
        {
            try
            {
                Label lblPoint = this.Controls.Find("lblPoint", true)[0] as Label;
                Label lblRealPoint = this.Controls.Find("lblRealPoint", true)[0] as Label;
                Label lblScreenScale = this.Controls.Find("lblScreenScale", true)[0] as Label;

                if (pb.Image != null)
                {
                    float ratioX = (float)imagePoint.X / pb.Image.Width;
                    float ratioY = (float)imagePoint.Y / pb.Image.Height;

                    lblPoint.Text = $"Point: {pictureBoxPoint.X}, {pictureBoxPoint.Y}, {ratioX:F2}, {ratioY:F2}";
                    lblRealPoint.Text = $"Real Point: {imagePoint.X}, {imagePoint.Y}";
                    lblScreenScale.Text = $"Screen Scale: {imageScale:F2}";
                }
            }
            catch { }
        }

        private void CmbMTFArea_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox cmb = sender as ComboBox;
            int newSize = int.Parse(cmb.Text);

            // 更新所有區域尺寸
            for (int i = 0; i < 5; i++)
            {
                mtfRegions[i] = new Rectangle(
                    mtfRegions[i].X, mtfRegions[i].Y, newSize, newSize);
            }

            for (int i = 0; i < 3; i++)
            {
                colorRegions[i] = new Rectangle(
                    colorRegions[i].X, colorRegions[i].Y, newSize, newSize);
            }

            PictureBox pb = this.Controls.Find("pictureBox1", true)[0] as PictureBox;
            pb.Invalidate();
        }

        private void MtfCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox chk = sender as CheckBox;
            string[] mtfNames = { "M0", "M1", "M2", "M3", "MC" };

            for (int i = 0; i < mtfNames.Length; i++)
            {
                if (chk.Name == "chk" + mtfNames[i])
                {
                    mtfRegionVisible[i] = chk.Checked;
                    mtfGroups[i].Visible = chk.Checked;
                    break;
                }
            }

            PictureBox pb = this.Controls.Find("pictureBox1", true)[0] as PictureBox;
            pb.Invalidate();
        }

        private void ColorCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox chk = sender as CheckBox;
            string[] colorNames = { "C1", "C2", "C3" };

            for (int i = 0; i < colorNames.Length; i++)
            {
                if (chk.Name == "chk" + colorNames[i])
                {
                    colorRegionVisible[i] = chk.Checked;
                    break;
                }
            }

            PictureBox pb = this.Controls.Find("pictureBox1", true)[0] as PictureBox;
            pb.Invalidate();
        }

        private void PerformMTFAnalysis()
        {
            // 這個方法現在主要用於手動觸發分析
            // 實際的即時分析由 ChartUpdateTimer_Tick 處理
            try
            {
                if (currentImage == null) return;

                for (int i = 0; i < 5; i++)
                {
                    if (mtfRegionVisible[i])
                    {
                        PerformSingleMTFAnalysis(i);
                    }
                }

                // 更新整體 MTF 結果顯示
                UpdateMTFResultDisplay();
            }
            catch (Exception ex)
            {
                MessageBox.Show("MTF 分析失敗: " + ex.Message);
            }
        }

        // 修正 PerformColorAnalysis 方法，計算實際的 Lab 色彩值
        // 修正 PerformColorAnalysis 方法，整合使用 IsValidColorRegion 和 LogColorAnalysis
        private void PerformColorAnalysis()
        {
            try
            {
                Bitmap imageToAnalyze = null;

                // 安全地取得當前影像
                lock (imageLock)
                {
                    if (currentImage == null) return;

                    try
                    {
                        imageToAnalyze = new Bitmap(currentImage);
                    }
                    catch (Exception)
                    {
                        return;
                    }
                }

                if (imageToAnalyze == null) return;

                try
                {
                    for (int i = 0; i < 3; i++)
                    {
                        if (colorRegionVisible[i])
                        {
                            Rectangle region = colorRegions[i];

                            // 使用 IsValidColorRegion 檢查區域有效性
                            if (IsValidColorRegion(region, imageToAnalyze))
                            {
                                // 計算區域內的平均 RGB 值
                                Color avgColor = CalculateAverageColor(imageToAnalyze, region);

                                // 轉換為 Lab 色彩空間
                                double[] labValues = RGBToLab(avgColor.R, avgColor.G, avgColor.B);

                                // 使用 LogColorAnalysis 記錄除錯資訊
                                LogColorAnalysis(i, avgColor, labValues);

                                // 更新顯示標籤
                                colorLabels[i].Text = $"C{i + 1} -- L:{labValues[0]:F1}, a:{labValues[1]:F1}, b:{labValues[2]:F1}";
                            }
                            else
                            {
                                // 區域無效時的處理
                                if (region.X < 0 || region.Y < 0 ||
                                    region.Right > imageToAnalyze.Width ||
                                    region.Bottom > imageToAnalyze.Height)
                                {
                                    colorLabels[i].Text = $"C{i + 1} -- Out of Range";
                                }
                                else
                                {
                                    colorLabels[i].Text = $"C{i + 1} -- Invalid Region";
                                }

                                // 記錄無效區域的除錯資訊
                                System.Diagnostics.Debug.WriteLine($"C{i + 1}: Invalid region - X:{region.X}, Y:{region.Y}, W:{region.Width}, H:{region.Height}, ImageSize:{imageToAnalyze.Width}x{imageToAnalyze.Height}");
                            }
                        }
                        else
                        {
                            // 區域不可見
                            colorLabels[i].Text = $"C{i + 1} -- Hidden";
                        }
                    }

                    // 更新整體顏色結果顯示
                    UpdateColorResultDisplay();
                }
                finally
                {
                    imageToAnalyze.Dispose();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("顏色分析失敗: " + ex.Message);
            }
        }



        private Bitmap ConvertToGrayscale(Bitmap original)
        {
            Bitmap grayscale = new Bitmap(original.Width, original.Height);

            for (int x = 0; x < original.Width; x++)
            {
                for (int y = 0; y < original.Height; y++)
                {
                    Color pixel = original.GetPixel(x, y);
                    int gray = (int)(pixel.R * 0.299 + pixel.G * 0.587 + pixel.B * 0.114);
                    Color grayColor = Color.FromArgb(gray, gray, gray);
                    grayscale.SetPixel(x, y, grayColor);
                }
            }

            return grayscale;
        }

        private double CalculateMTF(Bitmap grayImage)
        {
            // 簡化的 MTF 計算
            // 實際應用中需要更複雜的邊緣檢測和頻域分析
            double totalVariation = 0;
            int count = 0;

            for (int x = 1; x < grayImage.Width; x++)
            {
                for (int y = 1; y < grayImage.Height; y++)
                {
                    Color current = grayImage.GetPixel(x, y);
                    Color previous = grayImage.GetPixel(x - 1, y);

                    totalVariation += Math.Abs(current.R - previous.R);
                    count++;
                }
            }

            return count > 0 ? (totalVariation / count) * 0.5 : 0; // 簡化的 MTF 值
        }

        // 改進的平均顏色計算方法，加入統計資訊
        // 也可以在 CalculateAverageColor 方法中使用 IsValidColorRegion
        private Color CalculateAverageColor(Bitmap image, Rectangle region)
        {
            try
            {
                // 使用 IsValidColorRegion 進行預檢查
                if (!IsValidColorRegion(region, image))
                {
                    System.Diagnostics.Debug.WriteLine($"CalculateAverageColor: Invalid region detected");
                    return Color.Black;
                }

                long totalR = 0, totalG = 0, totalB = 0;
                int validPixelCount = 0;

                // 確保區域在影像範圍內（雙重保護）
                int startX = Math.Max(0, region.X);
                int endX = Math.Min(image.Width, region.Right);
                int startY = Math.Max(0, region.Y);
                int endY = Math.Min(image.Height, region.Bottom);

                for (int x = startX; x < endX; x++)
                {
                    for (int y = startY; y < endY; y++)
                    {
                        Color pixel = image.GetPixel(x, y);
                        totalR += pixel.R;
                        totalG += pixel.G;
                        totalB += pixel.B;
                        validPixelCount++;
                    }
                }

                if (validPixelCount > 0)
                {
                    Color result = Color.FromArgb(
                        (int)(totalR / validPixelCount),
                        (int)(totalG / validPixelCount),
                        (int)(totalB / validPixelCount)
                    );

                    System.Diagnostics.Debug.WriteLine($"CalculateAverageColor: Processed {validPixelCount} pixels");
                    return result;
                }

                System.Diagnostics.Debug.WriteLine($"CalculateAverageColor: No valid pixels found");
                return Color.Black;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"CalculateAverageColor exception: {ex.Message}");
                return Color.Black;
            }
        }


        // 改進的 RGB 到 Lab 轉換方法（使用標準 D65 白點）
        private double[] RGBToLab(int r, int g, int b)
        {
            try
            {
                // 第一步：RGB to XYZ
                double[] xyz = RGBToXYZ(r, g, b);

                // 第二步：XYZ to Lab
                double[] lab = XYZToLab(xyz[0], xyz[1], xyz[2]);

                return lab;
            }
            catch (Exception)
            {
                // 如果轉換失敗，返回預設值
                return new double[] { 0.0, 0.0, 0.0 };
            }
        }

        // RGB 到 XYZ 色彩空間轉換
        private double[] RGBToXYZ(int r, int g, int b)
        {
            // 正規化 RGB 值到 0-1 範圍
            double rNorm = r / 255.0;
            double gNorm = g / 255.0;
            double bNorm = b / 255.0;

            // 伽馬校正
            rNorm = (rNorm > 0.04045) ? Math.Pow((rNorm + 0.055) / 1.055, 2.4) : rNorm / 12.92;
            gNorm = (gNorm > 0.04045) ? Math.Pow((gNorm + 0.055) / 1.055, 2.4) : gNorm / 12.92;
            bNorm = (bNorm > 0.04045) ? Math.Pow((bNorm + 0.055) / 1.055, 2.4) : bNorm / 12.92;

            // 轉換到 XYZ（使用 sRGB 矩陣）
            double x = rNorm * 0.4124564 + gNorm * 0.3575761 + bNorm * 0.1804375;
            double y = rNorm * 0.2126729 + gNorm * 0.7151522 + bNorm * 0.0721750;
            double z = rNorm * 0.0193339 + gNorm * 0.1191920 + bNorm * 0.9503041;

            return new double[] { x, y, z };
        }

        // XYZ 到 Lab 色彩空間轉換
        private double[] XYZToLab(double x, double y, double z)
        {
            // D65 白點參考值
            const double Xn = 0.95047;
            const double Yn = 1.00000;
            const double Zn = 1.08883;

            // 正規化
            double xr = x / Xn;
            double yr = y / Yn;
            double zr = z / Zn;

            // Lab 轉換函數
            Func<double, double> f = (double t) =>
            {
                return (t > 0.008856) ? Math.Pow(t, 1.0 / 3.0) : (7.787 * t + 16.0 / 116.0);
            };

            double fx = f(xr);
            double fy = f(yr);
            double fz = f(zr);

            // 計算 Lab 值
            double L = 116.0 * fy - 16.0;
            double a = 500.0 * (fx - fy);
            double b = 200.0 * (fy - fz);

            return new double[] { L, a, b };
        }


        private void UpdateMTFResultDisplay()
        {
            Label lblMTFValues = this.Controls.Find("lblMTFValues", true)[0] as Label;
            StringBuilder sb = new StringBuilder();

            for (int i = 0; i < 5; i++)
            {
                if (mtfRegionVisible[i])
                {
                    sb.AppendLine($"MTF {i} -- {mtfCurrentLabels[i].Text}");
                }
            }

            lblMTFValues.Text = sb.ToString();
        }

        // 修正 UpdateColorResultDisplay 方法，顯示 Lab 值
        private void UpdateColorResultDisplay()
        {
            try
            {
                Label lblColorValues = this.Controls.Find("lblColorValues", true).FirstOrDefault() as Label;
                if (lblColorValues != null)
                {
                    StringBuilder sb = new StringBuilder();

                    for (int i = 0; i < 3; i++)
                    {
                        if (colorRegionVisible[i] && colorLabels[i] != null)
                        {
                            sb.AppendLine(colorLabels[i].Text);
                        }
                        else
                        {
                            sb.AppendLine($"C{i + 1} -- Hidden");
                        }
                    }

                    lblColorValues.Text = sb.ToString().TrimEnd();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("更新顏色結果顯示失敗: " + ex.Message);
            }
        }




        // 預留擴充功能的方法
        public void ShowForm2()
        {
            // 未來可以開啟第二個 Form
            // Form2 form2 = new Form2();
            // form2.Show();
        }

        public void ShowForm3()
        {
            // 未來可以開啟第三個 Form
            // Form3 form3 = new Form3();
            // form3.Show();
        }


        // Form 關閉
        // 修正 Form1_FormClosing 事件，確保正確儲存
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (chartUpdateTimer != null)
            {
                chartUpdateTimer.Stop();
                chartUpdateTimer.Dispose();
            }

            StopCamera();

            // 確保 Criterion 值正確儲存
            try
            {
                SaveCurrentSettingsWithCriterion();
                System.Diagnostics.Debug.WriteLine("FormClosing: 設定已自動儲存");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"FormClosing 儲存失敗: {ex.Message}");
            }
        }


        // PB1 Ratio 變更事件處理
        private void CmbPB1Ratio_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox cmb = sender as ComboBox;

            switch (cmb.Text)
            {
                case "50 %":
                    currentPB1Ratio = 2.0f; // 50% = 放大2倍
                    break;
                case "75 %":
                    currentPB1Ratio = 1.33f; // 75% = 放大1.33倍
                    break;
                case "100 %":
                    currentPB1Ratio = 1.0f; // 100% = 原始縮放
                    imageOffset = new PointF(0, 0); // 重置偏移
                    break;
            }

            // 重新繪製
            PictureBox pb = this.Controls.Find("pictureBox1", true)[0] as PictureBox;
            pb.Invalidate();
        }

        // 繪製 MTF 和顏色區域
        private void DrawMTFAndColorRegions(Graphics g, PictureBox pb, float scale, int offsetX, int offsetY)
        {
            // 繪製 MTF 區域 (紅框)
            Pen redPen = new Pen(Color.Red, 2);
            for (int i = 0; i < 5; i++)
            {
                if (mtfRegionVisible[i])
                {
                    Rectangle scaledRect = new Rectangle(
                        offsetX + (int)(mtfRegions[i].X * scale),
                        offsetY + (int)(mtfRegions[i].Y * scale),
                        (int)(mtfRegions[i].Width * scale),
                        (int)(mtfRegions[i].Height * scale)
                    );
                    g.DrawRectangle(redPen, scaledRect);

                    // 繪製區域標籤
                    string[] labels = { "M-0", "M-1", "M-2", "M-3", "M-C" };
                    g.DrawString(labels[i], this.Font, Brushes.Blue, scaledRect.Location);
                }
            }

            // 繪製 Color 區域 (黑框)
            Pen blackPen = new Pen(Color.Black, 2);
            for (int i = 0; i < 3; i++)
            {
                if (colorRegionVisible[i])
                {
                    Rectangle scaledRect = new Rectangle(
                        offsetX + (int)(colorRegions[i].X * scale),
                        offsetY + (int)(colorRegions[i].Y * scale),
                        (int)(colorRegions[i].Width * scale),
                        (int)(colorRegions[i].Height * scale)
                    );
                    g.DrawRectangle(blackPen, scaledRect);

                    // 繪製區域標籤
                    string label = "C-" + (i + 1);
                    g.DrawString(label, this.Font, Brushes.Black, scaledRect.Location);
                }
            }

            redPen.Dispose();
            blackPen.Dispose();
        }

        // 加入除錯輸出，幫助追蹤問題
        private void DebugOutput(string message)
        {
            System.Diagnostics.Debug.WriteLine($"[ImageAnalysis] {DateTime.Now:HH:mm:ss.fff} - {message}");
        }

        // 加入顏色分析品質檢查方法
        // 顏色分析品質檢查方法 - 放在類別的任何位置
        private bool IsValidColorRegion(Rectangle region, Bitmap image)
        {
            try
            {
                // 檢查影像是否有效
                if (image == null)
                {
                    System.Diagnostics.Debug.WriteLine("IsValidColorRegion: Image is null");
                    return false;
                }

                // 檢查區域是否完全在影像範圍內
                if (region.X < 0 || region.Y < 0 ||
                    region.Right > image.Width || region.Bottom > image.Height)
                {
                    System.Diagnostics.Debug.WriteLine($"IsValidColorRegion: Region out of bounds - Region({region.X},{region.Y},{region.Width},{region.Height}) vs Image({image.Width},{image.Height})");
                    return false;
                }

                // 檢查區域大小是否合理
                if (region.Width < 10 || region.Height < 10)
                {
                    System.Diagnostics.Debug.WriteLine($"IsValidColorRegion: Region too small - W:{region.Width}, H:{region.Height}");
                    return false;
                }

                // 檢查區域面積是否合理（不能太大）
                if (region.Width > image.Width || region.Height > image.Height)
                {
                    System.Diagnostics.Debug.WriteLine($"IsValidColorRegion: Region too large - Region({region.Width},{region.Height}) vs Image({image.Width},{image.Height})");
                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"IsValidColorRegion exception: {ex.Message}");
                return false;
            }
        }


        // 加入顏色分析除錯資訊方法
        // 顏色分析除錯資訊方法 - 放在類別的任何位置
        private void LogColorAnalysis(int regionIndex, Color avgColor, double[] labValues)
        {
            try
            {
                // 詳細的除錯資訊
                string debugInfo = $"[Color Analysis] C{regionIndex + 1}: " +
                                  $"RGB({avgColor.R:D3},{avgColor.G:D3},{avgColor.B:D3}) -> " +
                                  $"Lab(L:{labValues[0]:F2}, a:{labValues[1]:F2}, b:{labValues[2]:F2})";

                System.Diagnostics.Debug.WriteLine(debugInfo);

                // 可選：也可以寫入日誌檔案
                // WriteToLogFile(debugInfo);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"LogColorAnalysis exception: {ex.Message}");
            }
        }

        // 可選：寫入日誌檔案的方法
        private void WriteToLogFile(string message)
        {
            try
            {
                string logPath = Path.Combine(Application.StartupPath, "ColorAnalysis.log");
                string logEntry = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] {message}";
                File.AppendAllText(logPath, logEntry + Environment.NewLine);
            }
            catch (Exception)
            {
                // 忽略日誌寫入錯誤
            }
        }

        // 第三個下拉式選單的事件處理
        private void CmbDenoise3_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox cmb = sender as ComboBox;

            if (cmb.Text == "MTF Setting")
            {
                // 開啟 MTF Setting Form
                ShowMTFSettingForm();

                // 重置選項回 None
                cmb.Text = "None";
            }
        }

        // 顯示 MTF Setting Form
        // 修正 Form1 中的 ShowMTFSettingForm 方法，加入除錯資訊
        // 修正 ShowMTFSettingForm 方法，使用非模態對話框
        private void ShowMTFSettingForm()
        {
            try
            {
                System.Diagnostics.Debug.WriteLine("開始開啟 MTF Setting Form");

                // 如果已經有開啟的 MTF Setting Form，就將其帶到前面
                if (mtfSettingForm != null && !mtfSettingForm.IsDisposed)
                {
                    mtfSettingForm.BringToFront();
                    mtfSettingForm.Focus();
                    return;
                }

                // 建立新的 MTF Setting Form
                mtfSettingForm = new MTFSettingForm();

                // 設定為非模態並設定擁有者
                mtfSettingForm.Owner = this;

                // 取得當前資料
                MTFSettingData currentData = GetCurrentMTFData();
                System.Diagnostics.Debug.WriteLine($"取得 MTF 資料，區域數量: {currentData.MTFRegions.Count}");

                // 傳遞當前的 MTF 資料給 Form
                mtfSettingForm.SetMTFData(currentData);

                // 設定雙向同步事件
                mtfSettingForm.MTFDataUpdated += OnMTFSettingFormDataUpdated;
                mtfSettingForm.MTFRegionSelected += OnMTFSettingFormRegionSelected;
                mtfSettingForm.FormClosed += OnMTFSettingFormClosed;

                // 設定 Form1 的同步事件
                this.MTFDataChanged += mtfSettingForm.OnForm1DataChanged;
                this.MTFRegionMoved += mtfSettingForm.OnForm1RegionMoved;

                // 顯示為非模態對話框
                mtfSettingForm.Show();

                System.Diagnostics.Debug.WriteLine("MTF Setting Form 已開啟（非模態）");
            }
            catch (Exception ex)
            {
                MessageBox.Show("開啟 MTF 設定視窗失敗: " + ex.Message);
                System.Diagnostics.Debug.WriteLine($"MTF Setting Form 錯誤: {ex.Message}");
            }
        }


        // 取得當前 MTF 資料
        // 在 Form1 的 GetCurrentMTFData 方法中加入更多除錯資訊
        // 修正 GetCurrentMTFData 方法，使用儲存的 Criterion 值
        private MTFSettingData GetCurrentMTFData()
        {
            System.Diagnostics.Debug.WriteLine("GetCurrentMTFData: 開始取得 MTF 資料");

            MTFSettingData data = new MTFSettingData();

            // 設定結果檔案目錄
            data.ResultFilesDirectory = Application.StartupPath;
            data.SettingFile = currentSettingsFilePath ?? "";
            data.MTFMethod = currentMTFMethod;

            // 確保 mtfRegions 已初始化
            if (mtfRegions == null || mtfRegions.Length < 5)
            {
                InitializeRegions();
            }

            // 設定當前 MTF 區域資料
            string[] regionNames = { "MTF0", "MTF1", "MTF2", "MTF3", "MTF4" };
            for (int i = 0; i < 5; i++)
            {
                MTFRegionData regionData = new MTFRegionData();
                regionData.Name = regionNames[i];

                // 計算中心點座標
                if (mtfRegions[i].Width > 0 && mtfRegions[i].Height > 0)
                {
                    regionData.X = mtfRegions[i].X + mtfRegions[i].Width / 2;
                    regionData.Y = mtfRegions[i].Y + mtfRegions[i].Height / 2;
                }
                else
                {
                    // 使用預設位置
                    int mtfSize = GetCurrentMTFAreaSize();
                    int halfSize = mtfSize / 2;
                    Point[] defaultCenters = {
                new Point(100 + halfSize, 100 + halfSize),
                new Point(100 + halfSize, 400 + halfSize),
                new Point(600 + halfSize, 400 + halfSize),
                new Point(600 + halfSize, 100 + halfSize),
                new Point(350 + halfSize, 250 + halfSize)
            };
                    regionData.X = defaultCenters[i].X;
                    regionData.Y = defaultCenters[i].Y;
                }

                // 取得當前值和最大值
                regionData.CurrentValue = (mtfCurrentLabels != null && mtfCurrentLabels[i] != null) ? mtfCurrentLabels[i].Text : "0.00";
                regionData.MaxValue = (mtfMaxLabels != null && mtfMaxLabels[i] != null) ? mtfMaxLabels[i].Text : "0.00";
                regionData.IsSelected = (mtfRegionVisible != null) ? mtfRegionVisible[i] : true;

                // 使用儲存的 Criterion 值
                regionData.Criterion = mtfCriterionValues[i];

                data.MTFRegions.Add(regionData);

                System.Diagnostics.Debug.WriteLine($"區域 {i}: {regionData.Name} at ({regionData.X}, {regionData.Y}), Criterion={regionData.Criterion}");
            }

            System.Diagnostics.Debug.WriteLine($"GetCurrentMTFData: 完成，區域總數: {data.MTFRegions.Count}");
            return data;
        }




        // 套用 MTF 設定
        // 修正 ApplyMTFSettings 方法，確保 Criterion 值正確更新
        // 修正 ApplyMTFSettings 方法，加入同步檢查
        private void ApplyMTFSettings(MTFSettingData data)
        {
            try
            {
                // 更新 MTF 方法
                currentMTFMethod = data.MTFMethod;

                // 更新區域位置
                for (int i = 0; i < Math.Min(5, data.MTFRegions.Count); i++)
                {
                    var regionData = data.MTFRegions[i];
                    int halfSize = GetCurrentMTFAreaSize() / 2;

                    mtfRegions[i] = new Rectangle(
                        regionData.X - halfSize,
                        regionData.Y - halfSize,
                        GetCurrentMTFAreaSize(),
                        GetCurrentMTFAreaSize()
                    );

                    // 更新 Criterion 值到成員變數
                    mtfCriterionValues[i] = regionData.Criterion;
                }

                // 更新顯示
                PictureBox pb = this.Controls.Find("pictureBox1", true).FirstOrDefault() as PictureBox;
                if (pb != null)
                    pb.Invalidate();

                // 只有在非同步狀態下才儲存到檔案
                if (!isFormSyncing)
                {
                    SaveMTFSettingsToFile(data);
                }

                System.Diagnostics.Debug.WriteLine("MTF 設定已更新");
            }
            catch (Exception ex)
            {
                MessageBox.Show("套用 MTF 設定失敗: " + ex.Message);
            }
        }
    

    // 取得當前 MTF Area 大小的輔助方法
    private int GetCurrentMTFAreaSize()
    {
        try
        {
            ComboBox cmbMTFArea = this.Controls.Find("cmbMTFArea", true).FirstOrDefault() as ComboBox;
            if (cmbMTFArea != null && int.TryParse(cmbMTFArea.Text, out int size))
            {
                return size;
            }
            return 128; // 預設值
        }
        catch (Exception)
        {
            return 128;
        }
    }


    // MTF50 主計算方法
    private double CalculateMTF50(Bitmap grayImage)
    {
        try
        {
            // 方法1: 使用 Slanted Edge 方法計算 MTF50
            return CalculateMTF50_SlantedEdge(grayImage);
        }
        catch (Exception)
        {
            // 如果失敗，使用備用方法
            return CalculateMTF50_Contrast(grayImage);
        }
    }

    // MTF50 計算 - Slanted Edge 方法 (ISO 12233 標準)
    private double CalculateMTF50_SlantedEdge(Bitmap grayImage)
    {
        try
        {
            // 1. 找到最強的邊緣
            Point edgeStart, edgeEnd;
            double edgeAngle;
            if (!DetectStrongestEdge(grayImage, out edgeStart, out edgeEnd, out edgeAngle))
            {
                return 0.0;
            }

            // 2. 建立垂直於邊緣的線性剖面
            double[] edgeProfile = ExtractEdgeProfile(grayImage, edgeStart, edgeEnd, edgeAngle);
            if (edgeProfile == null || edgeProfile.Length < 10)
            {
                return 0.0;
            }

            // 3. 計算邊緣擴散函數 (ESF)
            double[] esf = SmoothProfile(edgeProfile);

            // 4. 微分得到線擴散函數 (LSF)
            double[] lsf = DifferentiateESF(esf);

            // 5. 傅立葉變換得到 MTF
            double[] mtf = CalculateMTFFromLSF(lsf);

            // 6. 找到 MTF50 值
            return FindMTF50Value(mtf);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"MTF50 Slanted Edge 計算失敗: {ex.Message}");
            return 0.0;
        }
    }

    // 偵測最強邊緣
    private bool DetectStrongestEdge(Bitmap grayImage, out Point start, out Point end, out double angle)
    {
        start = Point.Empty;
        end = Point.Empty;
        angle = 0.0;

        try
        {
            double maxGradient = 0;
            Point bestStart = Point.Empty;
            Point bestEnd = Point.Empty;

            // Sobel 邊緣檢測
            for (int x = 1; x < grayImage.Width - 1; x++)
            {
                for (int y = 1; y < grayImage.Height - 1; y++)
                {
                    // 計算 Sobel 梯度
                    double gx = GetSobelGx(grayImage, x, y);
                    double gy = GetSobelGy(grayImage, x, y);
                    double gradient = Math.Sqrt(gx * gx + gy * gy);

                    if (gradient > maxGradient)
                    {
                        maxGradient = gradient;
                        bestStart = new Point(x, y);
                        angle = Math.Atan2(gy, gx);
                    }
                }
            }

            if (maxGradient > 10) // 閾值
            {
                start = bestStart;
                // 計算邊緣終點（簡化）
                int length = Math.Min(grayImage.Width, grayImage.Height) / 4;
                end = new Point(
                    start.X + (int)(length * Math.Cos(angle)),
                    start.Y + (int)(length * Math.Sin(angle))
                );
                return true;
            }

            return false;
        }
        catch (Exception)
        {
            return false;
        }
    }

    // Sobel X 方向梯度
    private double GetSobelGx(Bitmap image, int x, int y)
    {
        try
        {
            double gx = 0;
            int[,] kernelX = { { -1, 0, 1 }, { -2, 0, 2 }, { -1, 0, 1 } };

            for (int i = -1; i <= 1; i++)
            {
                for (int j = -1; j <= 1; j++)
                {
                    if (x + i >= 0 && x + i < image.Width && y + j >= 0 && y + j < image.Height)
                    {
                        Color pixel = image.GetPixel(x + i, y + j);
                        gx += pixel.R * kernelX[i + 1, j + 1];
                    }
                }
            }

            return gx;
        }
        catch (Exception)
        {
            return 0;
        }
    }

    // Sobel Y 方向梯度
    private double GetSobelGy(Bitmap image, int x, int y)
    {
        try
        {
            double gy = 0;
            int[,] kernelY = { { -1, -2, -1 }, { 0, 0, 0 }, { 1, 2, 1 } };

            for (int i = -1; i <= 1; i++)
            {
                for (int j = -1; j <= 1; j++)
                {
                    if (x + i >= 0 && x + i < image.Width && y + j >= 0 && y + j < image.Height)
                    {
                        Color pixel = image.GetPixel(x + i, y + j);
                        gy += pixel.R * kernelY[i + 1, j + 1];
                    }
                }
            }

            return gy;
        }
        catch (Exception)
        {
            return 0;
        }
    }

    // 擷取邊緣剖面
    private double[] ExtractEdgeProfile(Bitmap image, Point start, Point end, double angle)
    {
        try
        {
            List<double> profile = new List<double>();

            // 垂直於邊緣方向採樣
            double perpAngle = angle + Math.PI / 2;
            int profileLength = 50; // 剖面長度

            for (int i = -profileLength / 2; i < profileLength / 2; i++)
            {
                int x = start.X + (int)(i * Math.Cos(perpAngle));
                int y = start.Y + (int)(i * Math.Sin(perpAngle));

                if (x >= 0 && x < image.Width && y >= 0 && y < image.Height)
                {
                    Color pixel = image.GetPixel(x, y);
                    profile.Add(pixel.R);
                }
            }

            return profile.ToArray();
        }
        catch (Exception)
        {
            return null;
        }
    }

    // 平滑剖面（移動平均）
    private double[] SmoothProfile(double[] profile)
    {
        try
        {
            if (profile == null || profile.Length < 3) return profile;

            double[] smoothed = new double[profile.Length];

            for (int i = 0; i < profile.Length; i++)
            {
                double sum = 0;
                int count = 0;

                for (int j = Math.Max(0, i - 1); j <= Math.Min(profile.Length - 1, i + 1); j++)
                {
                    sum += profile[j];
                    count++;
                }

                smoothed[i] = sum / count;
            }

            return smoothed;
        }
        catch (Exception)
        {
            return profile;
        }
    }

    // 微分 ESF 得到 LSF
    private double[] DifferentiateESF(double[] esf)
    {
        try
        {
            if (esf == null || esf.Length < 2) return new double[0];

            double[] lsf = new double[esf.Length - 1];

            for (int i = 0; i < lsf.Length; i++)
            {
                lsf[i] = esf[i + 1] - esf[i];
            }

            return lsf;
        }
        catch (Exception)
        {
            return new double[0];
        }
    }

    // 從 LSF 計算 MTF（簡化的傅立葉變換）
    private double[] CalculateMTFFromLSF(double[] lsf)
    {
        try
        {
            if (lsf == null || lsf.Length == 0) return new double[0];

            int n = lsf.Length;
            double[] mtf = new double[n / 2];

            // 簡化的 DFT 計算
            for (int k = 0; k < mtf.Length; k++)
            {
                double real = 0, imag = 0;

                for (int t = 0; t < n; t++)
                {
                    double angle = -2.0 * Math.PI * k * t / n;
                    real += lsf[t] * Math.Cos(angle);
                    imag += lsf[t] * Math.Sin(angle);
                }

                mtf[k] = Math.Sqrt(real * real + imag * imag);
            }

            // 正規化 MTF
            if (mtf[0] > 0)
            {
                for (int i = 0; i < mtf.Length; i++)
                {
                    mtf[i] /= mtf[0];
                }
            }

            return mtf;
        }
        catch (Exception)
        {
            return new double[0];
        }
    }

    // 找到 MTF50 值
    private double FindMTF50Value(double[] mtf)
    {
        try
        {
            if (mtf == null || mtf.Length < 2) return 0.0;

            // 找到 MTF 值首次低於 0.5 的頻率
            for (int i = 1; i < mtf.Length; i++)
            {
                if (mtf[i] < 0.5)
                {
                    // 線性插值找到精確的 MTF50 頻率
                    double f1 = (double)(i - 1) / mtf.Length;
                    double f2 = (double)i / mtf.Length;
                    double m1 = mtf[i - 1];
                    double m2 = mtf[i];

                    // 插值計算 MTF = 0.5 時的頻率
                    double mtf50_freq = f1 + (0.5 - m1) * (f2 - f1) / (m2 - m1);

                    // 轉換為合理的數值範圍
                    return mtf50_freq * 100; // 乘以 100 得到百分比頻率
                }
            }

            return 0.0; // 沒有找到 MTF50
        }
        catch (Exception)
        {
            return 0.0;
        }
    }

    // 備用的對比度基礎 MTF 計算
    private double CalculateMTF50_Contrast(Bitmap grayImage)
    {
        try
        {
            // 使用對比度變化來估算 MTF
            double maxContrast = 0;
            double minContrast = 255;
            double totalContrast = 0;
            int contrastCount = 0;

            // 計算水平方向的對比度
            for (int y = 0; y < grayImage.Height; y++)
            {
                for (int x = 1; x < grayImage.Width; x++)
                {
                    Color current = grayImage.GetPixel(x, y);
                    Color previous = grayImage.GetPixel(x - 1, y);

                    double contrast = Math.Abs(current.R - previous.R);
                    maxContrast = Math.Max(maxContrast, contrast);
                    minContrast = Math.Min(minContrast, contrast);
                    totalContrast += contrast;
                    contrastCount++;
                }
            }

            // 計算垂直方向的對比度
            for (int x = 0; x < grayImage.Width; x++)
            {
                for (int y = 1; y < grayImage.Height; y++)
                {
                    Color current = grayImage.GetPixel(x, y);
                    Color previous = grayImage.GetPixel(x, y - 1);

                    double contrast = Math.Abs(current.R - previous.R);
                    maxContrast = Math.Max(maxContrast, contrast);
                    minContrast = Math.Min(minContrast, contrast);
                    totalContrast += contrast;
                    contrastCount++;
                }
            }

            // 計算對比度比率作為 MTF 估算
            if (maxContrast > 0 && contrastCount > 0)
            {
                double avgContrast = totalContrast / contrastCount;
                double contrastRatio = avgContrast / maxContrast;

                // 將對比度比率轉換為 MTF50 類似的數值
                return contrastRatio * 50; // 縮放到合理範圍
            }

            return 0.0;
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"備用 MTF 計算失敗: {ex.Message}");
            return 0.0;
        }
    }

    // 改進的邊緣檢測方法 (可選)
    private double CalculateMTF50_EdgeBased(Bitmap grayImage)
    {
        try
        {
            double totalEdgeStrength = 0;
            int edgeCount = 0;

            // 使用 Laplacian 運算子檢測邊緣
            for (int x = 1; x < grayImage.Width - 1; x++)
            {
                for (int y = 1; y < grayImage.Height - 1; y++)
                {
                    // Laplacian 核心
                    double laplacian = 0;
                    Color center = grayImage.GetPixel(x, y);
                    Color top = grayImage.GetPixel(x, y - 1);
                    Color bottom = grayImage.GetPixel(x, y + 1);
                    Color left = grayImage.GetPixel(x - 1, y);
                    Color right = grayImage.GetPixel(x + 1, y);

                    laplacian = -4 * center.R + top.R + bottom.R + left.R + right.R;
                    laplacian = Math.Abs(laplacian);

                    if (laplacian > 10) // 邊緣閾值
                    {
                        totalEdgeStrength += laplacian;
                        edgeCount++;
                    }
                }
            }

            if (edgeCount > 0)
            {
                double avgEdgeStrength = totalEdgeStrength / edgeCount;
                return Math.Min(avgEdgeStrength / 10, 50); // 限制在合理範圍
            }

            return 0.0;
        }
        catch (Exception)
        {
            return 0.0;
        }
    }

    // 新增：儲存 MTF 設定到參數設定檔
    // 修正 SaveMTFSettingsToFile 方法，儲存 Criterion 值
    private void SaveMTFSettingsToFile(MTFSettingData data)
    {
        try
        {
            if (string.IsNullOrEmpty(currentSettingsFilePath))
                return;

            // 讀取現有設定
            ImageAnalysisSettings settings = LoadSettingsFromFile(currentSettingsFilePath) ?? new ImageAnalysisSettings();

            // 更新 MTF 區域設定和 Criterion 值
            for (int i = 0; i < Math.Min(5, data.MTFRegions.Count); i++)
            {
                var regionData = data.MTFRegions[i];

                // 更新位置
                switch (i)
                {
                    case 0:
                        settings.M0_Center = new Point(regionData.X, regionData.Y);
                        settings.M0_Criterion = regionData.Criterion;
                        break;
                    case 1:
                        settings.M1_Center = new Point(regionData.X, regionData.Y);
                        settings.M1_Criterion = regionData.Criterion;
                        break;
                    case 2:
                        settings.M2_Center = new Point(regionData.X, regionData.Y);
                        settings.M2_Criterion = regionData.Criterion;
                        break;
                    case 3:
                        settings.M3_Center = new Point(regionData.X, regionData.Y);
                        settings.M3_Criterion = regionData.Criterion;
                        break;
                    case 4:
                        settings.MC_Center = new Point(regionData.X, regionData.Y);
                        settings.MC_Criterion = regionData.Criterion;
                        break;
                }

                // 同時更新成員變數
                mtfCriterionValues[i] = regionData.Criterion;
            }

            // 儲存到檔案
            SaveSettingsToFile(settings, currentSettingsFilePath);

            System.Diagnostics.Debug.WriteLine("MTF 設定（包含 Criterion 值）已儲存到設定檔");
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine("儲存 MTF 設定失敗: " + ex.Message);
        }
    }

    // 新增：儲存 Criterion 值到成員變數
    private void StoreCriterionValues(ImageAnalysisSettings settings)
    {
        mtfCriterionValues[0] = settings.M0_Criterion;
        mtfCriterionValues[1] = settings.M1_Criterion;
        mtfCriterionValues[2] = settings.M2_Criterion;
        mtfCriterionValues[3] = settings.M3_Criterion;
        mtfCriterionValues[4] = settings.MC_Criterion;

        System.Diagnostics.Debug.WriteLine($"載入 Criterion 值: M0={mtfCriterionValues[0]}, M1={mtfCriterionValues[1]}, M2={mtfCriterionValues[2]}, M3={mtfCriterionValues[3]}, MC={mtfCriterionValues[4]}");
    }

    // 新增：正確儲存包含 Criterion 值的設定
    private void SaveCurrentSettingsWithCriterion()
    {
        try
        {
            if (string.IsNullOrEmpty(currentSettingsFilePath))
            {
                // 如果沒有設定檔路徑，建立新檔
                string exePath = Application.StartupPath;
                string fileName = $"IASetsup_{DateTime.Now:yyyyMMdd}.set";
                currentSettingsFilePath = System.IO.Path.Combine(exePath, fileName);
            }

            // 取得當前完整設定（包含 UI 和 Criterion 值）
            ImageAnalysisSettings currentSettings = GetCurrentCompleteSettings();
            SaveSettingsToFile(currentSettings, currentSettingsFilePath);

            System.Diagnostics.Debug.WriteLine($"Exit 時儲存設定完成: {System.IO.Path.GetFileName(currentSettingsFilePath)}");
            System.Diagnostics.Debug.WriteLine($"儲存的 Criterion 值: M0={currentSettings.M0_Criterion}, M1={currentSettings.M1_Criterion}, M2={currentSettings.M2_Criterion}, M3={currentSettings.M3_Criterion}, MC={currentSettings.MC_Criterion}");
        }
        catch (Exception ex)
        {
            MessageBox.Show("Exit 儲存設定失敗: " + ex.Message);
        }
    }

    // 新增：取得當前完整設定（包含正確的 Criterion 值）
    private ImageAnalysisSettings GetCurrentCompleteSettings()
    {
        ImageAnalysisSettings settings = new ImageAnalysisSettings();

        try
        {
            // 取得 MTF Area
            ComboBox cmbMTFArea = this.Controls.Find("cmbMTFArea", true).FirstOrDefault() as ComboBox;
            if (cmbMTFArea != null && int.TryParse(cmbMTFArea.Text, out int mtfArea))
                settings.MTFArea = mtfArea;

            // 取得 Output Ratio
            ComboBox cmbOutputRatio = this.Controls.Find("cmbOutputRatio", true).FirstOrDefault() as ComboBox;
            if (cmbOutputRatio != null && int.TryParse(cmbOutputRatio.Text, out int outputRatio))
                settings.OutputRatio = outputRatio;

            // 取得 MTF 區域中心點
            if (mtfRegions != null && mtfRegions.Length >= 5)
            {
                settings.M0_Center = new Point(mtfRegions[0].X + mtfRegions[0].Width / 2, mtfRegions[0].Y + mtfRegions[0].Height / 2);
                settings.M1_Center = new Point(mtfRegions[1].X + mtfRegions[1].Width / 2, mtfRegions[1].Y + mtfRegions[1].Height / 2);
                settings.M2_Center = new Point(mtfRegions[2].X + mtfRegions[2].Width / 2, mtfRegions[2].Y + mtfRegions[2].Height / 2);
                settings.M3_Center = new Point(mtfRegions[3].X + mtfRegions[3].Width / 2, mtfRegions[3].Y + mtfRegions[3].Height / 2);
                settings.MC_Center = new Point(mtfRegions[4].X + mtfRegions[4].Width / 2, mtfRegions[4].Y + mtfRegions[4].Height / 2);
            }

            // 取得彩色區域中心點
            if (colorRegions != null && colorRegions.Length >= 3)
            {
                settings.C1_Center = new Point(colorRegions[0].X + colorRegions[0].Width / 2, colorRegions[0].Y + colorRegions[0].Height / 2);
                settings.C2_Center = new Point(colorRegions[1].X + colorRegions[1].Width / 2, colorRegions[1].Y + colorRegions[1].Height / 2);
                settings.C3_Center = new Point(colorRegions[2].X + colorRegions[2].Width / 2, colorRegions[2].Y + colorRegions[2].Height / 2);
            }

            // 使用當前的 Criterion 值（已經從 MTF Setting 更新的值）
            settings.M0_Criterion = mtfCriterionValues[0];
            settings.M1_Criterion = mtfCriterionValues[1];
            settings.M2_Criterion = mtfCriterionValues[2];
            settings.M3_Criterion = mtfCriterionValues[3];
            settings.MC_Criterion = mtfCriterionValues[4];

            return settings;
        }
        catch (Exception ex)
        {
            MessageBox.Show("取得當前設定失敗: " + ex.Message);
            return new ImageAnalysisSettings();
        }
    }

    // 新增：除錯用 - 顯示當前 Criterion 值
    private void DebugShowCurrentCriterion()
    {
        System.Diagnostics.Debug.WriteLine($"當前 Criterion 值: M0={mtfCriterionValues[0]}, M1={mtfCriterionValues[1]}, M2={mtfCriterionValues[2]}, M3={mtfCriterionValues[3]}, MC={mtfCriterionValues[4]}");
    }

    // MTF Setting Form 資料更新事件處理
    private void OnMTFSettingFormDataUpdated(MTFSettingData updatedData)
    {
        if (isFormSyncing) return;

        isFormSyncing = true;
        try
        {
            System.Diagnostics.Debug.WriteLine("收到 MTF Setting Form 資料更新");
            ApplyMTFSettings(updatedData);

            // 觸發 Form1 資料變更事件（通知其他可能的監聽者）
            MTFDataChanged?.Invoke(updatedData);
        }
        finally
        {
            isFormSyncing = false;
        }
    }

    // MTF Setting Form 區域選擇事件處理
    private void OnMTFSettingFormRegionSelected(int regionIndex)
    {
        if (regionIndex >= 0 && regionIndex < 5)
        {
            // 可以在 Form1 上高亮顯示選中的區域
            HighlightMTFRegion(regionIndex);
        }
    }

    // MTF Setting Form 關閉事件處理
    private void OnMTFSettingFormClosed(object sender, FormClosedEventArgs e)
    {
        if (mtfSettingForm != null)
        {
            // 移除事件訂閱
            mtfSettingForm.MTFDataUpdated -= OnMTFSettingFormDataUpdated;
            mtfSettingForm.MTFRegionSelected -= OnMTFSettingFormRegionSelected;
            mtfSettingForm.FormClosed -= OnMTFSettingFormClosed;
            this.MTFDataChanged -= mtfSettingForm.OnForm1DataChanged;
            this.MTFRegionMoved -= mtfSettingForm.OnForm1RegionMoved;

            mtfSettingForm = null;
        }

        System.Diagnostics.Debug.WriteLine("MTF Setting Form 已關閉");
    }

    // 新增：高亮顯示選中的 MTF 區域
    private void HighlightMTFRegion(int regionIndex)
    {
        // 這裡可以實作高亮效果，例如改變邊框顏色
        PictureBox pb = this.Controls.Find("pictureBox1", true).FirstOrDefault() as PictureBox;
        if (pb != null)
        {
            pb.Invalidate(); // 重繪以顯示高亮效果
        }
    }

    






    }

}