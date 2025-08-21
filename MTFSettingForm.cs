using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace ImageAnalysis
{
    public partial class MTFSettingForm : Form
    {

        private MTFSettingData mtfData;
        private ListBox lstMTFRegions;
        private TextBox txtResultDirectory;
        private TextBox txtSettingFile;
        private RadioButton rbPixelDifference;
        private RadioButton rbMTF50;
        private NumericUpDown nudPointX;
        private NumericUpDown nudPointY;
        private NumericUpDown nudCriterion; // 改為 NumericUpDown
        private int selectedRegionIndex = 0;
        private bool isUpdating = false; // 防止遞迴更新

        public event Action<MTFSettingData> MTFDataUpdated;

        // 新增：同步相關事件
        public event Action<int> MTFRegionSelected; // regionIndex



        public MTFSettingForm()
        {
            Initialize_Component();
        }


        private void Initialize_Component()
        {
            this.Text = "MTF Setting";
            this.Size = new Size(600, 400);
            this.StartPosition = FormStartPosition.Manual;
            this.FormBorderStyle = FormBorderStyle.Sizable; // 允許調整大小
            this.MaximizeBox = true;
            this.MinimizeBox = true;
            this.ShowInTaskbar = true; // 在工作列顯示

            CreateControls();

            // 設定初始位置（避免與 Form1 重疊）
            this.Location = new Point(100, 100);
        }

        private void CreateControls()
        {
            // Result Files Directory
            Label lblResultDir = new Label();
            lblResultDir.Text = "Result Files Directory";
            lblResultDir.Location = new Point(20, 20);
            lblResultDir.Size = new Size(150, 20);
            this.Controls.Add(lblResultDir);

            txtResultDirectory = new TextBox();
            txtResultDirectory.Location = new Point(20, 45);
            txtResultDirectory.Size = new Size(400, 25);
            this.Controls.Add(txtResultDirectory);

            Button btnBrowseResult = new Button();
            btnBrowseResult.Text = "...";
            btnBrowseResult.Location = new Point(430, 45);
            btnBrowseResult.Size = new Size(30, 25);
            btnBrowseResult.Click += BtnBrowseResult_Click;
            this.Controls.Add(btnBrowseResult);

            // Setting File
            Label lblSettingFile = new Label();
            lblSettingFile.Text = "Setting File";
            lblSettingFile.Location = new Point(470, 20);
            lblSettingFile.Size = new Size(80, 20);
            this.Controls.Add(lblSettingFile);

            txtSettingFile = new TextBox();
            txtSettingFile.Location = new Point(470, 45);
            txtSettingFile.Size = new Size(100, 25);
            txtSettingFile.ReadOnly = true;
            this.Controls.Add(txtSettingFile);

            Button btnSave = new Button();
            btnSave.Text = "Save";
            btnSave.Location = new Point(470, 80);
            btnSave.Size = new Size(60, 30);
            btnSave.Click += BtnSave_Click;
            this.Controls.Add(btnSave);

            Button btnExit = new Button();
            btnExit.Text = "Exit";
            btnExit.Location = new Point(470, 120);
            btnExit.Size = new Size(60, 30);
            btnExit.Click += BtnExit_Click;
            this.Controls.Add(btnExit);

            // MTF Method
            Label lblMTFMethod = new Label();
            lblMTFMethod.Text = "MTF Method";
            lblMTFMethod.Location = new Point(20, 85);
            lblMTFMethod.Size = new Size(80, 20);
            this.Controls.Add(lblMTFMethod);

            rbPixelDifference = new RadioButton();
            rbPixelDifference.Text = "像素灰度差異";
            rbPixelDifference.Location = new Point(20, 110);
            rbPixelDifference.Size = new Size(120, 20);
            rbPixelDifference.Checked = true;
            rbPixelDifference.CheckedChanged += RbMTFMethod_CheckedChanged;
            this.Controls.Add(rbPixelDifference);

            rbMTF50 = new RadioButton();
            rbMTF50.Text = "MTF50";
            rbMTF50.Location = new Point(150, 110);
            rbMTF50.Size = new Size(80, 20);
            rbMTF50.CheckedChanged += RbMTFMethod_CheckedChanged;
            this.Controls.Add(rbMTF50);

            // MTF Setting (List)
            Label lblMTFSetting = new Label();
            lblMTFSetting.Text = "MTF Setting";
            lblMTFSetting.Location = new Point(20, 140);
            lblMTFSetting.Size = new Size(80, 20);
            this.Controls.Add(lblMTFSetting);

            lstMTFRegions = new ListBox();
            lstMTFRegions.Location = new Point(20, 165);
            lstMTFRegions.Size = new Size(200, 120);
            lstMTFRegions.SelectedIndexChanged += LstMTFRegions_SelectedIndexChanged;
            this.Controls.Add(lstMTFRegions);

            // Point X
            Label lblPointX = new Label();
            lblPointX.Text = "Point X";
            lblPointX.Location = new Point(240, 140);
            lblPointX.Size = new Size(50, 20);
            this.Controls.Add(lblPointX);

            nudPointX = new NumericUpDown();
            nudPointX.Location = new Point(240, 165);
            nudPointX.Size = new Size(80, 25);
            nudPointX.Minimum = 0;
            nudPointX.Maximum = 9999;
            nudPointX.ValueChanged += NudPoint_ValueChanged;
            this.Controls.Add(nudPointX);

            Button btnPointXMinus = new Button();
            btnPointXMinus.Text = "<";
            btnPointXMinus.Location = new Point(330, 165);
            btnPointXMinus.Size = new Size(25, 25);
            btnPointXMinus.Click += (s, e) => { nudPointX.Value = Math.Max(nudPointX.Minimum, nudPointX.Value - 1); };
            this.Controls.Add(btnPointXMinus);

            Button btnPointXPlus = new Button();
            btnPointXPlus.Text = ">";
            btnPointXPlus.Location = new Point(360, 165);
            btnPointXPlus.Size = new Size(25, 25);
            btnPointXPlus.Click += (s, e) => { nudPointX.Value = Math.Min(nudPointX.Maximum, nudPointX.Value + 1); };
            this.Controls.Add(btnPointXPlus);

            // Point Y
            Label lblPointY = new Label();
            lblPointY.Text = "Point Y";
            lblPointY.Location = new Point(240, 200);
            lblPointY.Size = new Size(50, 20);
            this.Controls.Add(lblPointY);

            nudPointY = new NumericUpDown();
            nudPointY.Location = new Point(240, 225);
            nudPointY.Size = new Size(80, 25);
            nudPointY.Minimum = 0;
            nudPointY.Maximum = 9999;
            nudPointY.ValueChanged += NudPoint_ValueChanged;
            this.Controls.Add(nudPointY);

            Button btnPointYMinus = new Button();
            btnPointYMinus.Text = "<";
            btnPointYMinus.Location = new Point(330, 225);
            btnPointYMinus.Size = new Size(25, 25);
            btnPointYMinus.Click += (s, e) => { nudPointY.Value = Math.Max(nudPointY.Minimum, nudPointY.Value - 1); };
            this.Controls.Add(btnPointYMinus);

            Button btnPointYPlus = new Button();
            btnPointYPlus.Text = ">";
            btnPointYPlus.Location = new Point(360, 225);
            btnPointYPlus.Size = new Size(25, 25);
            btnPointYPlus.Click += (s, e) => { nudPointY.Value = Math.Min(nudPointY.Maximum, nudPointY.Value + 1); };
            this.Controls.Add(btnPointYPlus);

            // Criterion
            Label lblCriterion = new Label();
            lblCriterion.Text = "Criterion";
            lblCriterion.Location = new Point(240, 260);
            lblCriterion.Size = new Size(60, 20);
            this.Controls.Add(lblCriterion);

            // 改為 NumericUpDown
            nudCriterion = new NumericUpDown();
            nudCriterion.Location = new Point(240, 285);
            nudCriterion.Size = new Size(80, 25);
            nudCriterion.Minimum = 1;
            nudCriterion.Maximum = 100;
            nudCriterion.Value = 50;
            nudCriterion.ValueChanged += NudCriterion_ValueChanged;
            this.Controls.Add(nudCriterion);

            Button btnCriterionMinus = new Button();
            btnCriterionMinus.Text = "<";
            btnCriterionMinus.Location = new Point(330, 285);
            btnCriterionMinus.Size = new Size(25, 25);
            btnCriterionMinus.Click += (s, e) => { nudCriterion.Value = Math.Max(nudCriterion.Minimum, nudCriterion.Value - 1); };
            this.Controls.Add(btnCriterionMinus);

            Button btnCriterionPlus = new Button();
            btnCriterionPlus.Text = ">";
            btnCriterionPlus.Location = new Point(360, 285);
            btnCriterionPlus.Size = new Size(25, 25);
            btnCriterionPlus.Click += (s, e) => { nudCriterion.Value = Math.Min(nudCriterion.Maximum, nudCriterion.Value + 1); };
            this.Controls.Add(btnCriterionPlus);
        }


        // 修正 MTFSettingForm 的 SetMTFData 方法，加入除錯資訊
        public void SetMTFData(MTFSettingData data)
        {
            if (data == null)
            {
                MessageBox.Show("MTF 資料為空");
                return;
            }

            mtfData = data;

            // 先不設定 isUpdating，讓 UpdateMTFRegionsList 可以執行
            System.Diagnostics.Debug.WriteLine($"設定 MTF 資料，區域數量: {data.MTFRegions.Count}");

            // 更新 UI
            txtResultDirectory.Text = data.ResultFilesDirectory ?? "";
            txtSettingFile.Text = string.IsNullOrEmpty(data.SettingFile) ? "" : System.IO.Path.GetFileName(data.SettingFile);

            rbPixelDifference.Checked = (data.MTFMethod == MTFMethod.PixelDifference);
            rbMTF50.Checked = (data.MTFMethod == MTFMethod.MTF50);

            // 確保有 MTF 區域資料
            if (data.MTFRegions.Count == 0)
            {
                System.Diagnostics.Debug.WriteLine("沒有 MTF 區域資料，建立預設資料");
                CreateDefaultMTFRegions();
            }

            // 直接更新列表，不使用 isUpdating 標記
            UpdateMTFRegionsListDirect();

            // 設定選擇
            if (lstMTFRegions.Items.Count > 0)
            {
                lstMTFRegions.SelectedIndex = 0;
                System.Diagnostics.Debug.WriteLine($"選擇第一個項目: {lstMTFRegions.Items[0]}");

                // 手動觸發選擇事件以更新右側數值
                LstMTFRegions_SelectedIndexChanged(lstMTFRegions, EventArgs.Empty);
            }
            else
            {
                System.Diagnostics.Debug.WriteLine("MTF 區域列表仍然為空");
            }
        }


        // 修正 UpdateMTFRegionsList 方法，加入除錯資訊
        private void UpdateMTFRegionsList()
        {
            if (mtfData == null)
            {
                System.Diagnostics.Debug.WriteLine("UpdateMTFRegionsList: mtfData 為空");
                return;
            }

            if (isUpdating)
            {
                System.Diagnostics.Debug.WriteLine("UpdateMTFRegionsList: 正在更新中，跳過");
                return;
            }

            isUpdating = true;
            try
            {
                int currentSelection = lstMTFRegions.SelectedIndex;

                System.Diagnostics.Debug.WriteLine($"UpdateMTFRegionsList: 更新列表，當前選擇: {currentSelection}");

                UpdateMTFRegionsListDirect();

                // 恢復選擇
                if (currentSelection >= 0 && currentSelection < lstMTFRegions.Items.Count)
                {
                    lstMTFRegions.SelectedIndex = currentSelection;
                }
                else if (lstMTFRegions.Items.Count > 0)
                {
                    lstMTFRegions.SelectedIndex = 0;
                }
            }
            finally
            {
                isUpdating = false;
            }
        }


        // 修正 LstMTFRegions_SelectedIndexChanged 方法，加入除錯資訊
        // 修正選擇變更事件，加入同步通知
        private void LstMTFRegions_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (isUpdating) return;

            selectedRegionIndex = lstMTFRegions.SelectedIndex;

            System.Diagnostics.Debug.WriteLine($"MTF Setting Form: 選擇索引 {selectedRegionIndex}");

            // 通知 Form1 選中的區域
            MTFRegionSelected?.Invoke(selectedRegionIndex);

            if (selectedRegionIndex >= 0 && selectedRegionIndex < mtfData.MTFRegions.Count)
            {
                var region = mtfData.MTFRegions[selectedRegionIndex];

                isUpdating = true;
                try
                {
                    nudPointX.Value = Math.Max(nudPointX.Minimum, Math.Min(nudPointX.Maximum, region.X));
                    nudPointY.Value = Math.Max(nudPointY.Minimum, Math.Min(nudPointY.Maximum, region.Y));
                    nudCriterion.Value = Math.Max(nudCriterion.Minimum, Math.Min(nudCriterion.Maximum, region.Criterion));
                }
                finally
                {
                    isUpdating = false;
                }
            }
        }




        private void NudPoint_ValueChanged(object sender, EventArgs e)
        {
            if (isUpdating) return;
            UpdateCurrentRegionData();
        }

        private void NudCriterion_ValueChanged(object sender, EventArgs e)
        {
            if (isUpdating) return;
            UpdateCurrentRegionData();
        }

        // 修正數值變更事件，加入即時同步
        private void UpdateCurrentRegionData()
        {
            if (selectedRegionIndex >= 0 && selectedRegionIndex < mtfData.MTFRegions.Count && !isUpdating)
            {
                // 更新資料
                mtfData.MTFRegions[selectedRegionIndex].X = (int)nudPointX.Value;
                mtfData.MTFRegions[selectedRegionIndex].Y = (int)nudPointY.Value;
                mtfData.MTFRegions[selectedRegionIndex].Criterion = (int)nudCriterion.Value;

                // 更新顯示列表
                UpdateMTFRegionsList();

                // 即時同步到 Form1
                MTFDataUpdated?.Invoke(mtfData);

                System.Diagnostics.Debug.WriteLine($"MTF Setting Form: 即時更新區域 {selectedRegionIndex}");
            }
        }


        private void RbMTFMethod_CheckedChanged(object sender, EventArgs e)
        {
            if (isUpdating) return;

            if (rbPixelDifference.Checked)
            {
                mtfData.MTFMethod = MTFMethod.PixelDifference;
            }
            else if (rbMTF50.Checked)
            {
                mtfData.MTFMethod = MTFMethod.MTF50;
            }

            // 即時更新 Form1
            MTFDataUpdated?.Invoke(mtfData);
        }


        private void BtnBrowseResult_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            fbd.SelectedPath = txtResultDirectory.Text;

            if (fbd.ShowDialog() == DialogResult.OK)
            {
                txtResultDirectory.Text = fbd.SelectedPath;
                mtfData.ResultFilesDirectory = fbd.SelectedPath;
            }
        }

        // 修正 MTFSettingForm 的 BtnSave_Click，確保主 Form 收到更新
        // 修正 BtnSave_Click，移除自動關閉
        private void BtnSave_Click(object sender, EventArgs e)
        {
            try
            {
                // 觸發事件通知主表單儲存設定
                MTFDataUpdated?.Invoke(mtfData);

                MessageBox.Show("設定已儲存到參數設定檔");

                System.Diagnostics.Debug.WriteLine("MTF Setting Form: Save 按鈕被按下");
            }
            catch (Exception ex)
            {
                MessageBox.Show("儲存設定失敗: " + ex.Message);
            }
        }


        // 修正 BtnExit_Click，正確關閉視窗
        private void BtnExit_Click(object sender, EventArgs e)
        {
            this.Close(); // 關閉視窗而不是隱藏
        }

        // 新增：建立預設 MTF 區域資料
        // 修正 CreateDefaultMTFRegions 方法，加入更多除錯資訊
        private void CreateDefaultMTFRegions()
        {
            if (mtfData == null)
            {
                System.Diagnostics.Debug.WriteLine("CreateDefaultMTFRegions: mtfData 為空，建立新的");
                mtfData = new MTFSettingData();
            }

            string[] regionNames = { "MTF0", "MTF1", "MTF2", "MTF3", "MTF4" };
            Point[] defaultCenters = {
        new Point(164, 164),   // M0 center
        new Point(164, 464),   // M1 center  
        new Point(664, 464),   // M2 center
        new Point(664, 164),   // M3 center
        new Point(414, 314)    // MC center
    };

            mtfData.MTFRegions.Clear();

            for (int i = 0; i < 5; i++)
            {
                MTFRegionData regionData = new MTFRegionData();
                regionData.Name = regionNames[i];
                regionData.X = defaultCenters[i].X;
                regionData.Y = defaultCenters[i].Y;
                regionData.CurrentValue = "0.00";
                regionData.MaxValue = "0.00";
                regionData.IsSelected = true;
                regionData.Criterion = 50;

                mtfData.MTFRegions.Add(regionData);

                System.Diagnostics.Debug.WriteLine($"建立預設區域: {regionData.Name} at ({regionData.X}, {regionData.Y})");
            }

            System.Diagnostics.Debug.WriteLine($"預設區域建立完成，總數: {mtfData.MTFRegions.Count}");
        }


        // 新增：直接更新列表的方法（不受 isUpdating 影響）
        private void UpdateMTFRegionsListDirect()
        {
            if (mtfData == null) return;

            System.Diagnostics.Debug.WriteLine($"直接更新 MTF 區域列表，資料數量: {mtfData.MTFRegions.Count}");

            lstMTFRegions.Items.Clear();

            foreach (var region in mtfData.MTFRegions)
            {
                // 格式：MTF0--50,164,164
                string displayText = $"{region.Name}--{region.Criterion},{region.X},{region.Y}";
                lstMTFRegions.Items.Add(displayText);

                System.Diagnostics.Debug.WriteLine($"加入項目: {displayText}");
            }

            System.Diagnostics.Debug.WriteLine($"列表項目總數: {lstMTFRegions.Items.Count}");
        }

        // 新增：接收 Form1 資料變更的方法
        public void OnForm1DataChanged(MTFSettingData updatedData)
        {
            if (isUpdating || this.IsDisposed) return;

            System.Diagnostics.Debug.WriteLine("MTF Setting Form: 收到 Form1 資料變更");

            isUpdating = true;
            try
            {
                mtfData = updatedData;
                UpdateMTFRegionsListDirect();

                // 如果有選中的項目，更新右側數值
                if (selectedRegionIndex >= 0 && selectedRegionIndex < mtfData.MTFRegions.Count)
                {
                    var region = mtfData.MTFRegions[selectedRegionIndex];
                    nudPointX.Value = Math.Max(nudPointX.Minimum, Math.Min(nudPointX.Maximum, region.X));
                    nudPointY.Value = Math.Max(nudPointY.Minimum, Math.Min(nudPointY.Maximum, region.Y));
                    nudCriterion.Value = Math.Max(nudCriterion.Minimum, Math.Min(nudCriterion.Maximum, region.Criterion));
                }
            }
            finally
            {
                isUpdating = false;
            }
        }

        // 新增：接收 Form1 區域移動的方法
        public void OnForm1RegionMoved(int regionIndex, Point newCenter)
        {
            if (isUpdating || this.IsDisposed) return;

            if (regionIndex >= 0 && regionIndex < mtfData.MTFRegions.Count)
            {
                System.Diagnostics.Debug.WriteLine($"MTF Setting Form: 收到區域 {regionIndex} 移動到 ({newCenter.X}, {newCenter.Y})");

                isUpdating = true;
                try
                {
                    // 更新資料
                    mtfData.MTFRegions[regionIndex].X = newCenter.X;
                    mtfData.MTFRegions[regionIndex].Y = newCenter.Y;

                    // 更新列表顯示
                    UpdateMTFRegionsListDirect();

                    // 如果移動的是當前選中的區域，更新右側數值
                    if (regionIndex == selectedRegionIndex)
                    {
                        nudPointX.Value = newCenter.X;
                        nudPointY.Value = newCenter.Y;
                    }
                }
                finally
                {
                    isUpdating = false;
                }
            }
        }

        // 覆寫 FormClosing 事件，確保正確清理
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            System.Diagnostics.Debug.WriteLine("MTF Setting Form: 正在關閉");
            base.OnFormClosing(e);
        }

        // 新增：帶到前面的方法
        public void BringToFrontAndFocus()
        {
            this.BringToFront();
            this.Focus();
            if (this.WindowState == FormWindowState.Minimized)
            {
                this.WindowState = FormWindowState.Normal;
            }
        }



    }
}
