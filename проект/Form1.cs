using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using ClosedXML.Excel;

namespace BytovkaBuilder
{
    public partial class Form1 : Form
    {
        // Внешняя обшивка
        private RadioButton rbIronLoadBearing, rbIronWall, rbIronLoadWall;
        private RadioButton rbImitationTimberExt, rbLiningExt, rbSiding;

        // Внутренняя отделка стен
        private RadioButton rbOSBWall, rbMDFSheetWall, rbPVCWall, rbMDFPanelWall, rbLiningWall, rbImitationTimberWall;

        // Пол
        private RadioButton rbPlywoodFloor, rbOSBFloor, rbRoughBoardFloor;

        // Дверь и окна
        private CheckBox chkDoor;
        private CheckBox chkWindow;
        private NumericUpDown numWidth, numHeight;

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        // Поле для пожеланий
        private TextBox txtWishes;

        // Превью и кнопки
        private PictureBox pictureBox;
        private Button btnShow, btnExportExcel;

        public Form1()
        {
            InitializeComponent();
            SetupForm();
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(240)))), ((int)(((byte)(242)))), ((int)(((byte)(245)))));
            this.ClientSize = new System.Drawing.Size(1100, 750);
            this.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Конструктор бытовки";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);

        }

        private void SetupForm()
        {
            int leftMargin = 20;
            int topMargin = 20;
            int groupWidth = 280;

            int currentY = topMargin;

            // ====== Внешняя обшивка ======
            GroupBox gbExternal = new GroupBox()
            {
                Text = "Внешняя обшивка",
                Location = new Point(leftMargin, currentY),
                Size = new Size(groupWidth, 180),
                BackColor = Color.White
            };

            rbIronLoadBearing = new RadioButton() { Text = "Железо (несущий)", Location = new Point(15, 30), AutoSize = true };
            rbIronWall = new RadioButton() { Text = "Железо (стеновой)", Location = new Point(15, 55), AutoSize = true };
            rbIronLoadWall = new RadioButton() { Text = "Железо (несущий-стеновой)", Location = new Point(15, 80), AutoSize = true };
            rbImitationTimberExt = new RadioButton() { Text = "Имитация бруса", Location = new Point(15, 105), AutoSize = true };
            rbLiningExt = new RadioButton() { Text = "Вагонка", Location = new Point(15, 130), AutoSize = true };
            rbSiding = new RadioButton() { Text = "Сайдинг", Location = new Point(15, 155), AutoSize = true };

            rbIronLoadBearing.Checked = true;

            gbExternal.Controls.AddRange(new Control[] { rbIronLoadBearing, rbIronWall, rbIronLoadWall, rbImitationTimberExt, rbLiningExt, rbSiding });
            this.Controls.Add(gbExternal);

            currentY += 195;

            // ====== Внутренняя отделка стен ======
            GroupBox gbInternal = new GroupBox()
            {
                Text = "Внутренняя отделка стен",
                Location = new Point(leftMargin, currentY),
                Size = new Size(groupWidth, 180),
                BackColor = Color.White
            };

            rbOSBWall = new RadioButton() { Text = "ОСБ", Location = new Point(15, 30), AutoSize = true };
            rbMDFSheetWall = new RadioButton() { Text = "МДФ листы", Location = new Point(15, 55), AutoSize = true };
            rbPVCWall = new RadioButton() { Text = "ПВХ", Location = new Point(15, 80), AutoSize = true };
            rbMDFPanelWall = new RadioButton() { Text = "МДФ панели", Location = new Point(15, 105), AutoSize = true };
            rbLiningWall = new RadioButton() { Text = "Вагонка", Location = new Point(15, 130), AutoSize = true };
            rbImitationTimberWall = new RadioButton() { Text = "Имитация бруса", Location = new Point(15, 155), AutoSize = true };

            rbOSBWall.Checked = true;

            gbInternal.Controls.AddRange(new Control[] { rbOSBWall, rbMDFSheetWall, rbPVCWall, rbMDFPanelWall, rbLiningWall, rbImitationTimberWall });
            this.Controls.Add(gbInternal);

            currentY += 195;

            // ====== Пол ======
            GroupBox gbFloor = new GroupBox()
            {
                Text = "Пол",
                Location = new Point(leftMargin, currentY),
                Size = new Size(groupWidth, 120),
                BackColor = Color.White
            };

            rbPlywoodFloor = new RadioButton() { Text = "Фанера", Location = new Point(15, 30), AutoSize = true };
            rbOSBFloor = new RadioButton() { Text = "ОСБ", Location = new Point(15, 55), AutoSize = true };
            rbRoughBoardFloor = new RadioButton() { Text = "Черновая доска", Location = new Point(15, 80), AutoSize = true };

            rbPlywoodFloor.Checked = true;

            gbFloor.Controls.AddRange(new Control[] { rbPlywoodFloor, rbOSBFloor, rbRoughBoardFloor });
            this.Controls.Add(gbFloor);

            currentY += 135;

            // ====== Дверь ======
            GroupBox gbDoor = new GroupBox()
            {
                Text = "Дверь",
                Location = new Point(leftMargin, currentY),
                Size = new Size(groupWidth, 60),
                BackColor = Color.White
            };

            chkDoor = new CheckBox() { Text = "Стандартная ГОСТ дверь", Location = new Point(15, 28), AutoSize = true, Checked = true };
            gbDoor.Controls.Add(chkDoor);
            this.Controls.Add(gbDoor);

            currentY += 75;

            // ====== Окна ======
            GroupBox gbWindow = new GroupBox()
            {
                Text = "Окна",
                Location = new Point(leftMargin, currentY),
                Size = new Size(groupWidth, 100),
                BackColor = Color.White
            };

            chkWindow = new CheckBox()
            {
                Text = "Наличие окна",
                Location = new Point(15, 28),
                AutoSize = true,
                Checked = true
            };

            Label lblSize = new Label() { Text = "Размер (см):", Location = new Point(15, 55), AutoSize = true };

            numWidth = new NumericUpDown()
            {
                Location = new Point(95, 52),
                Width = 50,
                Minimum = 50,
                Maximum = 200,
                Value = 80
            };

            Label lblX = new Label() { Text = "x", Location = new Point(150, 55), AutoSize = true };

            numHeight = new NumericUpDown()
            {
                Location = new Point(165, 52),
                Width = 50,
                Minimum = 50,
                Maximum = 200,
                Value = 60
            };

            Label lblCm = new Label() { Text = "см", Location = new Point(220, 55), AutoSize = true };

            chkWindow.CheckedChanged += (s, e) =>
            {
                numWidth.Enabled = chkWindow.Checked;
                numHeight.Enabled = chkWindow.Checked;
            };

            gbWindow.Controls.AddRange(new Control[] { chkWindow, lblSize, numWidth, lblX, numHeight, lblCm });
            this.Controls.Add(gbWindow);

            currentY += 115;

            // ====== Дополнительные пожелания ======
            GroupBox gbWishes = new GroupBox()
            {
                Text = "Дополнительные пожелания",
                Location = new Point(leftMargin, currentY),
                Size = new Size(groupWidth, 120),
                BackColor = Color.White
            };

            txtWishes = new TextBox()
            {
                Location = new Point(10, 30),
                Size = new Size(260, 75),
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                Text = "Напишите здесь ваши пожелания..."
            };

            txtWishes.Enter += (s, e) =>
            {
                if (txtWishes.Text == "Напишите здесь ваши пожелания...")
                {
                    txtWishes.Text = "";
                }
            };
            txtWishes.Leave += (s, e) =>
            {
                if (string.IsNullOrWhiteSpace(txtWishes.Text))
                {
                    txtWishes.Text = "Напишите здесь ваши пожелания...";
                }
            };

            gbWishes.Controls.Add(txtWishes);
            this.Controls.Add(gbWishes);

            // ====== ПРАВАЯ ПАНЕЛЬ ======
            int rightX = leftMargin + groupWidth + 30;

            // Кнопка "Показать бытовку"
            btnShow = new Button()
            {
                Text = "Показать бытовку",
                Location = new Point(rightX, topMargin),
                Size = new Size(200, 40),
                BackColor = Color.FromArgb(52, 152, 219),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 10, FontStyle.Bold)
            };
            btnShow.Click += BtnShow_Click;
            this.Controls.Add(btnShow);

            // Кнопка "Экспорт в Excel"
            btnExportExcel = new Button()
            {
                Text = "Сформировать ТЗ в Excel",
                Location = new Point(rightX + 215, topMargin),
                Size = new Size(200, 40),
                BackColor = Color.FromArgb(46, 204, 113),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 10, FontStyle.Bold)
            };
            btnExportExcel.Click += BtnExportExcel_Click;
            this.Controls.Add(btnExportExcel);

            // PictureBox для фото
            pictureBox = new PictureBox()
            {
                Location = new Point(rightX, topMargin + 60),
                Size = new Size(415, 400),
                BorderStyle = BorderStyle.FixedSingle,
                SizeMode = PictureBoxSizeMode.Zoom,
                BackColor = Color.White
            };
            this.Controls.Add(pictureBox);

            // Информационная панель с путём к папке
            string imagesPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Images");

            Label lblInfo = new Label()
            {
                Text = $"Все изображения найдены в интернете,\n по большей части не похожи на те что в жизни,\n уточняйте у владельца компании бытовок на счет внешнего вида.\n Контактная имнформация: 8-999-999-99-99 Артем2033",
                Location = new Point(rightX, topMargin + 480),
                Size = new Size(415, 60),
                ForeColor = Color.DarkBlue,
                Font = new Font("Segoe UI", 8),
                BackColor = Color.FromArgb(255, 255, 200),
                BorderStyle = BorderStyle.FixedSingle
            };
            this.Controls.Add(lblInfo);

            // Создаём папку Images, если её нет
            if (!Directory.Exists(imagesPath))
            {
                Directory.CreateDirectory(imagesPath);
            }
        }

        // Функция возвращает имя файла в зависимости от выбранного радио-баттона
        private string GetImageFileName()
        {
            if (rbIronLoadBearing.Checked) return "iron_loadbearing.jpg";
            if (rbIronWall.Checked) return "iron_wall.jpg";
            if (rbIronLoadWall.Checked) return "iron_loadwall.jpg";
            if (rbImitationTimberExt.Checked) return "imitation_timber.jpg";
            if (rbLiningExt.Checked) return "lining.jpg";
            if (rbSiding.Checked) return "siding.jpg";
            return "iron_loadbearing.jpg";
        }

        private string GetSelectedExteriorMaterialName()
        {
            if (rbIronLoadBearing.Checked) return "Железо (несущий)";
            if (rbIronWall.Checked) return "Железо (стеновой)";
            if (rbIronLoadWall.Checked) return "Железо (несущий-стеновой)";
            if (rbImitationTimberExt.Checked) return "Имитация бруса";
            if (rbLiningExt.Checked) return "Вагонка";
            if (rbSiding.Checked) return "Сайдинг";
            return "Не выбрано";
        }

        private string GetSelectedInterior()
        {
            if (rbOSBWall.Checked) return "ОСБ";
            if (rbMDFSheetWall.Checked) return "МДФ листы";
            if (rbPVCWall.Checked) return "ПВХ";
            if (rbMDFPanelWall.Checked) return "МДФ панели";
            if (rbLiningWall.Checked) return "Вагонка";
            if (rbImitationTimberWall.Checked) return "Имитация бруса";
            return "Не выбрано";
        }

        private string GetSelectedFloor()
        {
            if (rbPlywoodFloor.Checked) return "Фанера";
            if (rbOSBFloor.Checked) return "ОСБ";
            if (rbRoughBoardFloor.Checked) return "Черновая доска";
            return "Не выбрано";
        }

        // Улучшенный метод поиска изображений в нескольких местах
        private string FindImagePath(string imageName)
        {
            // Получаем базовую директорию (где находится EXE файл)
            string baseDirectory = AppDomain.CurrentDomain.BaseDirectory;

            // Список путей для поиска
            string[] searchPaths = new string[]
            {
                // 1. Папка Images рядом с EXE (BytovkaBuilder2-master\проект\bin\Debug\Images)
                Path.Combine(baseDirectory, "Images", imageName),
                
                // 2. Папка Images на уровень выше (BytovkaBuilder2-master\проект\bin\Images)
                Path.Combine(Directory.GetParent(baseDirectory).FullName, "Images", imageName),
                
                // 3. Папка Images в корне проекта (BytovkaBuilder2-master\проект\Images)
                Path.Combine(Directory.GetParent(Directory.GetParent(baseDirectory).FullName).FullName, "Images", imageName),
                
                // 4. Папка Images на 3 уровня выше (BytovkaBuilder2-master\Images)
                Path.Combine(Directory.GetParent(Directory.GetParent(Directory.GetParent(baseDirectory).FullName).FullName).FullName, "Images", imageName),
                
                // 5. Просто в папке с EXE
                Path.Combine(baseDirectory, imageName),
                
                // 6. В папке Resources (если есть)
                Path.Combine(baseDirectory, "Resources", imageName)
            };

            // Проверяем каждый путь
            foreach (string path in searchPaths)
            {
                if (File.Exists(path))
                {
                    return path;
                }
            }

            // Если файл не найден, возвращаем первый путь для отображения ошибки
            return searchPaths[0];
        }

        private void BtnShow_Click(object sender, EventArgs e)
        {
            string imageName = GetImageFileName();

            // Ищем изображение в нескольких местах
            string imagePath = FindImagePath(imageName);

            if (File.Exists(imagePath))
            {
                try
                {
                    pictureBox.Image = Image.FromFile(imagePath);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при загрузке изображения:\n{ex.Message}",
                        "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                // Формируем информативное сообщение об ошибке
                string baseDirectory = AppDomain.CurrentDomain.BaseDirectory;
                string message = $"Файл не найден!\n\n" +
                    $"Ищем файл: {imageName}\n\n" +
                    $"Проверьте следующие папки:\n" +
                    $"1. {Path.Combine(baseDirectory, "Images")}\n" +
                    $"2. {Path.Combine(Directory.GetParent(baseDirectory).FullName, "Images")}\n" +
                    $"3. {Path.Combine(Directory.GetParent(Directory.GetParent(baseDirectory).FullName).FullName, "Images")}\n\n" +
                    $"Текущая папка программы: {baseDirectory}\n\n" +
                    $"Рекомендация:\n" +
                    $"1. Создайте папку 'Images' в папке с программой\n" +
                    $"2. Положите туда файл: {imageName}\n" +
                    $"3. Или поместите изображения в папку 'Images' в корне проекта";

                MessageBox.Show(message,
                    "Изображение не найдено",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
        }

        private void BtnExportExcel_Click(object sender, EventArgs e)
        {
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string fileName = $"ТехЗадание_Бытовка_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
            string filePath = Path.Combine(desktopPath, fileName);

            try
            {
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("Спецификация");

                    ws.Cell(1, 1).Value = "Категория";
                    ws.Cell(1, 2).Value = "Выбранный материал";
                    ws.Cell(1, 3).Value = "Примечание";

                    var headerRange = ws.Range(1, 1, 1, 3);
                    headerRange.Style.Font.Bold = true;
                    headerRange.Style.Fill.BackgroundColor = XLColor.FromArgb(52, 73, 94);
                    headerRange.Style.Font.FontColor = XLColor.White;

                    int row = 2;

                    ws.Cell(row, 1).Value = "Внешняя обшивка";
                    ws.Cell(row, 2).Value = GetSelectedExteriorMaterialName();
                    row++;

                    ws.Cell(row, 1).Value = "Внутренняя отделка стен";
                    ws.Cell(row, 2).Value = GetSelectedInterior();
                    row++;

                    ws.Cell(row, 1).Value = "Пол";
                    ws.Cell(row, 2).Value = GetSelectedFloor();
                    row++;

                    ws.Cell(row, 1).Value = "Дверь";
                    ws.Cell(row, 2).Value = chkDoor.Checked ? "Стандартная ГОСТ дверь" : "Не требуется";
                    row++;

                    if (chkWindow.Checked)
                    {
                        ws.Cell(row, 1).Value = "Окно";
                        ws.Cell(row, 2).Value = $"Размер: {numWidth.Value} x {numHeight.Value} см";
                        row++;
                    }
                    else
                    {
                        ws.Cell(row, 1).Value = "Окно";
                        ws.Cell(row, 2).Value = "Не требуется";
                        row++;
                    }

                    string wishText = txtWishes.Text;
                    if (wishText == "Напишите здесь ваши пожелания...")
                        wishText = "";

                    if (!string.IsNullOrWhiteSpace(wishText))
                    {
                        row++;
                        ws.Cell(row, 1).Value = "Дополнительные пожелания";
                        ws.Cell(row, 2).Value = wishText;
                        ws.Cell(row, 2).Style.Alignment.WrapText = true;
                        ws.Row(row).Height = 50;
                    }

                    row += 2;
                    ws.Cell(row, 1).Value = "Дата формирования";
                    ws.Cell(row, 2).Value = DateTime.Now.ToString("dd.MM.yyyy HH:mm:ss");

                    ws.Columns().AdjustToContents();
                    ws.Column(2).Width = Math.Max(ws.Column(2).Width, 50);

                    workbook.SaveAs(filePath);
                }

                MessageBox.Show($"✓ Техзадание сохранено на рабочем столе!\n\n📁 {filePath}",
                    "Готово", MessageBoxButtons.OK, MessageBoxIcon.Information);

                DialogResult result = MessageBox.Show("Открыть файл?", "Открыть Excel",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при сохранении файла:\n{ex.Message}",
                    "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
