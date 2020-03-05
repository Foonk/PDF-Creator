namespace PDF_Creator
{
    partial class MainForm
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.wordToHtmlLink = new System.Windows.Forms.LinkLabel();
            this.progressLabel = new System.Windows.Forms.Label();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.excelFileSelectBtn = new System.Windows.Forms.Button();
            this.excelFileSelectLabel = new System.Windows.Forms.Label();
            this.htmlFileSelectBtn = new System.Windows.Forms.Button();
            this.htmlFileSelectLabel = new System.Windows.Forms.Label();
            this.loadDataFromFileBtn = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.folderWithFilesSelectBtn = new System.Windows.Forms.Button();
            this.filesPathLabel = new System.Windows.Forms.Label();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.setPassword = new System.Windows.Forms.CheckBox();
            this.paramsLabel = new System.Windows.Forms.Label();
            this.addParamBtn = new System.Windows.Forms.Button();
            this.pTitleLabel = new System.Windows.Forms.Label();
            this.newParamTitle = new System.Windows.Forms.TextBox();
            this.newParamRowNumber = new System.Windows.Forms.TextBox();
            this.pRowNumberLabel = new System.Windows.Forms.Label();
            this.newParamCSSclassName = new System.Windows.Forms.TextBox();
            this.pCSSclassNameLabel = new System.Windows.Forms.Label();
            this.parametersLabel = new System.Windows.Forms.Label();
            this.passRowNumberLabel = new System.Windows.Forms.Label();
            this.passRowNumber = new System.Windows.Forms.TextBox();
            this.deleteParamsBtn = new System.Windows.Forms.Button();
            this.newParamDataType = new System.Windows.Forms.ComboBox();
            this.pDataTypeLabel = new System.Windows.Forms.Label();
            this.saveConfigBtn = new System.Windows.Forms.Button();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.loadConfigBtn = new System.Windows.Forms.Button();
            this.exampleLabel = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // wordToHtmlLink
            // 
            this.wordToHtmlLink.AutoSize = true;
            this.wordToHtmlLink.Location = new System.Drawing.Point(12, 9);
            this.wordToHtmlLink.Name = "wordToHtmlLink";
            this.wordToHtmlLink.Size = new System.Drawing.Size(146, 13);
            this.wordToHtmlLink.TabIndex = 49;
            this.wordToHtmlLink.TabStop = true;
            this.wordToHtmlLink.Text = "Преобразовать Word в html";
            this.wordToHtmlLink.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.wordToHtmlLink_LinkClicked);
            // 
            // progressLabel
            // 
            this.progressLabel.AutoSize = true;
            this.progressLabel.Location = new System.Drawing.Point(539, 33);
            this.progressLabel.Name = "progressLabel";
            this.progressLabel.Size = new System.Drawing.Size(0, 13);
            this.progressLabel.TabIndex = 48;
            // 
            // progressBar
            // 
            this.progressBar.Location = new System.Drawing.Point(433, 28);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(100, 23);
            this.progressBar.TabIndex = 47;
            // 
            // excelFileSelectBtn
            // 
            this.excelFileSelectBtn.Location = new System.Drawing.Point(12, 28);
            this.excelFileSelectBtn.Name = "excelFileSelectBtn";
            this.excelFileSelectBtn.Size = new System.Drawing.Size(146, 23);
            this.excelFileSelectBtn.TabIndex = 38;
            this.excelFileSelectBtn.Text = "Выбери файл Excel";
            this.excelFileSelectBtn.UseVisualStyleBackColor = true;
            this.excelFileSelectBtn.Click += new System.EventHandler(this.excelFileSelectBtn_Click);
            // 
            // excelFileSelectLabel
            // 
            this.excelFileSelectLabel.AutoSize = true;
            this.excelFileSelectLabel.Location = new System.Drawing.Point(12, 54);
            this.excelFileSelectLabel.Name = "excelFileSelectLabel";
            this.excelFileSelectLabel.Size = new System.Drawing.Size(0, 13);
            this.excelFileSelectLabel.TabIndex = 37;
            // 
            // htmlFileSelectBtn
            // 
            this.htmlFileSelectBtn.Location = new System.Drawing.Point(12, 70);
            this.htmlFileSelectBtn.Name = "htmlFileSelectBtn";
            this.htmlFileSelectBtn.Size = new System.Drawing.Size(146, 23);
            this.htmlFileSelectBtn.TabIndex = 41;
            this.htmlFileSelectBtn.Text = "Выбери файл html";
            this.htmlFileSelectBtn.UseVisualStyleBackColor = true;
            this.htmlFileSelectBtn.Click += new System.EventHandler(this.htmlFileSelectBtn_Click);
            // 
            // htmlFileSelectLabel
            // 
            this.htmlFileSelectLabel.AutoSize = true;
            this.htmlFileSelectLabel.Location = new System.Drawing.Point(12, 96);
            this.htmlFileSelectLabel.Name = "htmlFileSelectLabel";
            this.htmlFileSelectLabel.Size = new System.Drawing.Size(0, 13);
            this.htmlFileSelectLabel.TabIndex = 40;
            // 
            // loadDataFromFileBtn
            // 
            this.loadDataFromFileBtn.Location = new System.Drawing.Point(277, 28);
            this.loadDataFromFileBtn.Name = "loadDataFromFileBtn";
            this.loadDataFromFileBtn.Size = new System.Drawing.Size(146, 23);
            this.loadDataFromFileBtn.TabIndex = 39;
            this.loadDataFromFileBtn.Text = "Загрузить";
            this.loadDataFromFileBtn.UseVisualStyleBackColor = true;
            this.loadDataFromFileBtn.Click += new System.EventHandler(this.loadDataFromFileBtn_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // folderWithFilesSelectBtn
            // 
            this.folderWithFilesSelectBtn.Location = new System.Drawing.Point(12, 112);
            this.folderWithFilesSelectBtn.Name = "folderWithFilesSelectBtn";
            this.folderWithFilesSelectBtn.Size = new System.Drawing.Size(146, 23);
            this.folderWithFilesSelectBtn.TabIndex = 51;
            this.folderWithFilesSelectBtn.Text = "Выбери куда сохранить";
            this.folderWithFilesSelectBtn.UseVisualStyleBackColor = true;
            this.folderWithFilesSelectBtn.Click += new System.EventHandler(this.folderWithFilesSelectBtn_Click);
            // 
            // filesPathLabel
            // 
            this.filesPathLabel.AutoSize = true;
            this.filesPathLabel.Location = new System.Drawing.Point(12, 138);
            this.filesPathLabel.Name = "filesPathLabel";
            this.filesPathLabel.Size = new System.Drawing.Size(0, 13);
            this.filesPathLabel.TabIndex = 50;
            // 
            // setPassword
            // 
            this.setPassword.AutoSize = true;
            this.setPassword.Location = new System.Drawing.Point(168, 116);
            this.setPassword.Name = "setPassword";
            this.setPassword.Size = new System.Drawing.Size(86, 17);
            this.setPassword.TabIndex = 52;
            this.setPassword.Text = "Запаролить";
            this.setPassword.UseVisualStyleBackColor = true;
            // 
            // paramsLabel
            // 
            this.paramsLabel.AutoSize = true;
            this.paramsLabel.Location = new System.Drawing.Point(12, 155);
            this.paramsLabel.Name = "paramsLabel";
            this.paramsLabel.Size = new System.Drawing.Size(69, 13);
            this.paramsLabel.TabIndex = 53;
            this.paramsLabel.Text = "Параметры:";
            // 
            // addParamBtn
            // 
            this.addParamBtn.Location = new System.Drawing.Point(9, 221);
            this.addParamBtn.Name = "addParamBtn";
            this.addParamBtn.Size = new System.Drawing.Size(294, 23);
            this.addParamBtn.TabIndex = 54;
            this.addParamBtn.Text = "Добавить";
            this.addParamBtn.UseVisualStyleBackColor = true;
            this.addParamBtn.Click += new System.EventHandler(this.addParamBtn_Click);
            // 
            // pTitleLabel
            // 
            this.pTitleLabel.AutoSize = true;
            this.pTitleLabel.Location = new System.Drawing.Point(33, 178);
            this.pTitleLabel.Name = "pTitleLabel";
            this.pTitleLabel.Size = new System.Drawing.Size(57, 13);
            this.pTitleLabel.TabIndex = 55;
            this.pTitleLabel.Text = "Название";
            // 
            // newParamTitle
            // 
            this.newParamTitle.Location = new System.Drawing.Point(9, 195);
            this.newParamTitle.Name = "newParamTitle";
            this.newParamTitle.Size = new System.Drawing.Size(100, 20);
            this.newParamTitle.TabIndex = 56;
            // 
            // newParamRowNumber
            // 
            this.newParamRowNumber.Location = new System.Drawing.Point(115, 195);
            this.newParamRowNumber.Name = "newParamRowNumber";
            this.newParamRowNumber.Size = new System.Drawing.Size(82, 20);
            this.newParamRowNumber.TabIndex = 58;
            // 
            // pRowNumberLabel
            // 
            this.pRowNumberLabel.AutoSize = true;
            this.pRowNumberLabel.Location = new System.Drawing.Point(112, 178);
            this.pRowNumberLabel.Name = "pRowNumberLabel";
            this.pRowNumberLabel.Size = new System.Drawing.Size(85, 13);
            this.pRowNumberLabel.TabIndex = 57;
            this.pRowNumberLabel.Text = "Номер столбца";
            // 
            // newParamCSSclassName
            // 
            this.newParamCSSclassName.Location = new System.Drawing.Point(203, 195);
            this.newParamCSSclassName.Name = "newParamCSSclassName";
            this.newParamCSSclassName.Size = new System.Drawing.Size(100, 20);
            this.newParamCSSclassName.TabIndex = 60;
            // 
            // pCSSclassNameLabel
            // 
            this.pCSSclassNameLabel.AutoSize = true;
            this.pCSSclassNameLabel.Location = new System.Drawing.Point(209, 178);
            this.pCSSclassNameLabel.Name = "pCSSclassNameLabel";
            this.pCSSclassNameLabel.Size = new System.Drawing.Size(92, 13);
            this.pCSSclassNameLabel.TabIndex = 59;
            this.pCSSclassNameLabel.Text = "Имя CSS класса";
            // 
            // parametersLabel
            // 
            this.parametersLabel.AutoSize = true;
            this.parametersLabel.Location = new System.Drawing.Point(6, 247);
            this.parametersLabel.MaximumSize = new System.Drawing.Size(260, 0);
            this.parametersLabel.Name = "parametersLabel";
            this.parametersLabel.Size = new System.Drawing.Size(0, 13);
            this.parametersLabel.TabIndex = 61;
            // 
            // passRowNumberLabel
            // 
            this.passRowNumberLabel.AutoSize = true;
            this.passRowNumberLabel.Location = new System.Drawing.Point(275, 117);
            this.passRowNumberLabel.Name = "passRowNumberLabel";
            this.passRowNumberLabel.Size = new System.Drawing.Size(124, 13);
            this.passRowNumberLabel.TabIndex = 62;
            this.passRowNumberLabel.Text = "Номер столбца пароля";
            // 
            // passRowNumber
            // 
            this.passRowNumber.Location = new System.Drawing.Point(405, 114);
            this.passRowNumber.Name = "passRowNumber";
            this.passRowNumber.Size = new System.Drawing.Size(42, 20);
            this.passRowNumber.TabIndex = 63;
            // 
            // deleteParamsBtn
            // 
            this.deleteParamsBtn.Location = new System.Drawing.Point(309, 221);
            this.deleteParamsBtn.Name = "deleteParamsBtn";
            this.deleteParamsBtn.Size = new System.Drawing.Size(72, 23);
            this.deleteParamsBtn.TabIndex = 64;
            this.deleteParamsBtn.Text = "Очистить";
            this.deleteParamsBtn.UseVisualStyleBackColor = true;
            this.deleteParamsBtn.Click += new System.EventHandler(this.deleteParamsBtn_Click);
            // 
            // newParamDataType
            // 
            this.newParamDataType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.newParamDataType.FormattingEnabled = true;
            this.newParamDataType.Items.AddRange(new object[] {
            "Строка",
            "Целое число",
            "Дробное число",
            "Дата",
            "Время"});
            this.newParamDataType.Location = new System.Drawing.Point(309, 194);
            this.newParamDataType.Name = "newParamDataType";
            this.newParamDataType.Size = new System.Drawing.Size(123, 21);
            this.newParamDataType.TabIndex = 65;
            // 
            // pDataTypeLabel
            // 
            this.pDataTypeLabel.AutoSize = true;
            this.pDataTypeLabel.Location = new System.Drawing.Point(333, 178);
            this.pDataTypeLabel.Name = "pDataTypeLabel";
            this.pDataTypeLabel.Size = new System.Drawing.Size(66, 13);
            this.pDataTypeLabel.TabIndex = 66;
            this.pDataTypeLabel.Text = "Тип данных";
            // 
            // saveConfigBtn
            // 
            this.saveConfigBtn.Location = new System.Drawing.Point(387, 221);
            this.saveConfigBtn.Name = "saveConfigBtn";
            this.saveConfigBtn.Size = new System.Drawing.Size(75, 23);
            this.saveConfigBtn.TabIndex = 67;
            this.saveConfigBtn.Text = "Сохранить";
            this.saveConfigBtn.UseVisualStyleBackColor = true;
            this.saveConfigBtn.Click += new System.EventHandler(this.saveConfigBtn_Click);
            // 
            // loadConfigBtn
            // 
            this.loadConfigBtn.Location = new System.Drawing.Point(468, 221);
            this.loadConfigBtn.Name = "loadConfigBtn";
            this.loadConfigBtn.Size = new System.Drawing.Size(75, 23);
            this.loadConfigBtn.TabIndex = 68;
            this.loadConfigBtn.Text = "Загрузить";
            this.loadConfigBtn.UseVisualStyleBackColor = true;
            this.loadConfigBtn.Click += new System.EventHandler(this.loadConfigBtn_Click);
            // 
            // exampleLabel
            // 
            this.exampleLabel.AutoSize = true;
            this.exampleLabel.Location = new System.Drawing.Point(165, 75);
            this.exampleLabel.Name = "exampleLabel";
            this.exampleLabel.Size = new System.Drawing.Size(389, 13);
            this.exampleLabel.TabIndex = 69;
            this.exampleLabel.Text = "Пример параметра для html: <span class=\"fio\" style=\"font-size:11pt;\"></span>";
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(591, 560);
            this.Controls.Add(this.exampleLabel);
            this.Controls.Add(this.loadConfigBtn);
            this.Controls.Add(this.saveConfigBtn);
            this.Controls.Add(this.pDataTypeLabel);
            this.Controls.Add(this.newParamDataType);
            this.Controls.Add(this.deleteParamsBtn);
            this.Controls.Add(this.passRowNumber);
            this.Controls.Add(this.passRowNumberLabel);
            this.Controls.Add(this.parametersLabel);
            this.Controls.Add(this.newParamCSSclassName);
            this.Controls.Add(this.pCSSclassNameLabel);
            this.Controls.Add(this.newParamRowNumber);
            this.Controls.Add(this.pRowNumberLabel);
            this.Controls.Add(this.newParamTitle);
            this.Controls.Add(this.pTitleLabel);
            this.Controls.Add(this.addParamBtn);
            this.Controls.Add(this.paramsLabel);
            this.Controls.Add(this.setPassword);
            this.Controls.Add(this.folderWithFilesSelectBtn);
            this.Controls.Add(this.filesPathLabel);
            this.Controls.Add(this.wordToHtmlLink);
            this.Controls.Add(this.progressLabel);
            this.Controls.Add(this.progressBar);
            this.Controls.Add(this.excelFileSelectBtn);
            this.Controls.Add(this.excelFileSelectLabel);
            this.Controls.Add(this.htmlFileSelectBtn);
            this.Controls.Add(this.htmlFileSelectLabel);
            this.Controls.Add(this.loadDataFromFileBtn);
            this.Name = "MainForm";
            this.Text = "PDF Creator";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.LinkLabel wordToHtmlLink;
        private System.Windows.Forms.Label progressLabel;
        private System.Windows.Forms.ProgressBar progressBar;
        private System.Windows.Forms.Button excelFileSelectBtn;
        private System.Windows.Forms.Label excelFileSelectLabel;
        private System.Windows.Forms.Button htmlFileSelectBtn;
        private System.Windows.Forms.Label htmlFileSelectLabel;
        private System.Windows.Forms.Button loadDataFromFileBtn;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Button folderWithFilesSelectBtn;
        private System.Windows.Forms.Label filesPathLabel;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.CheckBox setPassword;
        private System.Windows.Forms.Label paramsLabel;
        private System.Windows.Forms.Button addParamBtn;
        private System.Windows.Forms.Label pTitleLabel;
        private System.Windows.Forms.TextBox newParamTitle;
        private System.Windows.Forms.TextBox newParamRowNumber;
        private System.Windows.Forms.Label pRowNumberLabel;
        private System.Windows.Forms.TextBox newParamCSSclassName;
        private System.Windows.Forms.Label pCSSclassNameLabel;
        private System.Windows.Forms.Label parametersLabel;
        private System.Windows.Forms.Label passRowNumberLabel;
        private System.Windows.Forms.TextBox passRowNumber;
        private System.Windows.Forms.Button deleteParamsBtn;
        private System.Windows.Forms.ComboBox newParamDataType;
        private System.Windows.Forms.Label pDataTypeLabel;
        private System.Windows.Forms.Button saveConfigBtn;
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
        private System.Windows.Forms.Button loadConfigBtn;
        private System.Windows.Forms.Label exampleLabel;
    }
}

