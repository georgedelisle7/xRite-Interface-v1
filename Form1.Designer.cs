namespace xRite_Interface_v1
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            toolConnect = new Button();
            messageBar = new TextBox();
            measure = new Button();
            reset = new Button();
            inputcount = new RichTextBox();
            progressBar1 = new ProgressBar();
            measureStatus = new TextBox();
            datafill = new RichTextBox();
            date = new DateTimePicker();
            done = new Button();
            SuspendLayout();
            // 
            // toolConnect
            // 
            toolConnect.Location = new Point(327, 236);
            toolConnect.Name = "toolConnect";
            toolConnect.Size = new Size(227, 52);
            toolConnect.TabIndex = 0;
            toolConnect.Text = "Connect Tool ";
            toolConnect.UseVisualStyleBackColor = true;
            toolConnect.Click += toolConnect_Click;
            // 
            // messageBar
            // 
            messageBar.Enabled = false;
            messageBar.Font = new Font("Segoe UI", 12F, FontStyle.Regular, GraphicsUnit.Point);
            messageBar.Location = new Point(12, 34);
            messageBar.Name = "messageBar";
            messageBar.PlaceholderText = "Welcome to Xrite Interface";
            messageBar.Size = new Size(893, 34);
            messageBar.TabIndex = 1;
            messageBar.TextAlign = HorizontalAlignment.Center;
            // 
            // measure
            // 
            measure.Location = new Point(327, 402);
            measure.Name = "measure";
            measure.Size = new Size(227, 52);
            measure.TabIndex = 4;
            measure.Text = "Measure";
            measure.UseVisualStyleBackColor = true;
            measure.Click += measure_Click;
            // 
            // reset
            // 
            reset.Location = new Point(449, 505);
            reset.Name = "reset";
            reset.Size = new Size(105, 42);
            reset.TabIndex = 5;
            reset.Text = "Reset";
            reset.UseVisualStyleBackColor = true;
            reset.Click += reset_Click;
            // 
            // inputcount
            // 
            inputcount.BorderStyle = BorderStyle.FixedSingle;
            inputcount.Location = new Point(327, 338);
            inputcount.Name = "inputcount";
            inputcount.Size = new Size(227, 30);
            inputcount.TabIndex = 6;
            inputcount.Text = "Input # of Measurements";
            inputcount.TextChanged += inputcount_TextChanged;
            // 
            // progressBar1
            // 
            progressBar1.Location = new Point(176, 78);
            progressBar1.MarqueeAnimationSpeed = 50;
            progressBar1.Name = "progressBar1";
            progressBar1.Size = new Size(550, 34);
            progressBar1.TabIndex = 7;
            // 
            // measureStatus
            // 
            measureStatus.Font = new Font("Segoe UI", 12F, FontStyle.Bold, GraphicsUnit.Point);
            measureStatus.Location = new Point(763, 78);
            measureStatus.Name = "measureStatus";
            measureStatus.PlaceholderText = "0/0";
            measureStatus.ReadOnly = true;
            measureStatus.Size = new Size(91, 34);
            measureStatus.TabIndex = 8;
            measureStatus.TextAlign = HorizontalAlignment.Center;
            // 
            // datafill
            // 
            datafill.Location = new Point(35, 589);
            datafill.Name = "datafill";
            datafill.Size = new Size(846, 544);
            datafill.TabIndex = 9;
            datafill.Text = "";
            // 
            // date
            // 
            date.ImeMode = ImeMode.NoControl;
            date.Location = new Point(276, 163);
            date.Name = "date";
            date.Size = new Size(335, 27);
            date.TabIndex = 11;
            // 
            // done
            // 
            done.Location = new Point(327, 505);
            done.Name = "done";
            done.Size = new Size(105, 42);
            done.TabIndex = 12;
            done.Text = "Done";
            done.UseVisualStyleBackColor = true;
            done.Click += done_Click;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(8F, 20F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(917, 1185);
            Controls.Add(done);
            Controls.Add(date);
            Controls.Add(datafill);
            Controls.Add(measureStatus);
            Controls.Add(progressBar1);
            Controls.Add(inputcount);
            Controls.Add(reset);
            Controls.Add(measure);
            Controls.Add(messageBar);
            Controls.Add(toolConnect);
            Name = "Form1";
            Text = "XRite Interface";
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Button toolConnect;
        private TextBox messageBar;
        private Button measure;
        private Button reset;
        private RichTextBox inputcount;
        private ProgressBar progressBar1;
        private TextBox measureStatus;
        private RichTextBox datafill;
        private DateTimePicker date;
        private Button done;
    }
}