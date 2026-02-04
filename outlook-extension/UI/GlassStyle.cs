using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace outlook_extension
{
    internal static class GlassStyle
    {
        private const int WmNclbuttondown = 0xA1;
        private const int HtCaption = 0x2;

        [DllImport("user32.dll")]
        private static extern bool ReleaseCapture();

        [DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, int wParam, int lParam);

        public static void ApplyFormStyle(Form form)
        {
            form.FormBorderStyle = FormBorderStyle.None;
            form.BackColor = Color.FromArgb(20, 20, 26);
            form.ForeColor = Color.White;
            form.Font = new Font("Segoe UI", 10F, FontStyle.Regular, GraphicsUnit.Point);
            form.Padding = new Padding(22);
            form.StartPosition = FormStartPosition.CenterParent;
            form.KeyPreview = true;
            form.Opacity = 0.97;
            form.AutoScaleMode = AutoScaleMode.Dpi;
            form.Resize += (sender, args) => ApplyRoundedRegion(form, 22);
            ApplyRoundedRegion(form, 22);
        }

        public static Panel CreateGlassCard(int radius = 18, int padding = 18)
        {
            var panel = new Panel
            {
                BackColor = Color.FromArgb(70, 255, 255, 255),
                Padding = new Padding(padding),
                Dock = DockStyle.Fill
            };
            panel.Paint += (sender, args) => DrawGlassBorder(panel, args.Graphics, radius);
            panel.Resize += (sender, args) => ApplyRoundedRegion(panel, radius);
            ApplyRoundedRegion(panel, radius);
            return panel;
        }

        public static Panel CreateInputPanel(int radius = 14, int padding = 10)
        {
            var panel = new Panel
            {
                BackColor = Color.FromArgb(50, 255, 255, 255),
                Padding = new Padding(padding),
                Dock = DockStyle.Fill
            };
            panel.Paint += (sender, args) => DrawGlassBorder(panel, args.Graphics, radius);
            panel.Resize += (sender, args) => ApplyRoundedRegion(panel, radius);
            ApplyRoundedRegion(panel, radius);
            return panel;
        }

        public static void StyleGlassButton(Button button)
        {
            button.FlatStyle = FlatStyle.Flat;
            button.FlatAppearance.BorderSize = 0;
            button.BackColor = Color.FromArgb(90, 255, 255, 255);
            button.ForeColor = Color.White;
            button.Padding = new Padding(12, 6, 12, 6);
            button.Height = 36;
            button.Font = new Font("Segoe UI Semibold", 9.5F, FontStyle.Bold, GraphicsUnit.Point);
            button.MouseEnter += (sender, args) =>
            {
                button.BackColor = Color.FromArgb(130, 255, 255, 255);
            };
            button.MouseLeave += (sender, args) =>
            {
                button.BackColor = Color.FromArgb(90, 255, 255, 255);
            };
        }

        public static void StyleSubtleButton(Button button)
        {
            button.FlatStyle = FlatStyle.Flat;
            button.FlatAppearance.BorderSize = 0;
            button.BackColor = Color.FromArgb(40, 255, 255, 255);
            button.ForeColor = Color.White;
            button.Padding = new Padding(10, 4, 10, 4);
            button.Height = 32;
            button.Font = new Font("Segoe UI", 9F, FontStyle.Regular, GraphicsUnit.Point);
            button.MouseEnter += (sender, args) =>
            {
                button.BackColor = Color.FromArgb(90, 255, 255, 255);
            };
            button.MouseLeave += (sender, args) =>
            {
                button.BackColor = Color.FromArgb(40, 255, 255, 255);
            };
        }

        public static void StyleTextInput(TextBox textBox)
        {
            textBox.BorderStyle = BorderStyle.None;
            textBox.BackColor = Color.FromArgb(30, 255, 255, 255);
            textBox.ForeColor = Color.White;
            textBox.Font = new Font("Segoe UI", 10.5F, FontStyle.Regular, GraphicsUnit.Point);
        }

        public static void StyleListBox(ListBox listBox)
        {
            listBox.BorderStyle = BorderStyle.None;
            listBox.BackColor = Color.FromArgb(30, 255, 255, 255);
            listBox.ForeColor = Color.White;
            listBox.DrawMode = DrawMode.OwnerDrawFixed;
            listBox.ItemHeight = 32;
            listBox.IntegralHeight = false;
        }

        public static void StyleCheckBox(CheckBox checkBox)
        {
            checkBox.ForeColor = Color.WhiteSmoke;
            checkBox.Font = new Font("Segoe UI", 9.5F, FontStyle.Regular, GraphicsUnit.Point);
        }

        public static void StyleNumericInput(NumericUpDown numericUpDown)
        {
            numericUpDown.BorderStyle = BorderStyle.None;
            numericUpDown.BackColor = Color.FromArgb(30, 255, 255, 255);
            numericUpDown.ForeColor = Color.White;
            numericUpDown.Font = new Font("Segoe UI", 10.5F, FontStyle.Regular, GraphicsUnit.Point);
        }

        public static void ApplyRoundedRegion(Control control, int radius)
        {
            var bounds = control.ClientRectangle;
            if (bounds.Width <= 0 || bounds.Height <= 0)
            {
                return;
            }

            using (var path = new GraphicsPath())
            {
                int diameter = radius * 2;
                path.AddArc(bounds.X, bounds.Y, diameter, diameter, 180, 90);
                path.AddArc(bounds.Right - diameter, bounds.Y, diameter, diameter, 270, 90);
                path.AddArc(bounds.Right - diameter, bounds.Bottom - diameter, diameter, diameter, 0, 90);
                path.AddArc(bounds.X, bounds.Bottom - diameter, diameter, diameter, 90, 90);
                path.CloseAllFigures();
                control.Region = new Region(path);
            }
        }

        public static void EnableDrag(Control control)
        {
            control.MouseDown += (sender, args) =>
            {
                if (args.Button == MouseButtons.Left)
                {
                    ReleaseCapture();
                    SendMessage(control.FindForm().Handle, WmNclbuttondown, HtCaption, 0);
                }
            };
        }

        private static void DrawGlassBorder(Control control, Graphics graphics, int radius)
        {
            graphics.SmoothingMode = SmoothingMode.AntiAlias;
            var rect = new Rectangle(1, 1, control.Width - 2, control.Height - 2);
            using (var path = new GraphicsPath())
            {
                int diameter = radius * 2;
                path.AddArc(rect.X, rect.Y, diameter, diameter, 180, 90);
                path.AddArc(rect.Right - diameter, rect.Y, diameter, diameter, 270, 90);
                path.AddArc(rect.Right - diameter, rect.Bottom - diameter, diameter, diameter, 0, 90);
                path.AddArc(rect.X, rect.Bottom - diameter, diameter, diameter, 90, 90);
                path.CloseFigure();
                using (var pen = new Pen(Color.FromArgb(90, 255, 255, 255), 1))
                {
                    graphics.DrawPath(pen, path);
                }
            }
        }
    }
}
