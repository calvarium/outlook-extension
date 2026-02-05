using System;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace outlook_extension
{
    public class ContinuousCornerBorder : Border
    {
        public static readonly DependencyProperty CornerStyleProperty = DependencyProperty.Register(
            nameof(CornerStyle),
            typeof(CornerStyle),
            typeof(ContinuousCornerBorder),
            new FrameworkPropertyMetadata(CornerStyle.Circular, FrameworkPropertyMetadataOptions.AffectsRender, OnCornerPropertyChanged));

        public static readonly DependencyProperty CornerSmoothingProperty = DependencyProperty.Register(
            nameof(CornerSmoothing),
            typeof(double),
            typeof(ContinuousCornerBorder),
            new FrameworkPropertyMetadata(4.0, FrameworkPropertyMetadataOptions.AffectsRender, OnCornerPropertyChanged));

        private Geometry _outerGeometry;
        private Geometry _innerGeometry;
        private string _outerKey;
        private string _innerKey;

        public CornerStyle CornerStyle
        {
            get => (CornerStyle)GetValue(CornerStyleProperty);
            set => SetValue(CornerStyleProperty, value);
        }

        public double CornerSmoothing
        {
            get => (double)GetValue(CornerSmoothingProperty);
            set => SetValue(CornerSmoothingProperty, value);
        }

        protected override void OnRender(DrawingContext dc)
        {
            if (CornerStyle == CornerStyle.Circular)
            {
                base.OnRender(dc);
                if (Child != null)
                {
                    Child.Clip = null;
                }
                return;
            }

            var size = RenderSize;
            if (size.Width <= 0 || size.Height <= 0)
            {
                return;
            }

            EnsureGeometry(size);

            if (_outerGeometry != null)
            {
                if (Background != null)
                {
                    dc.DrawGeometry(Background, null, _outerGeometry);
                }

                if (BorderBrush != null)
                {
                    var thickness = Math.Max(Math.Max(BorderThickness.Left, BorderThickness.Right), Math.Max(BorderThickness.Top, BorderThickness.Bottom));
                    if (thickness > 0)
                    {
                        var pen = new Pen(BorderBrush, thickness)
                        {
                            LineJoin = PenLineJoin.Round
                        };
                        dc.DrawGeometry(null, pen, _outerGeometry);
                    }
                }
            }

            UpdateChildClip();
        }

        protected override Size ArrangeOverride(Size finalSize)
        {
            var arranged = base.ArrangeOverride(finalSize);
            UpdateChildClip();
            return arranged;
        }

        private static void OnCornerPropertyChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is ContinuousCornerBorder border)
            {
                border.InvalidateGeometry();
            }
        }

        private void InvalidateGeometry()
        {
            _outerKey = null;
            _innerKey = null;
            _outerGeometry = null;
            _innerGeometry = null;
            InvalidateVisual();
        }

        private void UpdateChildClip()
        {
            if (Child == null)
            {
                return;
            }

            if (CornerStyle != CornerStyle.Continuous || _innerGeometry == null)
            {
                Child.Clip = null;
                return;
            }

            Child.Clip = _innerGeometry.CloneCurrentValue();
        }

        private void EnsureGeometry(Size size)
        {
            var smoothing = Math.Max(2.0, CornerSmoothing);
            var cornerRadius = CornerRadius;

            var outerKey = BuildKey(size, cornerRadius, smoothing, BorderThickness, Padding, "outer");
            if (!string.Equals(outerKey, _outerKey, StringComparison.Ordinal))
            {
                _outerGeometry = CreateContinuousGeometry(new Rect(0, 0, size.Width, size.Height), cornerRadius, smoothing);
                _outerGeometry?.Freeze();
                _outerKey = outerKey;
            }

            var innerSize = new Size(
                Math.Max(0, size.Width - BorderThickness.Left - BorderThickness.Right - Padding.Left - Padding.Right),
                Math.Max(0, size.Height - BorderThickness.Top - BorderThickness.Bottom - Padding.Top - Padding.Bottom));

            if (innerSize.Width <= 0 || innerSize.Height <= 0)
            {
                _innerGeometry = null;
                _innerKey = null;
                return;
            }

            var innerCornerRadius = DeflateCornerRadius(cornerRadius, BorderThickness, Padding);
            var innerKey = BuildKey(innerSize, innerCornerRadius, smoothing, new Thickness(0), new Thickness(0), "inner");
            if (!string.Equals(innerKey, _innerKey, StringComparison.Ordinal))
            {
                _innerGeometry = CreateContinuousGeometry(new Rect(0, 0, innerSize.Width, innerSize.Height), innerCornerRadius, smoothing);
                _innerGeometry?.Freeze();
                _innerKey = innerKey;
            }
        }

        private static string BuildKey(Size size, CornerRadius radius, double smoothing, Thickness border, Thickness padding, string suffix)
        {
            var culture = CultureInfo.InvariantCulture;
            return string.Format(culture,
                "{0}:{1:0.###}:{2:0.###}:{3:0.###}:{4:0.###}:{5:0.###}:{6:0.###}:{7:0.###}:{8:0.###}:{9:0.###}:{10:0.###}:{11:0.###}:{12:0.###}:{13:0.###}:{14:0.###}:{15:0.###}:{16:0.###}",
                suffix,
                size.Width,
                size.Height,
                radius.TopLeft,
                radius.TopRight,
                radius.BottomRight,
                radius.BottomLeft,
                smoothing,
                border.Left,
                border.Top,
                border.Right,
                border.Bottom,
                padding.Left,
                padding.Top,
                padding.Right,
                padding.Bottom,
                0.0);
        }

        private static CornerRadius DeflateCornerRadius(CornerRadius radius, Thickness border, Thickness padding)
        {
            var leftInset = border.Left + padding.Left;
            var topInset = border.Top + padding.Top;
            var rightInset = border.Right + padding.Right;
            var bottomInset = border.Bottom + padding.Bottom;

            return new CornerRadius(
                Math.Max(0, radius.TopLeft - Math.Max(leftInset, topInset)),
                Math.Max(0, radius.TopRight - Math.Max(rightInset, topInset)),
                Math.Max(0, radius.BottomRight - Math.Max(rightInset, bottomInset)),
                Math.Max(0, radius.BottomLeft - Math.Max(leftInset, bottomInset)));
        }

        private static Geometry CreateContinuousGeometry(Rect rect, CornerRadius radius, double smoothing)
        {
            var width = rect.Width;
            var height = rect.Height;
            if (width <= 0 || height <= 0)
            {
                return null;
            }

            var topLeft = Math.Max(0, radius.TopLeft);
            var topRight = Math.Max(0, radius.TopRight);
            var bottomRight = Math.Max(0, radius.BottomRight);
            var bottomLeft = Math.Max(0, radius.BottomLeft);

            var scale = 1.0;
            if (topLeft + topRight > width)
            {
                scale = Math.Min(scale, width / (topLeft + topRight));
            }
            if (bottomLeft + bottomRight > width)
            {
                scale = Math.Min(scale, width / (bottomLeft + bottomRight));
            }
            if (topLeft + bottomLeft > height)
            {
                scale = Math.Min(scale, height / (topLeft + bottomLeft));
            }
            if (topRight + bottomRight > height)
            {
                scale = Math.Min(scale, height / (topRight + bottomRight));
            }

            if (scale < 1.0)
            {
                topLeft *= scale;
                topRight *= scale;
                bottomRight *= scale;
                bottomLeft *= scale;
            }

            var maxRadiusX = width / 2.0;
            var maxRadiusY = height / 2.0;
            topLeft = Math.Min(topLeft, Math.Min(maxRadiusX, maxRadiusY));
            topRight = Math.Min(topRight, Math.Min(maxRadiusX, maxRadiusY));
            bottomRight = Math.Min(bottomRight, Math.Min(maxRadiusX, maxRadiusY));
            bottomLeft = Math.Min(bottomLeft, Math.Min(maxRadiusX, maxRadiusY));

            var geometry = new StreamGeometry
            {
                FillRule = FillRule.Nonzero
            };
            using (var ctx = geometry.Open())
            {
                var topLeftStart = new Point(rect.Left + topLeft, rect.Top);
                ctx.BeginFigure(topLeftStart, true, true);

                ctx.LineTo(new Point(rect.Right - topRight, rect.Top), true, false);
                AddCorner(ctx, rect, topRight, smoothing, CornerPosition.TopRight);

                ctx.LineTo(new Point(rect.Right, rect.Bottom - bottomRight), true, false);
                AddCorner(ctx, rect, bottomRight, smoothing, CornerPosition.BottomRight);

                ctx.LineTo(new Point(rect.Left + bottomLeft, rect.Bottom), true, false);
                AddCorner(ctx, rect, bottomLeft, smoothing, CornerPosition.BottomLeft);

                ctx.LineTo(new Point(rect.Left, rect.Top + topLeft), true, false);
                AddCorner(ctx, rect, topLeft, smoothing, CornerPosition.TopLeft);
            }

            return geometry;
        }

        private static void AddCorner(StreamGeometryContext ctx, Rect rect, double radius, double smoothing, CornerPosition position)
        {
            if (radius <= 0)
            {
                return;
            }

            var normalized = Math.Max(0, smoothing - 2.0) / 6.0;
            if (normalized > 1.0)
            {
                normalized = 1.0;
            }

            var circleKappa = 0.5522847498307936;
            var squircleKappa = 0.75;
            var kappa = circleKappa + (squircleKappa - circleKappa) * normalized;
            var handle = radius * kappa;

            switch (position)
            {
                case CornerPosition.TopRight:
                    ctx.BezierTo(
                        new Point(rect.Right - radius + handle, rect.Top),
                        new Point(rect.Right, rect.Top + radius - handle),
                        new Point(rect.Right, rect.Top + radius),
                        true,
                        false);
                    break;
                case CornerPosition.BottomRight:
                    ctx.BezierTo(
                        new Point(rect.Right, rect.Bottom - radius + handle),
                        new Point(rect.Right - radius + handle, rect.Bottom),
                        new Point(rect.Right - radius, rect.Bottom),
                        true,
                        false);
                    break;
                case CornerPosition.BottomLeft:
                    ctx.BezierTo(
                        new Point(rect.Left + radius - handle, rect.Bottom),
                        new Point(rect.Left, rect.Bottom - radius + handle),
                        new Point(rect.Left, rect.Bottom - radius),
                        true,
                        false);
                    break;
                case CornerPosition.TopLeft:
                    ctx.BezierTo(
                        new Point(rect.Left, rect.Top + radius - handle),
                        new Point(rect.Left + radius - handle, rect.Top),
                        new Point(rect.Left + radius, rect.Top),
                        true,
                        false);
                    break;
            }
        }

        private enum CornerPosition
        {
            TopLeft,
            TopRight,
            BottomRight,
            BottomLeft
        }
    }
}
