using System;
using System.Windows;

namespace outlook_extension
{
    public enum CornerStyle
    {
        Circular,
        Continuous
    }

    public static class CornerTokens
    {
        public static double RadiusXS { get; set; } = 6;
        public static double RadiusS { get; set; } = 8;
        public static double RadiusM { get; set; } = 12;
        public static double RadiusL { get; set; } = 16;
        public static double RadiusXL { get; set; } = 24;

        public static CornerRadius ToCornerRadius(double radius)
        {
            return new CornerRadius(radius);
        }
    }
}
