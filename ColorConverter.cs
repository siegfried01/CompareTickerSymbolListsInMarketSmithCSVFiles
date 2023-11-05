using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


    class ColorConverter
    {
        public static (int R,int G,int B) HlsToRgba(double h, double l, double s)
        {
            int r, g, b;
            if (s == 0)
            {
                // If saturation is 0, the color is grayscale, so R, G, and B are all equal to the lightness value.
                r = g = b = (int)Math.Round(l * 255);
                return (r, g, b);
            }

            double q = l < 0.5 ? l * (1 + s) : l + s - l * s;
            double p = 2 * l - q;

            r = (int)Math.Round(Hue2Rgb(p, q, h + 1 / 3.0) * 255);
            g = (int)Math.Round(Hue2Rgb(p, q, h) * 255);
            b = (int)Math.Round(Hue2Rgb(p, q, h - 1 / 3.0) * 255);
            return (r , g, b);
            
        }

        private static double Hue2Rgb(double p, double q, double t)
        {
            if (t < 0) t += 1;
            if (t > 1) t -= 1;
            if (t < 1 / 6.0) return p + (q - p) * 6 * t;
            if (t < 1 / 2.0) return q;
            if (t < 2 / 3.0) return p + (q - p) * (2 / 3.0 - t) * 6;
            return p;
        }
        public static string ToHexString(System.Drawing.Color c)
        {
            return String.Format("#{0:X2}{1:X2}{2:X2}", c.R, c.G, c.B);
        }

    }

