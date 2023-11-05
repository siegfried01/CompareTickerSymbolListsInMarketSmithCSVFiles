using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

public class ExcelStyleSaturationRange
{
    public double InputMetric { get; set; } = 100.0;
    public double Saturation
    {
        get
        {
            return InputMetric * SaturationScale + SaturationOffset; // ' use linear mapping to compute Saturation based some metric like composite value. SaturationOffest & SaturationScale were perviously computed from the composit rating
        }
    }
    public double SaturationScale { get; set; } = 2.3;
    public double SaturationOffset { get; set; } = 93.3;
    public double Hue { get; set; } = 200.0;
    public double Luminance { get; set; } = 200.0;
    public double SaturationMax
    {
        get
        {
            return ymax;
        }
        set
        {
            ymax = value;
            recalc();
        }
    }
    public double SaturationMin
    {
        get
        {
            return ymin;
        }
        set
        {
            ymin = value;
            recalc();
        }

    }
    private double xmax;
    private double xmin;
    public double InputMetricMax
    {
        get
        {
            return xmax;
        }
        set
        {
            xmax = value;
            recalc();
        }
    }
    public double InputMetricMin
    {
        get { return xmin; }
        set
        {
            xmin = value;
            recalc();
        }
    }
    private double ymax;
    private double ymin;
    private void recalc()
    {
        SaturationScale = (SaturationMin + (xmax * SaturationMin / xmin - SaturationMax) / (1 - xmax / xmin)) / xmin;
        SaturationOffset = (SaturationMax - xmax * SaturationMin / xmin) / (1 - xmax / xmin);
    }
    public string ColorHexRGB
    {
        get
        {
            var c = ColorConverter.HlsToRgba(Hue / 255.0f, Saturation / 255.0f, Luminance / 255.0f);
            var result = String.Format("{0:X2}{1:X2}{2:X2}", c.R, c.G, c.B);
            return result;
        }
    }
}

