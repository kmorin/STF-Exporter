using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace STFExporter
{
    class Interpolation
    {
        private Interpolation() { }

        static public double linear(double x, double x0, double x1, double y0, double y1)
        {
            if ((x1-x0) == 0)
            {
                return (y0 + y1) / 2;
            }
            return y0 + (x - x0) * (y1 - y0) / (x1 - x0);
        }
    }
}
