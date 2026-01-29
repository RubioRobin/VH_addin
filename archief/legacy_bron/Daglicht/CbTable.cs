using System;
using System.Collections.Generic;
using System.Linq;

namespace VH_DaglichtPlugin
{
    public static class CbTable
    {
        private static readonly int[] ALPHAS = { 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32 };

        private static readonly List<BetaRange> BETA_RANGES = new List<BetaRange>
        {
            new BetaRange { Start = null, End = 0, Values = new Dictionary<int, double> { {20, 0.80}, {21, 0.79}, {22, 0.79}, {23, 0.78}, {24, 0.77}, {25, 0.77}, {26, 0.76}, {27, 0.75}, {28, 0.75}, {29, 0.74}, {30, 0.73}, {31, 0.73}, {32, 0.72} } },
            new BetaRange { Start = 0, End = 5, Values = new Dictionary<int, double> { {20, 0.80}, {21, 0.79}, {22, 0.79}, {23, 0.78}, {24, 0.77}, {25, 0.77}, {26, 0.76}, {27, 0.75}, {28, 0.75}, {29, 0.74}, {30, 0.73}, {31, 0.73}, {32, 0.72} } },
            new BetaRange { Start = 5, End = 10, Values = new Dictionary<int, double> { {20, 0.80}, {21, 0.79}, {22, 0.78}, {23, 0.78}, {24, 0.77}, {25, 0.76}, {26, 0.76}, {27, 0.75}, {28, 0.74}, {29, 0.74}, {30, 0.73}, {31, 0.72}, {32, 0.72} } },
            new BetaRange { Start = 10, End = 15, Values = new Dictionary<int, double> { {20, 0.79}, {21, 0.79}, {22, 0.78}, {23, 0.77}, {24, 0.77}, {25, 0.76}, {26, 0.76}, {27, 0.75}, {28, 0.75}, {29, 0.74}, {30, 0.73}, {31, 0.72}, {32, 0.71} } },
            new BetaRange { Start = 15, End = 20, Values = new Dictionary<int, double> { {20, 0.79}, {21, 0.78}, {22, 0.77}, {23, 0.77}, {24, 0.76}, {25, 0.75}, {26, 0.75}, {27, 0.74}, {28, 0.73}, {29, 0.72}, {30, 0.72}, {31, 0.71}, {32, 0.70} } },
            new BetaRange { Start = 20, End = 25, Values = new Dictionary<int, double> { {20, 0.77}, {21, 0.76}, {22, 0.76}, {23, 0.75}, {24, 0.74}, {25, 0.73}, {26, 0.73}, {27, 0.72}, {28, 0.71}, {29, 0.70}, {30, 0.70}, {31, 0.69}, {32, 0.68} } },
            new BetaRange { Start = 25, End = 30, Values = new Dictionary<int, double> { {20, 0.76}, {21, 0.75}, {22, 0.74}, {23, 0.73}, {24, 0.72}, {25, 0.72}, {26, 0.71}, {27, 0.70}, {28, 0.69}, {29, 0.68}, {30, 0.68}, {31, 0.67}, {32, 0.66} } },
            new BetaRange { Start = 30, End = 35, Values = new Dictionary<int, double> { {20, 0.74}, {21, 0.73}, {22, 0.72}, {23, 0.71}, {24, 0.70}, {25, 0.69}, {26, 0.69}, {27, 0.68}, {28, 0.67}, {29, 0.66}, {30, 0.65}, {31, 0.64}, {32, 0.63} } },
            new BetaRange { Start = 35, End = 40, Values = new Dictionary<int, double> { {20, 0.72}, {21, 0.70}, {22, 0.70}, {23, 0.68}, {24, 0.68}, {25, 0.67}, {26, 0.66}, {27, 0.65}, {28, 0.64}, {29, 0.63}, {30, 0.62}, {31, 0.61}, {32, 0.60} } },
            new BetaRange { Start = 40, End = 45, Values = new Dictionary<int, double> { {20, 0.69}, {21, 0.68}, {22, 0.66}, {23, 0.65}, {24, 0.64}, {25, 0.63}, {26, 0.62}, {27, 0.61}, {28, 0.60}, {29, 0.59}, {30, 0.58}, {31, 0.57}, {32, 0.55} } },
            new BetaRange { Start = 45, End = 50, Values = new Dictionary<int, double> { {20, 0.65}, {21, 0.64}, {22, 0.63}, {23, 0.61}, {24, 0.60}, {25, 0.59}, {26, 0.58}, {27, 0.56}, {28, 0.55}, {29, 0.54}, {30, 0.53}, {31, 0.52}, {32, 0.50} } },
            new BetaRange { Start = 50, End = 55, Values = new Dictionary<int, double> { {20, 0.60}, {21, 0.59}, {22, 0.58}, {23, 0.56}, {24, 0.55}, {25, 0.54}, {26, 0.52}, {27, 0.51}, {28, 0.50}, {29, 0.49}, {30, 0.48}, {31, 0.47}, {32, 0.45} } },
            new BetaRange { Start = 55, End = 60, Values = new Dictionary<int, double> { {20, 0.53}, {21, 0.52}, {22, 0.50}, {23, 0.49}, {24, 0.47}, {25, 0.46}, {26, 0.44}, {27, 0.43}, {28, 0.41}, {29, 0.40}, {30, 0.38}, {31, 0.36}, {32, 0.34} } },
            new BetaRange { Start = 60, End = 65, Values = new Dictionary<int, double> { {20, 0.45}, {21, 0.43}, {22, 0.42}, {23, 0.40}, {24, 0.39}, {25, 0.37}, {26, 0.36}, {27, 0.34}, {28, 0.33}, {29, 0.31}, {30, 0.29}, {31, 0.28}, {32, 0.26} } },
            new BetaRange { Start = 65, End = 70, Values = new Dictionary<int, double> { {20, 0.34}, {21, 0.32}, {22, 0.30}, {23, 0.28}, {24, 0.26}, {25, 0.24}, {26, 0.23}, {27, 0.21}, {28, 0.19}, {29, 0.17}, {30, 0.00}, {31, 0.00}, {32, 0.00} } },
            new BetaRange { Start = 70, End = 75, Values = new Dictionary<int, double> { {20, 0.23}, {21, 0.21}, {22, 0.19}, {23, 0.17}, {24, 0.15}, {25, 0.00}, {26, 0.00}, {27, 0.00}, {28, 0.00}, {29, 0.00}, {30, 0.00}, {31, 0.00}, {32, 0.00} } },
            new BetaRange { Start = 75, End = 90, Values = new Dictionary<int, double> { {20, 0.00}, {21, 0.00}, {22, 0.00}, {23, 0.00}, {24, 0.00}, {25, 0.00}, {26, 0.00}, {27, 0.00}, {28, 0.00}, {29, 0.00}, {30, 0.00}, {31, 0.00}, {32, 0.00} } }
        };

        public static double? GetCb(double alphaDeg, double betaDeg)
        {
            // Find beta row
            BetaRange row = null;
            foreach (var r in BETA_RANGES)
            {
                if (r.Start == null)
                {
                    if (betaDeg <= r.End)
                    {
                        row = r;
                        break;
                    }
                }
                else
                {
                    if (betaDeg >= r.Start.Value && betaDeg <= r.End)
                    {
                        row = r;
                        break;
                    }
                }
            }

            if (row == null)
                row = BETA_RANGES.Last();

            // Clamp alpha
            double alpha = Math.Max(ALPHAS.First(), Math.Min(ALPHAS.Last(), alphaDeg));

            // Find surrounding alphas
            int a1 = ALPHAS.Where(a => a <= alpha).Max();
            int a2 = ALPHAS.Where(a => a >= alpha).Min();

            if (a1 == a2)
                return row.Values[a1];

            // Interpolate
            double t = (alpha - a1) / (double)(a2 - a1);
            return row.Values[a1] * (1.0 - t) + row.Values[a2] * t;
        }

        private class BetaRange
        {
            public double? Start { get; set; }
            public double End { get; set; }
            public Dictionary<int, double> Values { get; set; }
        }
    }
}
