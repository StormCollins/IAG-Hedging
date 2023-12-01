using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using mni = MathNet.Numerics.Interpolation;
using System.Numerics;

namespace IAGHedging;

public static class MathUtils
{
    [ExcelFunction(
        Name = "IAG.Math_Interpolate",
        Description = "Performs linear, exponential, or flat interpolation on a range for a single given point.n",
        Category = "IAG: Mathematics")]
    public static object Interpolate(
    [ExcelArgument(Name = "X-values", Description = "Independent variable.")]
        object[,] xValuesRange,
    [ExcelArgument(Name = "Y-values", Description = "Dependent variable.")]
        object[,] yValuesRange,
    [ExcelArgument(Name = "Xi", Description = "Value for which to interpolate.")]
        double xi,
    [ExcelArgument(
            Name = "Method",
            Description = "Method of interpolation: 'linear', 'exponential', 'flat'")]
        string method,
    [ExcelArgument(
            Name = "Extrapolation Method",
            Description = "Method of extrapolation: 'flat'")]
        string extrapolationMethod)
    {
        mni.IInterpolation? interpolator = null;

        List<double> xValues = new();
        List<double> yValues = new();

        for (int i = 0; i < xValuesRange.GetLength(0); i++)
        {
            xValues.Add((double)xValuesRange[i, 0]);
            yValues.Add((double)yValuesRange[i, 0]);
        }

        if (xi > xValues.Max() && extrapolationMethod.ToUpper() == "FLAT")
        {
            return yValues[^1];
        }

        switch (method.ToUpper())
        {
            case "LINEAR":
                interpolator = mni.LinearSpline.Interpolate(xValues, yValues);
                return interpolator.Interpolate(xi);
            case "EXPONENTIAL":
                return ExponentialInterpolation();
            case "FLAT":
                interpolator = mni.StepInterpolation.Interpolate(xValues, yValues);
                return interpolator.Interpolate(xi);
            default:
                return "Error";
        }

        // Log-linear interpolation fails for negative y-values therefore we move to the complex plane here then 
        // back to real numbers.
        double ExponentialInterpolation()
        {
            int lowerXIndex = xValues.IndexOf(xValues.Where(x => x <= xi).Max());
            int upperXIndex = xValues.IndexOf(xValues.Where(x => x >= xi).Min());

            if (lowerXIndex == upperXIndex)
            {
                return yValues[lowerXIndex];
            }

            Complex xiComplex = xi;
            Complex x0Complex = xValues[lowerXIndex];
            Complex x1Complex = xValues[upperXIndex];
            Complex y0Complex = yValues[lowerXIndex];
            Complex y1Complex = yValues[upperXIndex];
            Complex yi =
                (Complex.Log(y1Complex) - Complex.Log(y0Complex)) / (x1Complex - x0Complex) * (xiComplex - x0Complex) +
                Complex.Log(y0Complex);

            Complex outputY = Complex.Exp(yi);
            return outputY.Real;
        }
    }
}