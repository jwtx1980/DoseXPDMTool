$ErrorActionPreference = "Stop"

Add-Type -TypeDefinition @"
using System;
using System.Linq;

public static class Tg51PreviewMathVerifier
{
    public static double CalculatePpol(double[] high, double[] opposite)
    {
        double reference = Math.Abs(high.Average());
        return (Math.Abs(high.Average()) + Math.Abs(opposite.Average())) / (2.0 * reference);
    }

    public static double CalculatePion(double[] high, double[] low)
    {
        double highAbs = Math.Abs(high.Average());
        double lowAbs = Math.Abs(low.Average());
        double voltageRatioSquared = Math.Pow(300.0 / 150.0, 2);
        return (1.0 - voltageRatioSquared) / (highAbs / lowAbs - voltageRatioSquared);
    }

    public static double CalculatePtp(double temperatureC, double pressureMmHg)
    {
        return ((273.2 + temperatureC) / 295.2) * (760.0 / pressureMmHg);
    }

    public static double CalculateRoughOutputPerMu(double[] high, double ppol, double pion, double detectorCalibrationFactor, double temperatureC, double pressureMmHg, double deliveredMu, double prp, double doseToTissueCorrection, double qualityCorrectionFactor)
    {
        double ptp = CalculatePtp(temperatureC, pressureMmHg);
        double correction = ptp * ppol * pion * prp * doseToTissueCorrection * qualityCorrectionFactor;
        double doseGy = Math.Abs(high.Average()) * 1e-9 * detectorCalibrationFactor * correction;
        return doseGy * 100.0 / deliveredMu;
    }

    public static double CalculateReferenceOutputPerMu(double measuredPointOutputPerMu, double clinicalPddOrTmr)
    {
        double factor = clinicalPddOrTmr > 2.0 ? clinicalPddOrTmr / 100.0 : clinicalPddOrTmr;
        return measuredPointOutputPerMu / factor;
    }

    public static double CalculatePhotonKq(double measuredPdd10, double a, double b, double c)
    {
        return a + (b * measuredPdd10 / 1000.0) + (c * measuredPdd10 * measuredPdd10 / 100000.0);
    }
}
"@

function Assert-Close([string] $name, [double] $actual, [double] $expected, [double] $tolerance) {
    $delta = [Math]::Abs($actual - $expected)
    if ($delta -gt $tolerance) {
        throw "$name failed: actual=$actual expected=$expected tolerance=$tolerance"
    }

    Write-Host "$name OK: $actual"
}

$high = [double[]]@(-16.30, -16.22, -16.24)
$targetPion = 1.005
$targetPpol = 1.0008
$highMean = [Math]::Abs(($high | Measure-Object -Average).Average)
$lowMean = $highMean / (4.0 - (3.0 / $targetPion))
$oppositeMean = $highMean * ((2.0 * $targetPpol) - 1.0)
$low = [double[]]@(-$lowMean, -$lowMean, -$lowMean)
$opposite = [double[]]@($oppositeMean, $oppositeMean, $oppositeMean)

$ppol = [Tg51PreviewMathVerifier]::CalculatePpol($high, $opposite)
$pion = [Tg51PreviewMathVerifier]::CalculatePion($high, $low)
$ptp = [Tg51PreviewMathVerifier]::CalculatePtp(22.9, 770.3)
$output = [Tg51PreviewMathVerifier]::CalculateRoughOutputPerMu($high, $ppol, $pion, 48480000, 22.9, 770.3, 100, 1.0, 1.0, 1.0)
$sadReferenceOutput = [Tg51PreviewMathVerifier]::CalculateReferenceOutputPerMu($output, 0.7718)
$ssdReferenceOutput = [Tg51PreviewMathVerifier]::CalculateReferenceOutputPerMu($output, 77.18)
$electron20eOutput = [Tg51PreviewMathVerifier]::CalculateRoughOutputPerMu([double[]]@(-21.88), 1.0003, 1.0135, 48530000, 18.0, 754.6, 100, 1.0, 1.0, 0.904)
$electron20eReferenceOutput = [Tg51PreviewMathVerifier]::CalculateReferenceOutputPerMu($electron20eOutput, 96.05)
$photon6xKq = [Tg51PreviewMathVerifier]::CalculatePhotonKq(66.4, 0.9708, 1.972, -2.48)
$photon6xOutput = [Tg51PreviewMathVerifier]::CalculateRoughOutputPerMu([double[]]@(-13.515), 1.0008, 1.0035, 48290000, 22.0, 760.0, 100, 1.0, 1.0, $photon6xKq)
$photon6xReferenceOutput = [Tg51PreviewMathVerifier]::CalculateReferenceOutputPerMu($photon6xOutput, 66.4)

Assert-Close "Ppol" $ppol 1.0008 0.0000001
Assert-Close "Pion" $pion 1.005 0.0000001
Assert-Close "Ptp" $ptp 0.9896366002 0.0000001
Assert-Close "RoughOutput_cGyPerMU" $output 0.7843215728 0.0000001
Assert-Close "SADReferenceOutput_cGyPerMU" $sadReferenceOutput 1.0162238570 0.0000001
Assert-Close "SSDPercentReferenceOutput_cGyPerMU" $ssdReferenceOutput 1.0162238570 0.0000001
Assert-Close "Electron20EPointWithKecal_cGyPerMU" $electron20eOutput 0.9668339013 0.0000001
Assert-Close "Electron20EReferenceWithKecal_cGyPerMU" $electron20eReferenceOutput 1.0065943792 0.0000001
Assert-Close "Photon6XKq" $photon6xKq 0.9923985920 0.0000001
Write-Host "Photon6XExample_cGyPerMU: point=$photon6xOutput reference=$photon6xReferenceOutput"
