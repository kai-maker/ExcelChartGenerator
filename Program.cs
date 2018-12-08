using System;
using System.IO;
using System.Text;
using System.Linq;
using System.Threading;
using System.Globalization;
using System.Collections.Generic;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using System.Numerics;
using MathNet.Numerics.IntegralTransforms;
using MathNet.Numerics.Statistics;

namespace ChartGeneratorFor_SeitaiJyouhoKougakuZikken
{
    class statisticData
    {
        public double Ave, Median, PSD;
    }
    class ExperimentTrial
    {
        public string trialNumber;
        public FileInfo inputCSVFile;
        public FileInfo outputExcelFile;
        public ExcelPackage excelPackage;
        public statisticData x, y;
    }
    class Experiment
    {
        public string experimentName;
        public List<ExperimentTrial> experimentTrials;
        public FileInfo outputAvelageExcelFile;
    }
    class Student
    {
        public string studentId;
        public List<Experiment> experiments;
    }

    class Program
    {
        static float chartValRange = 6;
        static int chartSize = 600;

        static void Main(string[] args)
        {
            DirectoryInfo inputDirInfo;
            while (true)
            {
                Console.WriteLine("実験結果csvファイルの親ディレクトリのパスを入力");
                Console.Write(">>");
                string inputDirPath = Console.ReadLine();
                inputDirInfo = new DirectoryInfo(inputDirPath.Trim());
                if (!inputDirInfo.Exists)
                {
                    Console.WriteLine("{0}が見つかりませんでした", inputDirPath);
                    continue;
                }
                break;
            }
            IEnumerable<FileInfo> files = inputDirInfo.GetFiles();
            IEnumerable<FileInfo> csvFilesSorted;
            csvFilesSorted = files.Where(x => x.Extension == ".csv")
                                  .OrderBy(x => x.Name);

            var students = new List<Student>();
            foreach (var f in csvFilesSorted)
            {
                if (students.Count == 0 || f.Name.Split('_')[0] != students.Last().studentId)
                {
                    students.Add(new Student());
                    students.Last().studentId = f.Name.Split('_')[0];
                    students.Last().experiments = new List<Experiment>();
                }
                var experiments = students.Last().experiments;
                if (experiments.Count == 0 || f.Name.Split('_')[1] != experiments.Last().experimentName)
                {
                    experiments.Add(new Experiment());
                    experiments.Last().experimentName = f.Name.Split('_')[1];
                    experiments.Last().experimentTrials = new List<ExperimentTrial>();
                }
                var experimentTrials = students.Last().experiments.Last().experimentTrials;
                if (experimentTrials.Count == 0 || f.Name.Split('_')[2] != experimentTrials.Last().trialNumber)
                {
                    experimentTrials.Add(new ExperimentTrial());
                    experimentTrials.Last().trialNumber = f.Name.Split('_')[2].Split('.')[0];
                    experimentTrials.Last().inputCSVFile = f;
                }
            }

            foreach (var student in students)
            {
                Console.WriteLine("学籍番号: " + student.studentId);
                foreach (var experiment in student.experiments)
                {
                    Console.WriteLine("\t実験名: " + experiment.experimentName);
                    foreach (var trial in experiment.experimentTrials)
                    {
                        Console.Write("\t\t" + trial.trialNumber + "回目: ");
                        Console.WriteLine(trial.inputCSVFile.FullName);
                    }
                }
            }
            Console.WriteLine("散布図の値の範囲[mm]を設定(推奨4~8)");
            Console.Write(">>");
            chartValRange = Convert.ToSingle(Console.ReadLine());
            Console.WriteLine("散布図の辺の大きさ[px]を設定(推奨400~1000)");
            Console.Write(">>");
            chartSize = Convert.ToInt32(Console.ReadLine());

            ExcelTextFormat format = new ExcelTextFormat();
            {
                format.Delimiter = ',';
                format.Culture = new CultureInfo(Thread.CurrentThread.CurrentCulture.ToString());
                format.Culture.DateTimeFormat.ShortDatePattern = "dd-mm-yyyy";
                format.Encoding = new UTF8Encoding();
            }

            foreach (var student in students)
            {
                foreach (var experiment in student.experiments)
                {
                    foreach (var trial in experiment.experimentTrials)
                    {
                        trial.outputExcelFile = new FileInfo(trial.inputCSVFile.DirectoryName +
                        "/Generated/" + trial.inputCSVFile.Name.Split('.')[0] + ".xlsx");
                        if (!Directory.Exists(trial.outputExcelFile.DirectoryName))
                        {
                            Directory.CreateDirectory(trial.outputExcelFile.DirectoryName);
                        }
                        var excelPackage = new ExcelPackage();
                        {
                            ExcelWorksheet dataSheet = excelPackage.Workbook.Worksheets.Add("dataSheet");
                            dataSheet.Cells["A1"].LoadFromText(trial.inputCSVFile, format);
                            dataSheet.Cells[1, 1].Value = "時間";
                            dataSheet.Cells[1, 2].Value = "x座標";
                            dataSheet.Cells[1, 3].Value = "y座標";
                            var xList = new List<double>();
                            var yList = new List<double>();
                            using (ExcelRange rx = dataSheet.Cells[2, 2, dataSheet.Dimension.End.Row, 2])
                            using (ExcelRange ry = dataSheet.Cells[2, 3, dataSheet.Dimension.End.Row, 3])
                            {
                                foreach (var i in rx)
                                {
                                    if (i.Value == null)
                                        break;
                                    xList.Add((double)i.Value);
                                }
                                foreach (var i in ry)
                                {
                                    if (i.Value == null)
                                        break;
                                    yList.Add((double)i.Value);
                                }

                                var x = new statisticData
                                {
                                    Ave = xList.Mean(),
                                    Median = xList.Median(),
                                    PSD = xList.PopulationStandardDeviation()
                                };
                                var y = new statisticData
                                {
                                    Ave = yList.Mean(),
                                    Median = yList.Median(),
                                    PSD = yList.PopulationStandardDeviation()
                                };
                                trial.x = x;
                                trial.y = y;
                            }
                            Console.WriteLine("実験データのコピー: " + trial.outputExcelFile.Name);
                            Console.WriteLine("\t (x座標) 平均: {0:f3}, 標準偏差: {1:f3}, 中央値{2:f3}", trial.x.Ave, trial.x.PSD, trial.x.Median);
                            Console.WriteLine("\t (y座標) 平均: {0:f3}, 標準偏差: {1:f3}, 中央値{2:f3}", trial.y.Ave, trial.y.PSD, trial.y.Median);
                        }
                        trial.excelPackage = excelPackage;
                    }
                }
            }

            ExcelWorksheet chartTemplateWorksheet;
            {
                var chartTemplateFile = new FileInfo("ChartTemplate.xlsx");
                var chartTemplatePackage = new ExcelPackage(chartTemplateFile);
                chartTemplateWorksheet = chartTemplatePackage.Workbook.Worksheets["chartSheet"];
            }

            foreach (var student in students)
            {
                foreach (var experiment in student.experiments)
                {
                    foreach (var trial in experiment.experimentTrials)
                    {
                        var excelPackage = trial.excelPackage;
                        {
                            ExcelWorksheet chartSheet = excelPackage.Workbook.Worksheets.Add("chartSheet");
                            Console.WriteLine("散布図の生成: " + trial.outputExcelFile.Name);
                            var chart = chartSheet.Drawings.AddChart("散布図", eChartType.XYScatter) as ExcelScatterChart;
                            {
                                chart.XAxis.Title.Text = "x軸[mm]";
                                chart.YAxis.Title.Text = "y軸[mm]";
                                chart.SetPosition(1, 0, 1, 0);
                                chart.SetSize(chartSize, chartSize);
                                chart.RoundedCorners = false;
                            }
                            var dataSheet = excelPackage.Workbook.Worksheets["dataSheet"];
                            using (ExcelRange rx = dataSheet.Cells[2, 2, dataSheet.Dimension.End.Row, 2])
                            using (ExcelRange ry = dataSheet.Cells[2, 3, dataSheet.Dimension.End.Row, 3])
                            {
                                var serie = (ExcelScatterChartSerie)chart.Series.Add(ry.FullAddress, rx.FullAddress);
                                serie.Marker = eMarkerStyle.Circle;
                            }
                            {
                                chart.XAxis.MinValue = Math.Round(trial.x.Median / 0.5) * 0.5 - chartValRange / 2;
                                chart.XAxis.MaxValue = Math.Round(trial.x.Median / 0.5) * 0.5 + chartValRange / 2;
                                chart.YAxis.MinValue = Math.Round(trial.y.Median / 0.5) * 0.5 - chartValRange / 2;
                                chart.YAxis.MaxValue = Math.Round(trial.y.Median / 0.5) * 0.5 + chartValRange / 2;

                                chart.XAxis.MinorTickMark = eAxisTickMark.None;
                                chart.YAxis.MinorTickMark = eAxisTickMark.None;

                                chart.XAxis.CrossesAt = chart.YAxis.MinValue;
                                chart.YAxis.CrossesAt = chart.XAxis.MinValue;

                                chart.XAxis.Title.Font.Size = 30;
                                chart.YAxis.Title.Font.Size = 30;

                                chart.Legend.Remove();
                            }
                        }
                        trial.excelPackage = excelPackage;
                    }
                }
            }
            foreach (var student in students)
            {
                foreach (var experiment in student.experiments)
                {
                    experiment.outputAvelageExcelFile = new FileInfo(experiment.experimentTrials.First().inputCSVFile.DirectoryName +
                        "/Generated/" + student.studentId + "_" + experiment.experimentName + "AverageChart.xlsx");
                    var AverageExcel = new ExcelPackage();
                    var dataSheet = AverageExcel.Workbook.Worksheets.Add("dataSheet");
                    {
                        dataSheet.Cells[1, 2].Value = "x平均";
                        dataSheet.Cells[1, 3].Value = "y平均";
                        dataSheet.Cells[1, 4].Value = "x標準偏差";
                        dataSheet.Cells[1, 5].Value = "y標準偏差";
                        var count = 2;
                        foreach (var trial in experiment.experimentTrials)
                        {
                            dataSheet.Cells[count, 1].Value = (count - 1) + "回目";
                            dataSheet.Cells[count, 2].Value = trial.x.Ave;
                            dataSheet.Cells[count, 3].Value = trial.y.Ave;
                            dataSheet.Cells[count, 4].Value = trial.x.PSD;
                            dataSheet.Cells[count, 5].Value = trial.y.PSD;
                            count++;
                        }
                    }
                    AverageExcel.SaveAs(experiment.outputAvelageExcelFile);
                    Console.WriteLine("平均、標準偏差ファイルの生成、保存: " + experiment.outputAvelageExcelFile.Name);
                }
            }
            foreach (var student in students)
            {
                foreach (var experiment in student.experiments)
                {
                    foreach (var trial in experiment.experimentTrials)
                    {
                        ExcelWorksheet fourierSheet = trial.excelPackage.Workbook.Worksheets.Add("fourierSheet");
                        ExcelWorksheet dataSheet = trial.excelPackage.Workbook.Worksheets["dataSheet"];
                        {
                            var xArray = new Complex[4096];
                            var yArray = new Complex[4096];
                            using (ExcelRange rx = dataSheet.Cells[2, 2, dataSheet.Dimension.End.Row, 2])
                            using (ExcelRange ry = dataSheet.Cells[2, 3, dataSheet.Dimension.End.Row, 3])
                            {
                                int count = 0;
                                foreach (var i in rx)
                                {
                                    if (i.Value == null || count == 4096)
                                        break;
                                    xArray[count] = new Complex((double)i.Value,0);
                                    count++;
                                }
                                Fourier.Forward(xArray, FourierOptions.Matlab);
                                count = 0;
                                foreach (var i in ry)
                                {
                                    if (i.Value == null || count == 4096)
                                        break;
                                    yArray[count] = new Complex((double)i.Value, 0);
                                    count++;
                                }
                                Fourier.Forward(yArray, FourierOptions.Matlab); 
                            }
                            fourierSheet.Cells[1, 2].Value = "x方向振幅[mm]";
                            fourierSheet.Cells[1, 4].Value = "y方向振幅[mm]";
                            for (int i = 0; i < 4096; i++)
                            {
                                fourierSheet.Cells[i + 2, 2].Value = Complex.Abs(xArray[i]);
                                fourierSheet.Cells[i + 2, 4].Value = Complex.Abs(yArray[i]);
                            }
                            double samplingFreqency = (double)90 / (double)(dataSheet.Dimension.End.Row - 1);
                            double frequency = (double)1 / (4096 * samplingFreqency);
                            fourierSheet.Cells[1, 1].Value = "周波数[Hz]";
                            fourierSheet.Cells[1, 3].Value = "周波数[Hz]";
                            for (int i = 0; i < 4096; i++)
                            {
                                fourierSheet.Cells[i + 2, 1].Value = frequency * (i + 1);
                                fourierSheet.Cells[i + 2, 3].Value = frequency * (i + 1);
                            }
                        }
                        Console.WriteLine("フーリエ変換の計算、保存: " + trial.outputExcelFile.Name);
                    }
                }
            }
            foreach (var student in students)
            {
                foreach (var experiment in student.experiments)
                {
                    foreach (var trial in experiment.experimentTrials)
                    {
                        trial.excelPackage.SaveAs(trial.outputExcelFile);
                        Console.WriteLine("散布図、フーリエ変換保存: " + trial.outputExcelFile.Name);
                    }
                }
            }
            Console.WriteLine("完了 ! [Enter]で閉じる");
            Console.ReadLine();
        }
    }
}