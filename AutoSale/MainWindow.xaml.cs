using CsvHelper;
using CsvHelper.Configuration;
using Microsoft.ML;
using Microsoft.ML.Data;
using Microsoft.Win32;
using System.Data;
using System.Globalization;
using System.IO;
using System.Windows;

namespace AutoSale
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            //CÂU 1: CLASSIFICATION PROBLEM(3 điểm)
            //Sử dụng tập dữ liệu Auto Sales data.csv và thực hiện các yêu cầu sau:
            //1.1 Chuẩn bị và Kiểm tra Dữ liệu

            //Tải và kiểm tra tập dữ liệu.
            //Mô tả cấu trúc của tập dữ liệu và phân loại dữ liệu.
            //Thực hiện các bước làm sạch dữ liệu, chuẩn bị dữ liệu để phân tích. Đề xuất các phương pháp xử lý dữ liệu bị thiếu hoặc outliers(nếu có).
            //1.2 Phân tích phân loại khách hàng

            //Sử dụng Quartile và RFM Analysis để phân loại khách hàng thành các nhóm dựa trên Recency, Frequency, và Monetary.
            //Sử dụng thuật toán K - means Clustering để phân loại khách hàng.
            //Trình bày các kết quả phân loại khách hàng từ cả hai phương pháp. So sánh hai phương pháp RFM và K-means và đưa ra nhận xét về sự khác biệt và ưu điểm của mỗi phương pháp.
            //Dựa trên kết quả phân tích, đề xuất các chiến lược cụ thể để tối ưu hóa doanh thu.
        }

        private DataTable LoadCsv(string filePath)
        {
            var dataTable = new DataTable();
            using (var reader = new StreamReader(filePath))
            using (var csv = new CsvReader(reader, new CsvConfiguration(CultureInfo.InvariantCulture)))
            {
                using (var dr = new CsvDataReader(csv))
                {
                    dataTable.Load(dr);
                }
            }
            return dataTable;
        }

        private void BtnLoadData_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var openFileDialog = new OpenFileDialog
                {
                    Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*",
                    Title = "Select Auto Sales Data File"
                };

                if (openFileDialog.ShowDialog() == true)
                {
                    string filePath = openFileDialog.FileName;
                    var data = LoadCsv(filePath);
                    DataGridView.ItemsSource = data.DefaultView;
                }
            }
            catch (Exception ex)
            {
                LogError(ex);
                MessageBox.Show(ex.Message);
            }
        }

        private void BtnCleanData_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var data = ((DataView)DataGridView.ItemsSource)?.Table;

                if (data == null)
                {
                    MessageBox.Show("No data to clean.");
                    return;
                }

                foreach (DataColumn column in data.Columns)
                {
                    if (column.DataType == typeof(double) || column.DataType == typeof(int))
                    {
                        var avg = data.AsEnumerable()
                                      .Where(row => !row.IsNull(column))
                                      .Average(row => Convert.ToDouble(row[column]));
                        foreach (DataRow row in data.Rows)
                        {
                            if (row.IsNull(column))
                            {
                                row[column] = avg;
                            }
                        }

                        var values = data.AsEnumerable()
                                         .Where(row => !row.IsNull(column))
                                         .Select(row => Convert.ToDouble(row[column]))
                                         .ToList();

                        double mean = values.Average();
                        double stdDev = Math.Sqrt(values.Average(v => Math.Pow(v - mean, 2)));

                        foreach (DataRow row in data.Rows)
                        {
                            if (!row.IsNull(column))
                            {
                                double zScore = Math.Abs((Convert.ToDouble(row[column]) - mean) / stdDev);
                                if (zScore > 3)
                                {
                                    row[column] = mean;
                                }
                            }
                        }
                    }
                    else if (column.DataType == typeof(string))
                    {
                        foreach (DataRow row in data.Rows)
                        {
                            if (row.IsNull(column))
                            {
                                row[column] = "Unknown";
                            }
                        }
                    }
                }

                DataGridView.ItemsSource = data.DefaultView;
                MessageBox.Show("Data cleaning complete.");
            }
            catch (Exception ex)
            {
                LogError(ex);
                MessageBox.Show(ex.Message);
            }
        }

        private string DescribeData(DataTable data)
        {
            int rowCount = data.Rows.Count;
            int columnCount = data.Columns.Count;
            var columnInfo = string.Join("\n", data.Columns
                                                  .Cast<DataColumn>()
                                                  .Select(col => $"{col.ColumnName} ({col.DataType.Name})"));
            return $"Dataset contains {rowCount} rows and {columnCount} columns.\n\nColumns:\n{columnInfo}";
        }

        private void BtnAnalyzeData_Click(object sender, RoutedEventArgs e)
        {
            var data = ((DataView)DataGridView.ItemsSource)?.Table;
            if (data == null)
            {
                MessageBox.Show("No data loaded.");
                return;
            }

            string info = DescribeData(data);
            MessageBox.Show(info, "Dataset Information");
        }

        private DataTable CalculateRFM(DataTable data)
        {
            var today = DateTime.Today;

            if (!data.Columns.Contains("Recency"))
            {
                var recencyColumn = "Recency";
                data.Columns.Add(recencyColumn, typeof(int));

                foreach (DataRow row in data.Rows)
                {
                    string orderDateStr = row["ORDERDATE"].ToString();

                    try
                    {
                        var orderDate = DateTime.ParseExact(orderDateStr, "dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture);
                        row[recencyColumn] = (today - orderDate).Days;
                    }
                    catch (FormatException)
                    {
                        row[recencyColumn] = 9999;
                    }
                }
            }

            if (!data.Columns.Contains("Frequency"))
            {
                var frequencyColumn = "Frequency";
                data.Columns.Add(frequencyColumn, typeof(int));
                var frequencyData = data.AsEnumerable()
                                        .GroupBy(row => row["CUSTOMERNAME"].ToString())
                                        .ToDictionary(
                                            group => group.Key,
                                            group => group.Count()
                                        );
                foreach (DataRow row in data.Rows)
                {
                    row[frequencyColumn] = frequencyData[row["CUSTOMERNAME"].ToString()];
                }
            }

            if (!data.Columns.Contains("Monetary"))
            {
                var monetaryColumn = "Monetary";
                data.Columns.Add(monetaryColumn, typeof(double));
                var monetaryData = data.AsEnumerable()
                                       .GroupBy(row => row["CUSTOMERNAME"].ToString())
                                       .ToDictionary(
                                           group => group.Key,
                                           group => group.Sum(r => Convert.ToDouble(r["SALES"]))
                                       );
                foreach (DataRow row in data.Rows)
                {
                    row[monetaryColumn] = monetaryData[row["CUSTOMERNAME"].ToString()];
                }
            }

            return data;
        }

        private void AssignQuartileScores(DataTable data)
        {
            string[] metrics = { "Recency", "Frequency", "Monetary" };

            foreach (var metric in metrics)
            {
                var values = data.AsEnumerable()
                                 .Select(row => Convert.ToDouble(row[metric]))
                                 .OrderBy(value => value)
                                 .ToList();

                double q1 = values[(int)(values.Count * 0.25)];
                double q2 = values[(int)(values.Count * 0.5)];
                double q3 = values[(int)(values.Count * 0.75)];

                var scoreColumn = metric + "_Score";
                data.Columns.Add(scoreColumn, typeof(int));
                foreach (DataRow row in data.Rows)
                {
                    double value = Convert.ToDouble(row[metric]);
                    if (value <= q1)
                        row[scoreColumn] = 1;
                    else if (value <= q2)
                        row[scoreColumn] = 2;
                    else if (value <= q3)
                        row[scoreColumn] = 3;
                    else
                        row[scoreColumn] = 4;
                }
            }
        }

        private void BtnAnalyzeRFM_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var data = ((DataView)DataGridView.ItemsSource)?.Table;

                if (data == null)
                {
                    MessageBox.Show("No data loaded.");
                    return;
                }

                data = CalculateRFM(data);
                AssignQuartileScores(data);
                var rfmColumn = "RFM_Segment";
                data.Columns.Add(rfmColumn, typeof(string));
                foreach (DataRow row in data.Rows)
                {
                    var rScore = Convert.ToInt32(row["Recency_Score"]);
                    var fScore = Convert.ToInt32(row["Frequency_Score"]);
                    var mScore = Convert.ToInt32(row["Monetary_Score"]);
                    row[rfmColumn] = $"{rScore}{fScore}{mScore}";
                }

                DataGridView.ItemsSource = data.DefaultView;
                MessageBox.Show("RFM Analysis Complete.");
            }
            catch (Exception ex)
            {
                LogError(ex);
                MessageBox.Show(ex.Message);
            }
        }

        private void PerformKMeansClustering(DataTable data, int numberOfClusters)
        {
            var rfmData = data.AsEnumerable()
                              .Select(row => new CustomerRFM
                              {
                                  Recency = Convert.ToSingle(row["Recency"]),
                                  Frequency = Convert.ToSingle(row["Frequency"]),
                                  Monetary = Convert.ToSingle(row["Monetary"])
                              })
                              .ToList();

            var mlContext = new MLContext();
            var trainingData = mlContext.Data.LoadFromEnumerable(rfmData);
            var pipeline = mlContext.Clustering.Trainers.KMeans(
                featureColumnName: "Features",
                numberOfClusters: numberOfClusters);
            var dataProcessPipeline = mlContext.Transforms.Concatenate("Features", new[] { "Recency", "Frequency", "Monetary" });
            var model = dataProcessPipeline.Append(pipeline).Fit(trainingData);
            var predictions = model.Transform(trainingData);
            var clusteredData = mlContext.Data.CreateEnumerable<ClusterPrediction>(predictions, reuseRowObject: false).ToList();
            var clusterColumn = "Cluster";

            if (!data.Columns.Contains(clusterColumn))
                data.Columns.Add(clusterColumn, typeof(int));

            for (int i = 0; i < data.Rows.Count; i++)
            {
                data.Rows[i][clusterColumn] = clusteredData[i].PredictedClusterId;
            }

            DataGridView.ItemsSource = data.DefaultView;
            MessageBox.Show("K-means Clustering Complete!");
        }


        public class ClusterPrediction
        {
            [ColumnName("PredictedLabel")]
            public uint PredictedClusterId { get; set; }
        }

        private void BtnKMeans_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var data = ((DataView)DataGridView.ItemsSource)?.Table;

                if (data == null)
                {
                    MessageBox.Show("No data loaded.");
                    return;
                }

                if (!data.Columns.Contains("Recency") || !data.Columns.Contains("Frequency") || !data.Columns.Contains("Monetary"))
                {
                    data = CalculateRFM(data);
                }

                PerformKMeansClustering(data, numberOfClusters: 3);
            }
            catch (Exception ex)
            {
                LogError(ex);
                MessageBox.Show(ex.Message);
            }
        }

        private void LogError(Exception ex)
        {
            try
            {
                string logFileName = $"log_{DateTime.Now:ddMMyyyy}.txt";
                string logFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, logFileName);
                string logContent = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {ex.Message}\n{ex.StackTrace}\n";

                File.AppendAllText(logFilePath, logContent);
            }
            catch
            {
            }
        }
    }
}