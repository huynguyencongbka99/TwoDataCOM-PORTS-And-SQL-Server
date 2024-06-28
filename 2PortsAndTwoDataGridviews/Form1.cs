
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.IO.Ports;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using ClosedXML.Excel;
using System.Data.SqlClient;



namespace _2PortsAndTwoDataGridviews
{
    public partial class Form1 : Form
    {
        private Thread _initializeThread1;
        private Thread _initializeThread2;
        private SqlConnection _sqlConnection;
        public Form1()
        {
            InitializeComponent();
            InitializeDatabase();
            StartInitializationThreads();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
        }

        private void InitializeDatabase()
        {
            try
            {
                string connectionString = "Server=HuyNguyen\\TONYNGUYEN;Database=READ_DATA_COM_PUSH_TO_SQLSERVER;User Id=sa;Password=Nchbka1999";
                _sqlConnection = new SqlConnection(connectionString);
                _sqlConnection.Open(); // Open connection to SQL Server
                MessageBox.Show("Connected to SQLServer!");
            }
            catch(Exception e)
            {
                MessageBox.Show("Failed to connect to SQLServer! " + e.Message);
                this.Close();
            }
        }

        private void StartInitializationThreads()
        {
            _initializeThread1 = new Thread(InitializeSerialPort1);
            _initializeThread2 = new Thread(InitializeSerialPort2);

            _initializeThread1.Start();
            _initializeThread2.Start();
        }

        private void InitializeSerialPort1()
        {
            _serialPort1 = new SerialPort("COM22", 9600, Parity.None, 8, StopBits.One);
            _serialPort1.ReadTimeout = 500;
            _serialPort1.WriteTimeout = 500;
            _serialPort1.DataReceived += new SerialDataReceivedEventHandler(DataReceivedHandler1);

            try
            {
                _serialPort1.Open();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error opening COM22: {ex.Message}");
            }
        }


        private void InitializeSerialPort2()
        {
            _serialPort2 = new SerialPort("COM24", 9600, Parity.None, 8, StopBits.One);
            _serialPort2.ReadTimeout = 500;
            _serialPort2.WriteTimeout = 500;
            _serialPort2.DataReceived += new SerialDataReceivedEventHandler(DataReceivedHandler2);

            try
            {
                _serialPort2.Open();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error opening COM23: {ex.Message}");
            }
        }

        private void DataReceivedHandler1(object sender, SerialDataReceivedEventArgs e)
        {
            try
            {
                string data = _serialPort1.ReadExisting();
                string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");

                this.BeginInvoke(new Action(() =>
                {
                    dataGridView1.Rows.Add(timestamp, data);
                    CheckAndRemoveDuplicates(dataGridView1, 1, 2);
                }));

                InsertDataIntoDatabase1("COM22", timestamp, data);
            }
            catch (TimeoutException) { }
            catch (Exception ex)
            {
                Console.WriteLine($"Error reading COM22: {ex.Message}");
            }
        }

        private void DataReceivedHandler2(object sender, SerialDataReceivedEventArgs e)
        {
            try
            {
                string data = _serialPort2.ReadExisting();
                string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");

                this.BeginInvoke(new Action(() =>
                {
                    dataGridView2.Rows.Add(timestamp, data);
                    CheckAndRemoveDuplicates(dataGridView1, 1, 2);
                }));

                

                InsertDataIntoDatabase2("COM24", timestamp, data);
            }
            catch (TimeoutException) { }
            catch (Exception ex)
            {
                Console.WriteLine($"Error reading COM23: {ex.Message}");
            }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (_serialPort1 != null && _serialPort1.IsOpen)
            {
                _serialPort1.Close();
            }

            if (_serialPort2 != null && _serialPort2.IsOpen)
            {
                _serialPort2.Close();
            }

            if (_initializeThread1 != null && _initializeThread1.IsAlive)
            {
                _initializeThread1.Join(); // Wait for the thread to finish
            }

            if (_initializeThread2 != null && _initializeThread2.IsAlive)
            {
                _initializeThread2.Join(); // Wait for the thread to finish
            }

            _sqlConnection.Close();
        }


        private void SaveDataToExcel(string filePath)
        {
            using (var workbook = new XLWorkbook())
            {
                // Add a new worksheet for DataGridView1
                var worksheet1 = workbook.Worksheets.Add("COM22 Data");
                for (int i = 0; i < dataGridView1.Columns.Count; i++)
                {
                    worksheet1.Cell(1, i + 1).Value = dataGridView1.Columns[i].HeaderText;
                }
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGridView1.Columns.Count; j++)
                    {
                        worksheet1.Cell(i + 2, j + 1).Value = dataGridView1.Rows[i].Cells[j].Value?.ToString();
                    }
                }

                // Add a new worksheet for DataGridView2
                var worksheet2 = workbook.Worksheets.Add("COM23 Data");
                for (int i = 0; i < dataGridView2.Columns.Count; i++)
                {
                    worksheet2.Cell(1, i + 1).Value = dataGridView2.Columns[i].HeaderText;
                }
                for (int i = 0; i < dataGridView2.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGridView2.Columns.Count; j++)
                    {
                        worksheet2.Cell(i + 2, j + 1).Value = dataGridView2.Rows[i].Cells[j].Value?.ToString();
                    }
                }

                // Save the Excel file
                workbook.SaveAs(filePath);
            }

            MessageBox.Show("Data saved to Excel successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void btnSaveToExcel_Click(object sender, EventArgs e)
        {
            SaveDataToExcel("Z:\\A_Works\\Study\\SelfLearningC#\\WindowFormApps\\NangCao\\2PortsAndTwoDataGridviews\\2PortsAndTwoDataGridviews\\SerialPortData.xlsx");
        }

        private void InsertDataIntoDatabase1(string portName, string timestamp, string data)
        {
            string query = "INSERT INTO DATA_COMPORT1 (COM, DATE, DATA) VALUES (@PortName, @Timestamp, @Data)";
            SqlCommand command = new SqlCommand(query, _sqlConnection);
            command.Parameters.AddWithValue("@PortName", portName);
            command.Parameters.AddWithValue("@Timestamp", timestamp);
            command.Parameters.AddWithValue("@Data", data);

            try
            {
                command.ExecuteNonQuery();
                Console.WriteLine($"Data inserted into database: Port={portName}, Timestamp={timestamp}, Data={data}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error inserting data into database: {ex.Message}");
            }
        }

        private void InsertDataIntoDatabase2(string portName, string timestamp, string data)
        {
            string query = "INSERT INTO DATA_COMPORT2 (COM, DATE, DATA) VALUES (@PortName, @Timestamp, @Data)";
            SqlCommand command = new SqlCommand(query, _sqlConnection);
            command.Parameters.AddWithValue("@PortName", portName);
            command.Parameters.AddWithValue("@Timestamp", timestamp);
            command.Parameters.AddWithValue("@Data", data);

            try
            {
                command.ExecuteNonQuery();
                Console.WriteLine($"Data inserted into database: Port={portName}, Timestamp={timestamp}, Data={data}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error inserting data into database: {ex.Message}");
            }
        }


        private void CheckAndRemoveDuplicates(DataGridView dataGridView, int columnIndex1, int columnIndex2)
        {
            HashSet<string> seenValues = new HashSet<string>();

            for (int i = dataGridView.Rows.Count - 1; i >= 0; i--)
            {
                var row = dataGridView.Rows[i];
                string value1 = row.Cells[columnIndex1].Value?.ToString();
                string value2 = row.Cells[columnIndex2].Value?.ToString();

                if (value1 != null && value2 != null)
                {
                    string combinedValues = value1 + "|" + value2;
                    if (seenValues.Contains(combinedValues))
                    {
                        //dataGridView.Rows.RemoveAt(i);
                        dataGridView.Rows.
                    }
                    else
                    {
                        seenValues.Add(combinedValues);
                    }
                }
            }
        }
    }
}
