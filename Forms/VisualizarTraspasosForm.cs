using AdminSERMAC.Services;
using ClosedXML.Excel;
using System.Data.SQLite;
using System.Diagnostics;

public class VisualizarTraspasosForm : Form
{
    private readonly SQLiteService _sqliteService;
    private DateTimePicker fechaDesdePicker;
    private DateTimePicker fechaHastaPicker;
    private ComboBox ubicacionComboBox;
    private DataGridView traspasosDataGridView;
    private Button buscarButton;
    private Button exportarExcelButton;

    public VisualizarTraspasosForm(SQLiteService sqliteService)
    {
        _sqliteService = sqliteService;
        InitializeComponents();
        ConfigureEvents();
        CargarTraspasos();
    }

    private void InitializeComponents()
    {
        this.Text = "Visualizar Traspasos";
        this.Size = new Size(1200, 600);
        this.StartPosition = FormStartPosition.CenterScreen;
        this.BackColor = Color.WhiteSmoke;

        fechaDesdePicker = new DateTimePicker
        {
            Location = new Point(20, 20),
            Format = DateTimePickerFormat.Short,
            Width = 120
        };

        fechaHastaPicker = new DateTimePicker
        {
            Location = new Point(200, 20),
            Format = DateTimePickerFormat.Short,
            Width = 120
        };

        ubicacionComboBox = new ComboBox
        {
            Location = new Point(380, 20),
            Width = 200,
            DropDownStyle = ComboBoxStyle.DropDownList,
            Font = new Font("Segoe UI", 9F)
        };

        ubicacionComboBox.Items.AddRange(new string[] { "Todas", "Inventario Normal", "Cámara Congelados", "Local" });
        ubicacionComboBox.SelectedIndex = 0;

        buscarButton = new Button
        {
            Text = "Buscar",
            Location = new Point(600, 20),
            Width = 100,
            FlatStyle = FlatStyle.Flat,
            BackColor = Color.FromArgb(45, 66, 91),
            ForeColor = Color.White
        };

        exportarExcelButton = new Button
        {
            Text = "Exportar a Excel",
            Location = new Point(720, 20),
            Width = 120,
            FlatStyle = FlatStyle.Flat,
            BackColor = Color.FromArgb(45, 66, 91),
            ForeColor = Color.White
        };

        traspasosDataGridView = new DataGridView
        {
            Location = new Point(20, 60),
            Size = new Size(1140, 480),
            AllowUserToAddRows = false,
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
            ReadOnly = true,
            BackgroundColor = Color.White,
            BorderStyle = BorderStyle.None,
            RowHeadersVisible = false
        };

        ConfigurarColumnas();

        this.Controls.AddRange(new Control[] {
            fechaDesdePicker,
            fechaHastaPicker,
            ubicacionComboBox,
            buscarButton,
            exportarExcelButton,
            traspasosDataGridView
        });
    }

    private void ConfigurarColumnas()
    {
        traspasosDataGridView.Columns.Clear();
        traspasosDataGridView.Columns.AddRange(new DataGridViewColumn[]
        {
            new DataGridViewTextBoxColumn { Name = "Fecha", HeaderText = "Fecha", Width = 150 },
            new DataGridViewTextBoxColumn { Name = "Origen", HeaderText = "Origen", Width = 150 },
            new DataGridViewTextBoxColumn { Name = "Destino", HeaderText = "Destino", Width = 150 },
            new DataGridViewTextBoxColumn { Name = "Codigo", HeaderText = "Código", Width = 100 },
            new DataGridViewTextBoxColumn { Name = "Descripcion", HeaderText = "Descripción", Width = 200 },
            new DataGridViewTextBoxColumn { Name = "Unidades", HeaderText = "Unidades", Width = 100 },
            new DataGridViewTextBoxColumn { Name = "Kilos", HeaderText = "Kilos", Width = 100 }
        });

        traspasosDataGridView.EnableHeadersVisualStyles = false;
        traspasosDataGridView.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(45, 66, 91);
        traspasosDataGridView.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
        traspasosDataGridView.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
        traspasosDataGridView.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(245, 245, 245);
    }

    private void ConfigureEvents()
    {
        buscarButton.Click += BuscarButton_Click;
        exportarExcelButton.Click += ExportarExcelButton_Click;
    }

    private void CargarTraspasos()
    {
        string query = @"
       SELECT 
           h.FechaTraspaso,
           h.UbicacionOrigen as Origen,
           h.UbicacionDestino as Destino,
           h.CodigoProducto,
           h.Descripcion,
           h.Unidades,
           h.Kilos
       FROM HistorialTraspasos h
       WHERE h.FechaTraspaso BETWEEN @fechaDesde AND @fechaHasta
       AND (@ubicacion = 'Todas' 
           OR h.UbicacionOrigen = @ubicacion 
           OR h.UbicacionDestino = @ubicacion)
       ORDER BY h.FechaTraspaso DESC";

        try
        {
            using (var connection = new SQLiteConnection(_sqliteService.connectionString))
            {
                connection.Open();
                using (var command = new SQLiteCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@fechaDesde", fechaDesdePicker.Value.ToString("yyyy-MM-dd"));
                    command.Parameters.AddWithValue("@fechaHasta", fechaHastaPicker.Value.ToString("yyyy-MM-dd 23:59:59"));
                    command.Parameters.AddWithValue("@ubicacion", ubicacionComboBox.SelectedItem.ToString());

                    traspasosDataGridView.Rows.Clear();
                    using (var reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            traspasosDataGridView.Rows.Add(
                                DateTime.Parse(reader["FechaTraspaso"].ToString()).ToString("dd/MM/yyyy HH:mm"),
                                reader["Origen"],
                                reader["Destino"],
                                reader["CodigoProducto"],
                                reader["Descripcion"],
                                reader["Unidades"],
                                Convert.ToDouble(reader["Kilos"]).ToString("N2")
                            );
                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error al cargar traspasos: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void BuscarButton_Click(object sender, EventArgs e)
    {
        if (fechaDesdePicker.Value > fechaHastaPicker.Value)
        {
            MessageBox.Show("La fecha desde no puede ser mayor a la fecha hasta", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }
        CargarTraspasos();
    }

    private void ExportarExcelButton_Click(object sender, EventArgs e)
    {
        if (traspasosDataGridView.Rows.Count == 0)
        {
            MessageBox.Show("No hay datos para exportar", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        try
        {
            using (SaveFileDialog saveDialog = new SaveFileDialog())
            {
                saveDialog.Filter = "Excel files (*.xlsx)|*.xlsx";
                saveDialog.FileName = $"Traspasos_{DateTime.Now:yyyyMMdd}.xlsx";

                if (saveDialog.ShowDialog() == DialogResult.OK)
                {
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add("Traspasos");

                        // Encabezados
                        for (int i = 0; i < traspasosDataGridView.Columns.Count; i++)
                        {
                            worksheet.Cell(1, i + 1).Value = traspasosDataGridView.Columns[i].HeaderText;
                        }

                        // Datos
                        for (int i = 0; i < traspasosDataGridView.Rows.Count; i++)
                        {
                            for (int j = 0; j < traspasosDataGridView.Columns.Count; j++)
                            {
                                worksheet.Cell(i + 2, j + 1).Value = traspasosDataGridView.Rows[i].Cells[j].Value?.ToString();
                            }
                        }

                        worksheet.Columns().AdjustToContents();
                        workbook.SaveAs(saveDialog.FileName);
                        Process.Start(new ProcessStartInfo { FileName = saveDialog.FileName, UseShellExecute = true });
                    }
                }
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error al exportar: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}