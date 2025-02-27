﻿using System;
using System.Configuration;
using System.Drawing;
using System.Windows.Forms;
using AdminSERMAC.Core.Interfaces;
using AdminSERMAC.Services;
using AdminSERMAC.Services.Database;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Console;



namespace AdminSERMAC.Forms
{
    public class InventarioForm : Form
    {
        private readonly ILogger<InventarioForm> _logger;
        private readonly SQLiteService _sqliteService;
        private readonly IInventarioDatabaseService _inventarioDatabaseService;
        private readonly IProductoDatabaseService _productoDatabaseService;

        private Button comprarProductosButton;
        private Button visualizarCongeladosButton;
        private Button visualizarInventarioButton;
        private Button crearProductoButton;
        private Button traspasoLocalButton;
        private Button modificarProductoButton;
        private Button visualizarTraspasosButton;
        private Panel mainPanel;
        private Label titleLabel;


        public InventarioForm(
        ILogger<InventarioForm> logger,
        SQLiteService sqliteService,
        IInventarioDatabaseService inventarioDatabaseService,
        IProductoDatabaseService productoDatabaseService)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _sqliteService = sqliteService ?? throw new ArgumentNullException(nameof(sqliteService));
            _inventarioDatabaseService = inventarioDatabaseService ?? throw new ArgumentNullException(nameof(inventarioDatabaseService));
            _productoDatabaseService = productoDatabaseService ?? throw new ArgumentNullException(nameof(productoDatabaseService));

            InitializeComponents();
        }

        private void InitializeComponents()
        {
            // Configuración del formulario
            this.Text = "Gestión de Inventario - SERMAC";
            this.Size = new Size(800, 750); // Aumentado para acomodar todos los botones
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.BackColor = Color.WhiteSmoke;

            // Panel principal
            mainPanel = new Panel
            {
                Dock = DockStyle.None,
                Size = new Size(700, 650),
                Location = new Point((this.ClientSize.Width - 700) / 2, 20),
                Padding = new Padding(20),
                BackColor = Color.White
            };

            // Título
            titleLabel = new Label
            {
                Text = "Gestión de Inventario",
                Font = new Font("Segoe UI", 24, FontStyle.Bold),
                AutoSize = true,
                Location = new Point(50, 30),
                ForeColor = Color.FromArgb(45, 66, 91)
            };

            // Botones
            comprarProductosButton = CreateMenuButton("Comprar Productos", 100);
            visualizarInventarioButton = CreateMenuButton("Visualizar Inventario", 160);
            crearProductoButton = CreateMenuButton("Crear Producto", 220);
            modificarProductoButton = CreateMenuButton("Modificar Producto", 280);
            traspasoLocalButton = CreateMenuButton("Traspaso a Local", 340);
            visualizarCongeladosButton = CreateMenuButton("Visualizar Cámara de Congelados", 400);
            visualizarTraspasosButton = CreateMenuButton("Visualizar Traspasos", 460);




            // Configurar eventos para los botones
            comprarProductosButton.Click += ComprarProductosButton_Click;
            visualizarInventarioButton.Click += VisualizarInventarioButton_Click;
            crearProductoButton.Click += CrearProductoButton_Click;
            traspasoLocalButton.Click += TraspasoLocalButton_Click;
            modificarProductoButton.Click += ModificarProductoButton_Click;
            visualizarCongeladosButton.Click += VisualizarCongeladosButton_Click;
            visualizarTraspasosButton.Click += VisualizarTraspasosButton_Click;



            mainPanel.Controls.Add(visualizarTraspasosButton);

            // Agregar controles al panel
            mainPanel.Controls.AddRange(new Control[] {
            titleLabel,
            comprarProductosButton,
            visualizarInventarioButton,
            crearProductoButton,
            modificarProductoButton,
            traspasoLocalButton,
            visualizarCongeladosButton,
            visualizarTraspasosButton
        });

            // Agregar panel al formulario
            this.Controls.Add(mainPanel);

            AjustarEstilosBotones();
        }

        private Button CreateMenuButton(string text, int top)
        {
            return new Button
            {
                Text = text,
                Top = top,
                Left = 50,
                Width = 600,
                Height = 45,
                Font = new Font("Segoe UI", 11, FontStyle.Regular),
                TextAlign = ContentAlignment.MiddleCenter,
                Cursor = Cursors.Hand,
                FlatStyle = FlatStyle.Flat
            };
        }

        private void AjustarEstilosBotones()
        {
            var botones = new[] {
            comprarProductosButton,
            visualizarInventarioButton,
            crearProductoButton,
            modificarProductoButton,
            traspasoLocalButton,
            visualizarCongeladosButton,
            visualizarTraspasosButton
        };

            foreach (var boton in botones)
            {
                boton.FlatStyle = FlatStyle.Flat;
                boton.FlatAppearance.BorderSize = 1;
                boton.BackColor = Color.FromArgb(45, 66, 91);
                boton.ForeColor = Color.White;

                // Eventos para hover effect
                boton.MouseEnter += (s, e) => {
                    Button btn = (Button)s;
                    btn.BackColor = Color.FromArgb(55, 76, 101);
                };

                boton.MouseLeave += (s, e) => {
                    Button btn = (Button)s;
                    btn.BackColor = Color.FromArgb(45, 66, 91);
                };
            }
        }

        private void VisualizarCongeladosButton_Click(object sender, EventArgs e)
        {
            try
            {
                var logger = LoggerFactory.Create(builder => builder.AddConsole())
                    .CreateLogger<VisualizarInventarioForm>();
                var visualizarCongeladosForm = new VisualizarInventarioForm(_sqliteService, logger, true); // Agregamos parámetro
                visualizarCongeladosForm.Text = "Visualizar Cámara de Congelados";
                visualizarCongeladosForm.Show();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al abrir el visualizador de congelados");
                MessageBox.Show($"Error al abrir el visualizador de congelados: {ex.Message}",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void VisualizarTraspasosButton_Click(object sender, EventArgs e)
        {
            try
            {
                var visualizarTraspasosForm = new VisualizarTraspasosForm(_sqliteService);
                visualizarTraspasosForm.Show();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al abrir el visualizador de traspasos");
                MessageBox.Show("Error al abrir el visualizador de traspasos", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void ModificarProductoButton_Click(object sender, EventArgs e)
        {
            try
            {
                var inputForm = new Form()
                {
                    Width = 300,
                    Height = 150,
                    FormBorderStyle = FormBorderStyle.FixedDialog,
                    Text = "Modificar Producto",
                    StartPosition = FormStartPosition.CenterParent
                };

                var textBox = new TextBox() { Left = 50, Top = 20, Width = 200 };
                var label = new Label() { Left = 50, Top = 5, Text = "Código del producto:" };
                var buttonOk = new Button() { Text = "Aceptar", Left = 50, Width = 100, Top = 50, DialogResult = DialogResult.OK };
                var buttonCancel = new Button() { Text = "Cancelar", Left = 150, Width = 100, Top = 50, DialogResult = DialogResult.Cancel };

                buttonOk.Click += (sender, e) => { inputForm.Close(); };
                buttonCancel.Click += (sender, e) => { inputForm.Close(); };

                inputForm.Controls.Add(label);
                inputForm.Controls.Add(textBox);
                inputForm.Controls.Add(buttonOk);
                inputForm.Controls.Add(buttonCancel);

                if (inputForm.ShowDialog() == DialogResult.OK && !string.IsNullOrEmpty(textBox.Text))
                {
                    // Crear un nuevo logger específico para SQLiteService
                    var sqliteLogger = LoggerFactory.Create(builder => builder.AddConsole())
                                                  .CreateLogger<SQLiteService>();

                    var modificarForm = new ModificarProductoForm(
                        textBox.Text,
                        sqliteLogger,  // Usar el nuevo logger
                        _productoDatabaseService,
                        _inventarioDatabaseService);
                    modificarForm.ShowDialog();
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al abrir el formulario de modificación de productos");
                MessageBox.Show($"Error al abrir el formulario de modificación de productos: {ex.Message}",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        private void ComprarProductosButton_Click(object sender, EventArgs e)
        {
            try
            {
                var sqliteLogger = LoggerFactory.Create(builder => builder.AddConsole())
                    .CreateLogger<SQLiteService>();

                // Crear los servicios necesarios con los loggers correctos
                var proveedorDatabaseService = new ProveedorDatabaseService(
                    LoggerFactory.Create(builder => builder.AddConsole()).CreateLogger<ProveedorDatabaseService>(),
                    _sqliteService.connectionString
                );

                var comprasDatabaseService = new ComprasDatabaseService(
                    LoggerFactory.Create(builder => builder.AddConsole()).CreateLogger<ComprasDatabaseService>(),
                    _sqliteService.connectionString
                );

                var compraForm = new ComprarInventarioForm(
                    sqliteLogger,
                    _productoDatabaseService,
                    _inventarioDatabaseService,
                    proveedorDatabaseService,  // IProveedorService
                    comprasDatabaseService     // IComprasDatabaseService
                );

                compraForm.Show();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al abrir el formulario de compra de productos");
                MessageBox.Show($"Error al abrir el formulario de compra de productos: {ex.Message}",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        private void VisualizarInventarioButton_Click(object sender, EventArgs e)
        {
            try
            {
                var logger = LoggerFactory.Create(builder => builder.AddConsole())
                    .CreateLogger<VisualizarInventarioForm>();
                var visualizarInventarioForm = new VisualizarInventarioForm(_sqliteService, logger);
                visualizarInventarioForm.Show();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al abrir el visualizador de inventario");
                MessageBox.Show($"Error al abrir el visualizador de inventario: {ex.Message}",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void CrearProductoButton_Click(object sender, EventArgs e)
        {
            try
            {
                var logger = LoggerFactory.Create(builder => builder.AddConsole())
                    .CreateLogger<CrearProductoForm>();
                using (var crearProductoForm = new CrearProductoForm(logger))
                {
                    if (crearProductoForm.ShowDialog() == DialogResult.OK)
                    {
                        MessageBox.Show("Producto creado exitosamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al abrir el formulario de creación de productos");
                MessageBox.Show($"Error al abrir el formulario de creación de productos: {ex.Message}",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void TraspasoLocalButton_Click(object sender, EventArgs e)
        {
            try
            {
                var logger = LoggerFactory.Create(builder => builder.AddConsole())
                    .CreateLogger<TraspasosForm>();
                var traspasoForm = new TraspasosForm(logger, _inventarioDatabaseService, _sqliteService); // Cambio aquí
                traspasoForm.Show();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al abrir el formulario de traspasos");
                MessageBox.Show($"Error al abrir el formulario de traspasos: {ex.Message}",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
