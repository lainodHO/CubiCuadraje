using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace Caso_Cubicuadraje
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            // Establece el contexto de la licencia
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
        }

        private void btnCargarArchivo_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Archivos Excel|*.xlsx;*.xls";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string filePath = openFileDialog.FileName;
                CargarArchivoExcel(filePath);
            }
        }

        private void btnCalcular_Click(object sender, EventArgs e)
        {
            CalcularResultados();
        }

        private void CargarArchivoExcel(string filePath)
        {
            try
            {
                using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                    DataTable dataTable = new DataTable();

                    for (int i = 1; i <= worksheet.Dimension.Columns; i++)
                    {
                        dataTable.Columns.Add(worksheet.Cells[1, i].Value.ToString());
                    }

                    for (int row = 2; row <= worksheet.Dimension.Rows; row++)
                    {
                        DataRow dataRow = dataTable.NewRow();
                        for (int col = 1; col <= worksheet.Dimension.Columns; col++)
                        {
                            dataRow[col - 1] = worksheet.Cells[row, col].Value;
                        }
                        dataTable.Rows.Add(dataRow);
                    }

                    dataGridViewResultados.DataSource = dataTable;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al cargar el archivo Excel: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void CalcularResultados()
        {
            try
            {
                // Verifica si hay al menos una fila en el DataGridView
                if (dataGridViewResultados.Rows.Count > 0)
                {
                    // Verifica si las columnas requeridas están presentes en el DataGridView
                    if (dataGridViewResultados.Columns.Contains("DOCVTAS") &&
                        dataGridViewResultados.Columns.Contains("CENTRO") &&
                        dataGridViewResultados.Columns.Contains("DESTINAT") &&
                        dataGridViewResultados.Columns.Contains("P_B_KG") &&
                        dataGridViewResultados.Columns.Contains("VOLUMEN_M3") &&
                        dataGridViewResultados.Columns.Contains("PALLET") &&
                        dataGridViewResultados.Columns.Contains("APILABLESINO") &&
                        dataGridViewResultados.Columns.Contains("POS REALES") &&
                        dataGridViewResultados.Columns.Contains("TIPOCAM"))
                    {
                        // Recorre cada fila de la DataGridView
                        foreach (DataGridViewRow row in dataGridViewResultados.Rows)
                        {
                            // Verifica si la fila no es nueva y está completa
                            if (!row.IsNewRow && row.Cells.Count >= 9)
                            {
                                // Obtén los valores de las celdas
                                string centro = row.Cells["CENTRO"].Value.ToString();
                                string destino = row.Cells["DESTINAT"].Value.ToString();
                                double peso = Convert.ToDouble(row.Cells["P_B_KG"].Value);
                                double volumen = Convert.ToDouble(row.Cells["VOLUMEN_M3"].Value);
                                int pallet = Convert.ToInt32(row.Cells["PALLET"].Value);
                                string apilable = row.Cells["APILABLESINO"].Value.ToString();

                                // Calcula las posiciones reales
                                int posicionesReales = (apilable == "SI") ? pallet / 2 : pallet;

                                // Verifica el tipo de camión y actualiza la columna correspondiente
                                if ((peso >= 19 && peso <= 23) && (volumen >= 70 && volumen <= 90))
                                {
                                    row.Cells["TIPOCAM"].Value = "REG 1";
                                }
                                else if ((peso >= 14 && peso <= 23) && (volumen >= 42 && volumen <= 90) && (destino == "DESTINO 1" || destino == "DESTINO 2"))
                                {
                                    row.Cells["TIPOCAM"].Value = "BKHL 1";
                                }
                                else if ((peso >= 19 && peso <= 23) && (volumen >= 70 && volumen <= 90) && ((centro == "ORIGEN 1" && destino == "DESTINO 3") || (centro == "ORIGEN 1" && destino == "DESTINO 4") || (centro == "ORIGEN 1" && destino == "DESTINO 5")))
                                {
                                    row.Cells["TIPOCAM"].Value = "MR 1";
                                }
                                else
                                {
                                    row.Cells["TIPOCAM"].Value = "SIN CONSOLIDAR";
                                }
                            }
                        }
                        MessageBox.Show("Se han calculado los resultados y actualizado el DataGridView.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        MessageBox.Show("El DataGridView no contiene todas las columnas requeridas.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    MessageBox.Show("El DataGridView no contiene ninguna fila de datos.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al calcular y actualizar los resultados: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

    }
}
