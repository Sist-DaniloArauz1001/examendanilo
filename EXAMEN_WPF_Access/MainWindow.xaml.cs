using System;
using System.Data;
using System.Data.OleDb;
using System.Windows;
using System.Windows.Controls;

namespace EXAMEN_WPF_Access
{
    public partial class MainWindow : Window
    {
        string conexion = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\ExamenParcial\ExamenParcial.accdb;Persist Security Info=False;";
        OleDbConnection conn;

        // Variable para guardar el ID del registro que seleccionemos en la tabla
        int idSeleccionado = 0;

        public MainWindow()
        {
            InitializeComponent();
            conn = new OleDbConnection(conexion);
            // Ya no cargamos los datos aquí, la grilla iniciará vacía.
        }

        // --- 1. MOSTRAR (READ) ---
        private void MostrarDatos()
        {
            try
            {
                conn.Open();
                string query = "SELECT * FROM Proveedores";
                OleDbCommand cmd = new OleDbCommand(query, conn);
                OleDbDataAdapter adapter = new OleDbDataAdapter(cmd);
                DataTable dt = new DataTable();
                adapter.Fill(dt);
                dgvProveedores.ItemsSource = dt.DefaultView;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al cargar datos: " + ex.Message);
            }
            finally
            {
                if (conn.State == ConnectionState.Open) conn.Close();
            }
        }

        private void btnMostrar_Click(object sender, RoutedEventArgs e)
        {
            MostrarDatos();
            LimpiarCampos();
        }

        // --- 2. INSERTAR (CREATE) ---
        private void btnInsertar_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtNombre.Text))
            {
                MessageBox.Show("Por favor, ingrese al menos el Nombre.");
                return;
            }

            try
            {
                conn.Open();
                // Usamos parámetros (?) para evitar errores de sintaxis y seguridad
                string query = "INSERT INTO Proveedores (Nombre, CPost, Produc, Ubica) VALUES (?, ?, ?, ?)";
                OleDbCommand cmd = new OleDbCommand(query, conn);

                cmd.Parameters.AddWithValue("?", txtNombre.Text);
                cmd.Parameters.AddWithValue("?", Convert.ToInt32(txtCPost.Text)); // Es Entero Largo en Access
                cmd.Parameters.AddWithValue("?", txtProduc.Text);
                cmd.Parameters.AddWithValue("?", txtUbica.Text);

                cmd.ExecuteNonQuery();
                MessageBox.Show("Registro insertado exitosamente.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al insertar: Verifique que el Código Postal sea un número. Detalle: " + ex.Message);
            }
            finally
            {
                if (conn.State == ConnectionState.Open) conn.Close();
                MostrarDatos(); // Refrescar la tabla
                LimpiarCampos(); // Dejar limpio para el siguiente
            }
        }

        // --- 3. ACTUALIZAR (UPDATE) ---
        private void btnActualizar_Click(object sender, RoutedEventArgs e)
        {
            if (idSeleccionado == 0)
            {
                MessageBox.Show("Por favor, seleccione un proveedor de la tabla para actualizar.");
                return;
            }

            try
            {
                conn.Open();
                string query = "UPDATE Proveedores SET Nombre=?, CPost=?, Produc=?, Ubica=? WHERE Id=?";
                OleDbCommand cmd = new OleDbCommand(query, conn);

                cmd.Parameters.AddWithValue("?", txtNombre.Text);
                cmd.Parameters.AddWithValue("?", Convert.ToInt32(txtCPost.Text));
                cmd.Parameters.AddWithValue("?", txtProduc.Text);
                cmd.Parameters.AddWithValue("?", txtUbica.Text);
                cmd.Parameters.AddWithValue("?", idSeleccionado);

                cmd.ExecuteNonQuery();
                MessageBox.Show("Registro actualizado exitosamente.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al actualizar: " + ex.Message);
            }
            finally
            {
                if (conn.State == ConnectionState.Open) conn.Close();
                MostrarDatos();
                LimpiarCampos();
            }
        }

        // --- 4. ELIMINAR (DELETE) ---
        private void btnEliminar_Click(object sender, RoutedEventArgs e)
        {
            if (idSeleccionado == 0)
            {
                MessageBox.Show("Por favor, seleccione un proveedor de la tabla para eliminar.");
                return;
            }

            // Pregunta de confirmación antes de borrar
            MessageBoxResult respuesta = MessageBox.Show("¿Está seguro de eliminar este registro?", "Confirmación", MessageBoxButton.YesNo, MessageBoxImage.Warning);
            if (respuesta == MessageBoxResult.Yes)
            {
                try
                {
                    conn.Open();
                    string query = "DELETE FROM Proveedores WHERE Id=?";
                    OleDbCommand cmd = new OleDbCommand(query, conn);
                    cmd.Parameters.AddWithValue("?", idSeleccionado);

                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Registro eliminado exitosamente.");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error al eliminar: " + ex.Message);
                }
                finally
                {
                    if (conn.State == ConnectionState.Open) conn.Close();
                    MostrarDatos();
                    LimpiarCampos();
                }
            }
        }

        // --- EVENTO EXTRA: Llenar campos al seleccionar la tabla ---
        private void dgvProveedores_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (dgvProveedores.SelectedItem is DataRowView row)
            {
                idSeleccionado = Convert.ToInt32(row["Id"]);
                txtNombre.Text = row["Nombre"].ToString();
                txtCPost.Text = row["CPost"].ToString();
                txtProduc.Text = row["Produc"].ToString();
                txtUbica.Text = row["Ubica"].ToString();
            }
        }

        // Método auxiliar para vaciar los campos
        private void LimpiarCampos()
        {
            txtNombre.Clear();
            txtCPost.Clear();
            txtProduc.Clear();
            txtUbica.Clear();
            idSeleccionado = 0;
        }

        // --- 5. SALIR ---
        private void btnSalir_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}