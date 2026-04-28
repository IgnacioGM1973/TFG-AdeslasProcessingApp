<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class IncioForms
    Inherits System.Windows.Forms.Form

    'Form reemplaza a Dispose para limpiar la lista de componentes.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Requerido por el Diseñador de Windows Forms
    Private components As System.ComponentModel.IContainer

    'NOTA: el Diseñador de Windows Forms necesita el siguiente procedimiento
    'Se puede modificar usando el Diseñador de Windows Forms.  
    'No lo modifique con el editor de código.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(IncioForms))
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.ArchivoToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SeleccionarFicheroToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SalirToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ImportarDatosToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SeleccionarFicheroToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.SinAsistenciaViajeToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ImportarSeniorToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ImnportarEnviosToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ImportarEmpresaToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ImportarTaarjetasNoProcesarToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.MantenimientoToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ImportarHispapostToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.ProgressBar2 = New System.Windows.Forms.ProgressBar()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.eqT_Informativa = New System.Windows.Forms.Label()
        Me.MenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'MenuStrip1
        '
        Me.MenuStrip1.AutoSize = False
        Me.MenuStrip1.ImageScalingSize = New System.Drawing.Size(20, 20)
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ArchivoToolStripMenuItem, Me.ImportarDatosToolStripMenuItem, Me.MantenimientoToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Padding = New System.Windows.Forms.Padding(4, 2, 0, 2)
        Me.MenuStrip1.Size = New System.Drawing.Size(600, 23)
        Me.MenuStrip1.TabIndex = 0
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'ArchivoToolStripMenuItem
        '
        Me.ArchivoToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.SeleccionarFicheroToolStripMenuItem, Me.SalirToolStripMenuItem})
        Me.ArchivoToolStripMenuItem.Name = "ArchivoToolStripMenuItem"
        Me.ArchivoToolStripMenuItem.Size = New System.Drawing.Size(60, 19)
        Me.ArchivoToolStripMenuItem.Text = "Archivo"
        '
        'SeleccionarFicheroToolStripMenuItem
        '
        Me.SeleccionarFicheroToolStripMenuItem.Name = "SeleccionarFicheroToolStripMenuItem"
        Me.SeleccionarFicheroToolStripMenuItem.Size = New System.Drawing.Size(176, 22)
        Me.SeleccionarFicheroToolStripMenuItem.Text = "Seleccionar Fichero"
        '
        'SalirToolStripMenuItem
        '
        Me.SalirToolStripMenuItem.Name = "SalirToolStripMenuItem"
        Me.SalirToolStripMenuItem.Size = New System.Drawing.Size(176, 22)
        Me.SalirToolStripMenuItem.Text = "Salir"
        '
        'ImportarDatosToolStripMenuItem
        '
        Me.ImportarDatosToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.SeleccionarFicheroToolStripMenuItem1, Me.SinAsistenciaViajeToolStripMenuItem, Me.ImportarSeniorToolStripMenuItem, Me.ImnportarEnviosToolStripMenuItem, Me.ImportarEmpresaToolStripMenuItem, Me.ImportarTaarjetasNoProcesarToolStripMenuItem})
        Me.ImportarDatosToolStripMenuItem.Name = "ImportarDatosToolStripMenuItem"
        Me.ImportarDatosToolStripMenuItem.Size = New System.Drawing.Size(65, 19)
        Me.ImportarDatosToolStripMenuItem.Text = "Importar"
        '
        'SeleccionarFicheroToolStripMenuItem1
        '
        Me.SeleccionarFicheroToolStripMenuItem1.Name = "SeleccionarFicheroToolStripMenuItem1"
        Me.SeleccionarFicheroToolStripMenuItem1.Size = New System.Drawing.Size(230, 22)
        Me.SeleccionarFicheroToolStripMenuItem1.Text = "Importar MGA"
        '
        'SinAsistenciaViajeToolStripMenuItem
        '
        Me.SinAsistenciaViajeToolStripMenuItem.Name = "SinAsistenciaViajeToolStripMenuItem"
        Me.SinAsistenciaViajeToolStripMenuItem.Size = New System.Drawing.Size(230, 22)
        Me.SinAsistenciaViajeToolStripMenuItem.Text = "Importar Sin Extranjero"
        '
        'ImportarSeniorToolStripMenuItem
        '
        Me.ImportarSeniorToolStripMenuItem.Name = "ImportarSeniorToolStripMenuItem"
        Me.ImportarSeniorToolStripMenuItem.Size = New System.Drawing.Size(230, 22)
        Me.ImportarSeniorToolStripMenuItem.Text = "Importar Seniors"
        '
        'ImnportarEnviosToolStripMenuItem
        '
        Me.ImnportarEnviosToolStripMenuItem.Name = "ImnportarEnviosToolStripMenuItem"
        Me.ImnportarEnviosToolStripMenuItem.Size = New System.Drawing.Size(230, 22)
        Me.ImnportarEnviosToolStripMenuItem.Text = "Importar Envíos"
        '
        'ImportarEmpresaToolStripMenuItem
        '
        Me.ImportarEmpresaToolStripMenuItem.Name = "ImportarEmpresaToolStripMenuItem"
        Me.ImportarEmpresaToolStripMenuItem.Size = New System.Drawing.Size(230, 22)
        Me.ImportarEmpresaToolStripMenuItem.Text = "Importar Empresas"
        '
        'ImportarTaarjetasNoProcesarToolStripMenuItem
        '
        Me.ImportarTaarjetasNoProcesarToolStripMenuItem.Name = "ImportarTaarjetasNoProcesarToolStripMenuItem"
        Me.ImportarTaarjetasNoProcesarToolStripMenuItem.Size = New System.Drawing.Size(230, 22)
        Me.ImportarTaarjetasNoProcesarToolStripMenuItem.Text = "Importar Tarjetas No Procesar"
        '
        'MantenimientoToolStripMenuItem
        '
        Me.MantenimientoToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ImportarHispapostToolStripMenuItem1})
        Me.MantenimientoToolStripMenuItem.Name = "MantenimientoToolStripMenuItem"
        Me.MantenimientoToolStripMenuItem.Size = New System.Drawing.Size(101, 19)
        Me.MantenimientoToolStripMenuItem.Text = "Mantenimiento"
        '
        'ImportarHispapostToolStripMenuItem1
        '
        Me.ImportarHispapostToolStripMenuItem1.Name = "ImportarHispapostToolStripMenuItem1"
        Me.ImportarHispapostToolStripMenuItem1.Size = New System.Drawing.Size(176, 22)
        Me.ImportarHispapostToolStripMenuItem1.Text = "Importar Hispapost"
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(0, 25)
        Me.TextBox1.Margin = New System.Windows.Forms.Padding(2)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(601, 20)
        Me.TextBox1.TabIndex = 1
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(0, 46)
        Me.Button1.Margin = New System.Windows.Forms.Padding(2)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(182, 62)
        Me.Button1.TabIndex = 2
        Me.Button1.Text = "Ejecutar"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.ProgressBar1.Location = New System.Drawing.Point(0, 347)
        Me.ProgressBar1.Margin = New System.Windows.Forms.Padding(2)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(600, 19)
        Me.ProgressBar1.TabIndex = 3
        '
        'ProgressBar2
        '
        Me.ProgressBar2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ProgressBar2.Location = New System.Drawing.Point(0, 330)
        Me.ProgressBar2.Margin = New System.Windows.Forms.Padding(2)
        Me.ProgressBar2.Name = "ProgressBar2"
        Me.ProgressBar2.Size = New System.Drawing.Size(600, 36)
        Me.ProgressBar2.TabIndex = 4
        Me.ProgressBar2.Visible = False
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(205, 46)
        Me.Button2.Margin = New System.Windows.Forms.Padding(2)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(169, 62)
        Me.Button2.TabIndex = 5
        Me.Button2.Text = "Separar Tajetas"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'eqT_Informativa
        '
        Me.eqT_Informativa.AutoSize = True
        Me.eqT_Informativa.Location = New System.Drawing.Point(45, 280)
        Me.eqT_Informativa.Name = "eqT_Informativa"
        Me.eqT_Informativa.Size = New System.Drawing.Size(0, 13)
        Me.eqT_Informativa.TabIndex = 6
        '
        'IncioForms
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.ClientSize = New System.Drawing.Size(600, 366)
        Me.Controls.Add(Me.eqT_Informativa)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.ProgressBar2)
        Me.Controls.Add(Me.ProgressBar1)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.MenuStrip1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.Name = "IncioForms"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Proceso de datos para Adeslas"
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents MenuStrip1 As MenuStrip
    Friend WithEvents ArchivoToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents SeleccionarFicheroToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents SalirToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents TextBox1 As TextBox
    Friend WithEvents Button1 As Button
    Friend WithEvents ImportarDatosToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents SeleccionarFicheroToolStripMenuItem1 As ToolStripMenuItem
    Friend WithEvents ProgressBar1 As ProgressBar
    Friend WithEvents SinAsistenciaViajeToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ImportarSeniorToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ImnportarEnviosToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ImportarEmpresaToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ProgressBar2 As ProgressBar
    Friend WithEvents ImportarTaarjetasNoProcesarToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents MantenimientoToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ImportarHispapostToolStripMenuItem1 As ToolStripMenuItem
    Friend WithEvents Button2 As Button
    Friend WithEvents eqT_Informativa As Label
End Class
