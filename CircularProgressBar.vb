Imports System.ComponentModel
Imports System.Drawing.Drawing2D
Public Class CircularProgressBar
    Inherits Control
    Public Enum _ProgressShape
        Round
        Flat
    End Enum
    Public Enum _TextMode
        None
        Value
        Percentage
        [Custom]
    End Enum
    Private _Value As Integer = 0
    Private _Maximum As Integer = 100
    Private _LineWitdh As Integer = 5
    Private _BarWidth As Single = 14
    Private _ProgressColor1 As Color = Color.Orange
    Private _ProgressColor2 As Color = Color.Orange
    Private _LineColor As Color = Color.LightGray
    Private _GradientMode As LinearGradientMode = LinearGradientMode.ForwardDiagonal
    Private ProgressShapeVal As _ProgressShape
    Private ProgressTextMode As _TextMode
    Private _ShadowOffset As Single = 0
    Public Sub New()
        'SetStyle(ControlStyles.AllPaintingInWmPaint Or ControlStyles.UserPaint Or ControlStyles.ResizeRedraw Or ControlStyles.OptimizedDoubleBuffer Or ControlStyles.SupportsTransparentBackColor, True)
        SetStyle(ControlStyles.SupportsTransparentBackColor, True)
        SetStyle(ControlStyles.Opaque, True)
        'BackColor = Color.Transparent
        BackColor = SystemColors.Control
        ForeColor = Color.DimGray
        'Size = New Size(130, 130)
        Font = New Font("Segoe UI", 15)
        'MinimumSize = New Size(100, 100)
        DoubleBuffered = True
        LineColor = Color.LightGray
        Value = 50
        ProgressShape = _ProgressShape.Flat
        TextMode = _TextMode.Percentage
    End Sub
    <Description("Integer Value that determines the position of the Progress Bar."), Category("Behavior")>
    Public Property Value() As Long
        Get
            Return _Value
        End Get
        Set
            If Value > _Maximum Then
                Value = _Maximum
            End If
            _Value = Value
            Invalidate()
        End Set
    End Property
    <Description("Gets or Sets the Maximum Value of the Progress bar."), Category("Behavior")>
    Public Property Maximum() As Long
        Get
            Return _Maximum
        End Get
        Set
            If Value < 1 Then
                Value = 1
            End If
            _Maximum = Value
            Invalidate()
        End Set
    End Property
    <Description("Initial Color of Progress Bar"), Category("Appearance")>
    Public Property BarColor1() As Color
        Get
            Return _ProgressColor1
        End Get
        Set
            _ProgressColor1 = Value
            Invalidate()
        End Set
    End Property
    <Description("Final Color of Progress Bar"), Category("Appearance")> Public Property BarColor2() As Color
        Get
            Return _ProgressColor2
        End Get
        Set
            _ProgressColor2 = Value
            Invalidate()
        End Set
    End Property
    <Description("Progress Bar Width"), Category("Appearance")>
    Public Property BarWidth() As Single
        Get
            Return _BarWidth
        End Get
        Set
            _BarWidth = Value
            Invalidate()
        End Set
    End Property
    <Description("Modo del Gradiente de Color"), Category("Appearance")>
    Public Property GradientMode() As LinearGradientMode
        Get
            Return _GradientMode
        End Get
        Set
            _GradientMode = Value
            Invalidate()
        End Set
    End Property
    <Description("Color de la Linea Intermedia"), Category("Appearance")>
    Public Property LineColor() As Color
        Get
            Return _LineColor
        End Get
        Set
            _LineColor = Value
            Invalidate()
        End Set
    End Property
    <Description("Width of Intermediate Line"), Category("Appearance")>
    Public Property LineWidth() As Integer
        Get
            Return _LineWitdh
        End Get
        Set
            _LineWitdh = Value
            Invalidate()
        End Set
    End Property
    <Description("Get or Set the Shape of the progress bar terminals."), Category("Appearance")>
    Public Property ProgressShape() As _ProgressShape
        Get
            Return ProgressShapeVal
        End Get
        Set
            ProgressShapeVal = Value
            Invalidate()
        End Set
    End Property
    <Description("Modo del Gradiente de Color"), Category("Appearance")>
    Public Property ShadowOffset() As Integer
        Get
            Return _ShadowOffset
        End Get
        Set
            If Value > 2 Then
                Value = 2
            End If
            _ShadowOffset = Value
            Invalidate()
        End Set
    End Property
    <Description("Get or Set the Mode as the Text is displayed inside the Progress bar."), Category("Behavior")>
    Public Property TextMode() As _TextMode
        Get
            Return ProgressTextMode
        End Get
        Set
            ProgressTextMode = Value
            Invalidate()
        End Set
    End Property

    Protected Overloads Overrides Sub OnResize(e As EventArgs)
        MyBase.OnResize(e)
        SetStandardSize()
    End Sub
    Protected Overloads Overrides Sub OnSizeChanged(e As EventArgs)
        MyBase.OnSizeChanged(e)
        SetStandardSize()
    End Sub
    Protected Overloads Overrides Sub OnPaintBackground(p As PaintEventArgs)
        MyBase.OnPaintBackground(p)
    End Sub
    Private Sub SetStandardSize()
        Dim _Size As Integer = Math.Max(Width, Height)
        Size = New Size(_Size, _Size)
    End Sub
    Public Sub Increment(Val As Integer)
        _Value += Val
        Invalidate()
    End Sub
    Public Sub Decrement(Val As Integer)
        _Value -= Val
        Invalidate()
    End Sub
    'Protected Overloads Overrides Sub OnPaint(e As PaintEventArgs)
    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        Using bitmap As New Bitmap(Width, Height)
            Using graphics As Graphics = Graphics.FromImage(bitmap)
                graphics.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBilinear
                graphics.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality
                graphics.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality
                graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias
                graphics.TextRenderingHint = Drawing.Text.TextRenderingHint.AntiAliasGridFit
                PaintTransparentBackground(Me, e)
                Using mBackColor As Brush = New SolidBrush(BackColor)
                    graphics.FillEllipse(mBackColor, 18, 18, (Width - 48) + 12, (Height - 48) + 12)
                End Using
                Using pen2 As New Pen(LineColor, LineWidth)
                    graphics.DrawEllipse(pen2, 18, 18, (Width - 48) + 12, (Height - 48) + 12)
                End Using
                Using brush As New LinearGradientBrush(ClientRectangle, _ProgressColor1, _ProgressColor2, GradientMode)
                    Using pen As New Pen(brush, BarWidth)
                        Select Case ProgressShapeVal
                            Case _ProgressShape.Round
                                pen.StartCap = LineCap.Round
                                pen.EndCap = LineCap.Round
                                Exit Select
                            Case _ProgressShape.Flat
                                pen.StartCap = LineCap.Flat
                                pen.EndCap = LineCap.Flat
                                Exit Select
                        End Select
                        graphics.DrawArc(pen, 18, 18, (Width - 35) - 2, (Height - 35) - 2, -90, CType(Math.Round(360 / _Maximum) * _Value, Integer))

                    End Using
                End Using

                Select Case TextMode
                    Case _TextMode.None
                        Text = String.Empty
                        Exit Select
                    Case _TextMode.Value
                        Text = _Value.ToString()
                        Exit Select
                    Case _TextMode.Percentage
                        Text = Convert.ToString(Convert.ToInt32((100 / _Maximum) * _Value))
                        Exit Select
                    Case _TextMode.Custom
                        Text = Text
                        Exit Select
                    Case Else
                        Exit Select
                End Select

                If Not Text = String.Empty Then
                    Dim MS As SizeF = graphics.MeasureString(Text, Font)
                    Dim shadowBrush As New SolidBrush(Color.FromArgb(100, ForeColor))
                    If ShadowOffset > 0 Then graphics.DrawString(Text, Font, shadowBrush, (Width / 2 - MS.Width / 2) + ShadowOffset, (Height / 2 - MS.Height / 2) + ShadowOffset)
                    graphics.DrawString(Text, Font, New SolidBrush(ForeColor), (Width / 2 - MS.Width / 2), (Height / 2 - MS.Height / 2))
                End If
                MyBase.OnPaint(e)

                e.Graphics.DrawImage(bitmap, 0, 0)
                graphics.Dispose()
            End Using
        End Using
    End Sub
    Private Shared Sub PaintTransparentBackground(c As Control, e As PaintEventArgs)
        If c.Parent Is Nothing OrElse Not Application.RenderWithVisualStyles Then
            Return
        End If
        ButtonRenderer.DrawParentBackground(e.Graphics, c.ClientRectangle, c)
    End Sub

    Private Sub FillCircle(g As Graphics, brush As Brush, centerX As Single, centerY As Single, radius As Single)
        g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBilinear
        g.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality
        g.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality
        g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias
        Using gp As New System.Drawing.Drawing2D.GraphicsPath()
            g.FillEllipse(brush, centerX - radius, centerY - radius, radius + radius, radius + radius)
        End Using
    End Sub
End Class
