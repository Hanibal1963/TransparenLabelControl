' ****************************************************************************************************************
' TransparentLabel.vb
' © 2024 by Andreas Sauer
' ****************************************************************************************************************
'

Imports System.ComponentModel
Imports System.Drawing
Imports System.Windows.Forms

''' <summary></summary>
<ProvideToolboxControl("TransparentLabelControl.TransparenLabel", False)>
<Description("")>
Public Class TransparentLabel

    Public Sub New()

        'Dieser Aufruf ist für den Designer erforderlich.
        Me.InitializeComponent()

        'Fügen Sie Initialisierungen nach dem InitializeComponent()-Aufruf hinzu.
        Me.InitializeStyles()

    End Sub

    Private Sub InitializeStyles()

        Me.SetStyle(ControlStyles.Opaque, True)
        Me.SetStyle(ControlStyles.OptimizedDoubleBuffer, False)

    End Sub

    Protected Overrides ReadOnly Property CreateParams As CreateParams

        Get
            Dim cp As CreateParams = MyBase.CreateParams
            'WS EX TRANSPARENT aktivieren
            'https://learn.microsoft.com/en-us/windows/win32/winmsg/extended-window-styles
            cp.ExStyle = cp.ExStyle Or &H20
            Return cp
        End Get

    End Property

    ''' <summary>Nicht Relevant</summary>
    Public Overrides Property BackColor As Color
        Get
            Return MyBase.BackColor
        End Get
        Set(value As Color)
            MyBase.BackColor = value
        End Set
    End Property

    ''' <summary>Nicht Relevant</summary>
    Public Overrides Property BackgroundImage As Image
        Get
            Return MyBase.BackgroundImage
        End Get
        Set(value As Image)
            MyBase.BackgroundImage = value
        End Set
    End Property

    ''' <summary>Nicht Relevant</summary>
    Public Overrides Property BackgroundImageLayout As ImageLayout
        Get
            Return MyBase.BackgroundImageLayout
        End Get
        Set(value As ImageLayout)
            MyBase.BackgroundImageLayout = value
        End Set
    End Property

    Public Overrides Property Font As Font
        Get
            Return MyBase.Font
        End Get
        Set(value As Font)
            MyBase.Font = value
            Me.RecreateHandle()
        End Set
    End Property

    Public Overrides Property ForeColor As Color
        Get
            Return MyBase.ForeColor
        End Get
        Set(value As Color)
            MyBase.ForeColor = value
            Me.RecreateHandle()
        End Set
    End Property

    Public Overrides Property Text As String
        Get
            Return MyBase.Text
        End Get
        Set(value As String)
            MyBase.Text = value
            Me.RecreateHandle()
        End Set
    End Property

    Protected Overrides Sub OnPaint(ByVal e As System.Windows.Forms.PaintEventArgs)
        MyBase.OnPaint(e)

        'Benutzerdefinierten Zeichnungscode hier einfügen
        Dim g As Graphics = Me.CreateGraphics

        Dim h As Single = g.MeasureString(Me.Text, Me.Font).Height
        Dim w As Single = g.MeasureString(Me.Text, Me.Font).Width
        Dim s As String = $"Höhe: {h}, Breite: {w}{vbCrLf}{Me.Text}"

        'Me.Height = CInt(h)
        'Me.Width = CInt(w)

        g.DrawString(s, Me.Font, New SolidBrush(Me.ForeColor), Me.ClientRectangle)


    End Sub

End Class
