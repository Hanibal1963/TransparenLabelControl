' ****************************************************************************************************************
' TransparentLabel.vb
' © 2024 by Andreas Sauer
' ****************************************************************************************************************
'

Imports System.ComponentModel
Imports System.Drawing
Imports System.Windows.Forms

''' <summary>Ein Steuerelement zum Anzeigen eines Textes mit durchscheinendem Hintergrund.</summary>
<ProvideToolboxControl("SchlumpfSoft Controls", False)>
<Description("Ein Steuerelement zum Anzeigen eines Textes mit durchscheinendem Hintergrund.")>
<ToolboxItem(True)>
<ToolboxBitmap(GetType(TransparentLabel), "TransparentLabel.bmp")>
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

    ''' <summary>Hiermit wird die Möglichkeit der Transparenz aktiviert</summary>
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
    <Browsable(False)>
    <EditorBrowsable(EditorBrowsableState.Never)>
    Public Overrides Property BackColor As Color
        Get
            Return MyBase.BackColor
        End Get
        Set(value As Color)
            'MyBase.BackColor = value
        End Set
    End Property

    ''' <summary>Nicht Relevant</summary>
    <Browsable(False)>
    <EditorBrowsable(EditorBrowsableState.Never)>
    Public Overrides Property BackgroundImage As Image
        Get
            Return MyBase.BackgroundImage
        End Get
        Set(value As Image)
            'MyBase.BackgroundImage = value
        End Set
    End Property

    ''' <summary>Nicht Relevant</summary>
    <Browsable(False)>
    <EditorBrowsable(EditorBrowsableState.Never)>
    Public Overrides Property BackgroundImageLayout As ImageLayout
        Get
            Return MyBase.BackgroundImageLayout
        End Get
        Set(value As ImageLayout)
            'MyBase.BackgroundImageLayout = value
        End Set
    End Property

    ''' <summary>Nicht Relevant</summary>
    <Browsable(False)>
    <EditorBrowsable(EditorBrowsableState.Never)>
    Public Overloads Property FlatStyle As FlatStyle
        Get
            Return MyBase.FlatStyle
        End Get
        Set(value As FlatStyle)
            MyBase.FlatStyle = value
        End Set
    End Property

End Class
