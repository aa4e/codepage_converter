Imports System.ComponentModel
Imports System.Text
Imports System.Text.RegularExpressions

Public Class CodepageConverter
    Implements INotifyPropertyChanged

#Region "CTOR"

    Public Sub New()
        AddHandler Application.LanguageChanged, AddressOf LanguageChangedHandler
    End Sub

#End Region '/CTOR

#Region "PROPS"

    Public Property ForWeb As Boolean
        Get
            Return _ForWeb
        End Get
        Set(value As Boolean)
            If (_ForWeb <> value) Then
                _ForWeb = value
                NotifyPropertyChanged(NameOf(ForWeb))
                NotifyPropertyChanged(NameOf(StringDecoded))
                NotifyPropertyChanged(NameOf(HexBytes))
                NotifyPropertyChanged(NameOf(HexBytesLength))
            End If
        End Set
    End Property
    Private _ForWeb As Boolean = False

    Public Property TextEncoding As TextEncodings
        Get
            Return _TextEncoding
        End Get
        Set(value As TextEncodings)
            If (_TextEncoding <> value) Then
                _TextEncoding = value
                NotifyPropertyChanged(NameOf(TextEncoding))
                NotifyPropertyChanged(NameOf(StringDecoded))
                NotifyPropertyChanged(NameOf(HexBytes))
                NotifyPropertyChanged(NameOf(HexBytesLength))
            End If
        End Set
    End Property
    Private _TextEncoding As TextEncodings = TextEncodings.None

    Public Property CodePage As String
        Get
            Return _CodePage
        End Get
        Set(value As String)
            value = value.Trim()
            If (_CodePage <> value) Then
                _CodePage = value
                NotifyPropertyChanged(NameOf(CodePage))
                NotifyPropertyChanged(NameOf(TextEncoding))
                NotifyPropertyChanged(NameOf(StringDecoded))
                NotifyPropertyChanged(NameOf(HexBytes))
                NotifyPropertyChanged(NameOf(HexBytesLength))
            End If
        End Set
    End Property
    Private _CodePage As String = "10006"

    ''' <summary>
    ''' Направление преобразования.
    ''' </summary>
    Public Property FromBytesToString As Boolean
        Get
            Return _FromBytesToString
        End Get
        Set(value As Boolean)
            If (_FromBytesToString <> value) Then
                _FromBytesToString = value
                NotifyPropertyChanged(NameOf(FromBytesToString))
            End If
        End Set
    End Property
    Private _FromBytesToString As Boolean = True

    Public Property HexBytes As String
        Get
            If FromBytesToString Then
                Return _HexBytes
            Else
                Return GetStringEncoded(StringDecoded)
            End If
        End Get
        Set(value As String)
            _HexBytes = value
            NotifyPropertyChanged(NameOf(StringDecoded))
        End Set
    End Property
    Private _HexBytes As String = "25 26 27 28 29 30 31 32 33 34 35 36 37 38 39 40 41 42 43 44"

    Public Property StringDecoded As String
        Get
            If FromBytesToString Then
                Return GetBytesDecoded(HexBytes)
            Else
                Return _StringDecoded
            End If
        End Get
        Set(value As String)
            _StringDecoded = value
            NotifyPropertyChanged(NameOf(HexBytes))
            NotifyPropertyChanged(NameOf(HexBytesLength))
            NotifyPropertyChanged(NameOf(StringDecoded))
        End Set
    End Property
    Private _StringDecoded As String = ""

#End Region '/PROPS

#Region "READ-ONLY PROPS"

    Public ReadOnly Property HexBytesLength As Integer
        Get
            Return CInt((HexBytes.Length + 1) / 3)
        End Get
    End Property

    Public ReadOnly Property Encodings As Dictionary(Of TextEncodings, String)
        Get
            Return New Dictionary(Of TextEncodings, String) From {
                {TextEncodings.Ascii, "ASCII"},
                {TextEncodings.CP855, $"Code Page 855 ({Application.Lang("m_Cyr")})"},
                {TextEncodings.Windows1251, "Windows 1251"},
                {TextEncodings.Unicode, "Unicode"},
                {TextEncodings.BigEndianUnicode, "Big Endian Unicode"},
                {TextEncodings.Utf7, "UTF7"},
                {TextEncodings.Utf8, "UTF8"},
                {TextEncodings.Utf32, "UTF32"},
                {TextEncodings.Base64, "Base64"},
                {TextEncodings.ScanCodeXt, "Scan code XT"},
                {TextEncodings.ScanCodeAt, "Scan code AT"},
                {TextEncodings.Some, Application.Lang("m_CharsetCustom")}
            }
        End Get
    End Property

#End Region '/READ-ONLY PROPS

#Region "SCAN-CODE"

    Private ReadOnly ScanCodesXt As New Dictionary(Of String, Byte) From {
        {"A", &H1E},
        {"B", &H30},
        {"C", &H2E},
        {"D", &H20},
        {"E", &H12},
        {"F", &H21},
        {"G", &H22},
        {"H", &H23},
        {"I", &H17},
        {"J", &H24},
        {"K", &H25},
        {"L", &H26},
        {"M", &H32},
        {"N", &H31},
        {"O", &H18},
        {"P", &H19},
        {"Q", &H10},
        {"R", &H13},
        {"S", &H1F},
        {"T", &H14},
        {"U", &H16},
        {"V", &H2F},
        {"W", &H11},
        {"X", &H2D},
        {"Y", &H15},
        {"Z", &H2C},
        {"0", &HB},
        {"1", &H2},
        {"2", &H3},
        {"3", &H4},
        {"4", &H5},
        {"5", &H6},
        {"6", &H7},
        {"7", &H8},
        {"8", &H9},
        {"9", &HA},
        {"~", &H29},
        {"-", &HC},
        {"=", &HD},
        {"\", &H2B},
        {"[", &H1A},
        {"]", &H1B},
        {";", &H27},
        {"'", &H28},
        {",", &H33},
        {".", &H34},
        {"/", &H35},
        {" ", &H39},
        {"  ", &HF}
    }
    Private ReadOnly ScanCodesAt As New Dictionary(Of String, Byte) From {
        {"A", &H1C},
        {"B", &H32},
        {"C", &H21},
        {"D", &H23},
        {"E", &H24},
        {"F", &H2B},
        {"G", &H34},
        {"H", &H33},
        {"I", &H43},
        {"J", &H3B},
        {"K", &H42},
        {"L", &H4B},
        {"M", &H3A},
        {"N", &H31},
        {"O", &H44},
        {"P", &H4D},
        {"Q", &H15},
        {"R", &H2D},
        {"S", &H1B},
        {"T", &H2C},
        {"U", &H3C},
        {"V", &H2A},
        {"W", &H1D},
        {"X", &H22},
        {"Y", &H35},
        {"Z", &H1A},
        {"0", &H8B},
        {"1", &H82},
        {"2", &H83},
        {"3", &H84},
        {"4", &H85},
        {"5", &H86},
        {"6", &H87},
        {"7", &H88},
        {"8", &H89},
        {"9", &H8A},
        {"~", &H89},
        {"-", &H8C},
        {"=", &H82},
        {"\", &HAB},
        {"[", &H9A},
        {"]", &H9B},
        {";", &HA7},
        {"'", &HA8},
        {",", &HB3},
        {".", &HB4},
        {"/", &HB5},
        {" ", &HB9},
        {"  ", &H8F}
    }

    Private Function GetTextByScanCodeXt(bytes As Byte()) As String
        Dim sb As New StringBuilder()
        For Each b As Byte In bytes
            For Each kvp As KeyValuePair(Of String, Byte) In ScanCodesXt
                If (kvp.Value = b) Then
                    sb.Append(kvp.Key)
                    Exit For
                End If
            Next
        Next
        Return sb.ToString()
    End Function

    Private Function GetScanCodeByTextXt(s As String) As Byte()
        Dim bytes As New List(Of Byte)
        s = s.ToUpper()
        For i As Integer = 0 To s.Length - 1
            Dim c As Char = s.Chars(i)
            If ScanCodesXt.ContainsKey(c) Then
                bytes.Add(ScanCodesXt(c))
            End If
        Next
        Return bytes.ToArray()
    End Function

    Private Function GetTextByScanCodeAt(bytes As Byte()) As String
        Dim sb As New StringBuilder()
        For Each b As Byte In bytes
            For Each kvp As KeyValuePair(Of String, Byte) In ScanCodesAt
                If (kvp.Value = b) Then
                    sb.Append(kvp.Key)
                    Exit For
                End If
            Next
        Next
        Return sb.ToString()
    End Function

    Private Function GetScanCodeByTextAt(s As String) As Byte()
        Dim bytes As New List(Of Byte)
        s = s.ToUpper()
        For i As Integer = 0 To s.Length - 1
            Dim c As Char = s.Chars(i)
            If ScanCodesAt.ContainsKey(c) Then
                bytes.Add(ScanCodesAt(c))
            End If
        Next
        Return bytes.ToArray()
    End Function

#End Region '/SCAN-CODE

#Region "METHODS"

    Public Function GetStringEncoded(Optional isReversed As Boolean = False) As String
        If isReversed Then
            Return GetReversedBytes(GetStringEncoded(_StringDecoded))
        Else
            Return GetStringEncoded(_StringDecoded)
        End If
    End Function

    ''' <summary>
    ''' Обращает порядок байтов в строке.
    ''' </summary>
    ''' <param name="hexBytes"></param>
    Private Function GetReversedBytes(hexBytes As String) As String
        Dim sb As New StringBuilder
        Dim hexs As String() = hexBytes.Split(" "c, "-"c)
        For i As Integer = hexs.Length - 1 To 0 Step -1
            sb.Append(hexs(i))
            sb.Append(" ")
        Next
        Return sb.ToString().Trim()
    End Function

    ''' <summary>
    ''' Возвращает текстовую строку, полученную из массива байтов с учётом кодировки.
    ''' </summary>
    Private Function GetBytesDecoded(hexBytesStr As String) As String
        Dim bytes As Byte() = GetBytes(hexBytesStr)
        Dim s As String
        Select Case TextEncoding
            Case TextEncodings.Ascii
                s = Encoding.ASCII.GetString(bytes)
            Case TextEncodings.Unicode
                s = Encoding.Unicode.GetString(bytes)
            Case TextEncodings.Utf7
                s = Encoding.UTF7.GetString(bytes)
            Case TextEncodings.Utf8
                s = Encoding.UTF8.GetString(bytes)
            Case TextEncodings.Utf32
                s = Encoding.UTF32.GetString(bytes)
            Case TextEncodings.BigEndianUnicode
                s = Encoding.BigEndianUnicode.GetString(bytes)
            Case TextEncodings.Windows1251
                s = Encoding.GetEncoding(1251).GetString(bytes)
            Case TextEncodings.CP855
                s = Encoding.GetEncoding(855).GetString(bytes)
            Case TextEncodings.Base64
                s = Convert.ToBase64String(bytes)
            Case TextEncodings.ScanCodeXt
                s = GetTextByScanCodeXt(bytes)
            Case TextEncodings.ScanCodeAt
                s = GetTextByScanCodeAt(bytes)
            Case TextEncodings.Some
                Try
                    s = Encoding.GetEncoding(CInt(CodePage)).GetString(bytes)
                Catch ex As Exception
                    s = ex.Message
                End Try
            Case Else
                s = Encoding.ASCII.GetString(bytes)
        End Select
        Return s
    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="decodedString"></param>
    Private Function GetStringEncoded(decodedString As String) As String
        Dim s As String
        Select Case TextEncoding
            Case TextEncodings.Ascii
                s = GetHexString(Encoding.ASCII.GetBytes(decodedString))
            Case TextEncodings.Unicode
                s = GetHexString(Encoding.Unicode.GetBytes(decodedString))
            Case TextEncodings.Utf7
                s = GetHexString(Encoding.UTF7.GetBytes(decodedString))
            Case TextEncodings.Utf8
                s = GetHexString(Encoding.UTF8.GetBytes(decodedString))
            Case TextEncodings.Utf32
                s = GetHexString(Encoding.UTF32.GetBytes(decodedString))
            Case TextEncodings.BigEndianUnicode
                s = GetHexString(Encoding.BigEndianUnicode.GetBytes(decodedString))
            Case TextEncodings.Windows1251
                s = GetHexString(Encoding.GetEncoding(1251).GetBytes(decodedString))
            Case TextEncodings.CP855
                s = GetHexString(Encoding.GetEncoding(855).GetBytes(decodedString))
            Case TextEncodings.Base64
                s = GetHexString(Convert.FromBase64String(decodedString))
            Case TextEncodings.ScanCodeXt
                s = GetHexString(GetScanCodeByTextXt(decodedString))
            Case TextEncodings.ScanCodeAt
                s = GetHexString(GetScanCodeByTextAt(decodedString))
            Case TextEncodings.Some
                Try
                    s = GetHexString(Encoding.GetEncoding(CInt(CodePage)).GetBytes(decodedString))
                Catch ex As Exception
                    s = ex.Message
                End Try
            Case Else
                s = GetHexString(Encoding.ASCII.GetBytes(decodedString))
        End Select
        Return s
    End Function

    ''' <summary>
    ''' Возвращает массив байтов по hex-строке.
    ''' </summary>
    ''' <param name="hexString"></param>
    Private Function GetBytes(hexString As String) As Byte()
        Dim hexRegex As New Regex("[a-fA-F0-9]{2}")
        Dim bytes As New List(Of Byte)
        For Each m As Match In hexRegex.Matches(hexString)
            Try
                bytes.Add(Convert.ToByte(m.Value, 16))
            Catch ex As Exception
            End Try
        Next
        Return bytes.ToArray()
    End Function

    ''' <summary>
    ''' Возвращает hex-строку из массива байтов.
    ''' </summary>
    ''' <param name="bytes">Массив байтов.</param>
    Private Function GetHexString(bytes As Byte()) As String
        Dim sb As New StringBuilder()
        For Each b As Byte In bytes
            If ForWeb Then
                sb.Append("%")
            End If
            sb.Append(b.ToString("X2"))
            If (Not ForWeb) Then
                sb.Append(" "c)
            End If
        Next
        If (Not ForWeb) Then
            If (sb.Length > 1) Then
                sb = sb.Remove(sb.Length - 1, 1)
            End If
        End If
        Return sb.ToString()
    End Function

#End Region '/METHODS

#Region "ENUM"

    Public Enum TextEncodings As Integer
        Ascii
        CP855
        Windows1251
        Unicode
        BigEndianUnicode
        Utf7
        Utf8
        Utf32
        Some
        Base64
        ScanCodeXt
        ScanCodeAt
        AltScanCode
        None
    End Enum

#End Region '/ENUM

    Private Sub LanguageChangedHandler(sender As Object, e As EventArgs)
        NotifyPropertyChanged(NameOf(Encodings))
    End Sub

#Region "INOTIFY"

    Public Event PropertyChanged(sender As Object, ByVal e As PropertyChangedEventArgs) Implements INotifyPropertyChanged.PropertyChanged
    Private Sub NotifyPropertyChanged(propName As String)
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(propName))
    End Sub

#End Region '/INOTIFY

End Class
