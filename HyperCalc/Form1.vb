Option Explicit Off

Imports System.ComponentModel
Imports HyperLib


Public Class Form1


    Dim a, b As Hyper
    Dim OALprec% = 900
    Dim OALprecOnDIV% = 900
    Dim XLdivPrec1% = 1444
    Dim XLdivPrec2% = 44
    Dim lastTextInp$ = ""

    Function olepšaj(s$) As String
'this function removes unnecessary zeros from the decimal output string 
        s1$ = ""
        If Strings.Left(s, 1) <> "-" Then
            If InStr(s, "e") Then
                s1 = s.TrimStart("0")
            ElseIf InStr(s, ".") Then
                s1 = s.Trim("0")
            Else
                s1 = s.TrimStart("0")
            End If

            If Strings.Left(s1, 1) = "." Then s1 = "0" + s1

        Else
            s1$ = s.Remove(0, 1)
            If InStr(s, "e") Then
                s1 = s1.TrimStart("0")
            ElseIf InStr(s, ".") Then
                s1 = s1.Trim("0")
            Else
                s1 = s1.TrimStart("0")
            End If
            If Strings.Left(s1, 1) = "." Then
                s1 = "-0" + s1
            Else
                s1 = "-" + s1

            End If
        End If

        If Strings.Right(s1, 1) = "." Then
            s1 += "0"
        End If
        Return s1
    End Function
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Try
            If IsNothing(a) Then GoTo npr
            txt$ = TextBox2.Text

            s$ = txt.Replace(" ", "")

            If UCase(Strings.Left(s, 1)) = "M" Then
                txt = Mid(s, 2)
                memNr% = txt
                b = ms(memNr)
                GoTo napr
            End If

            If InStr(s, ",") Then
                txt = s.Replace(",", ".")
            End If
            If s.Length Then
                b = New Hyper(s)
            ElseIf Not IsNothing(b) Then

            Else
                Exit Sub
            End If

            't = DateAndTime.Timer

napr:

            Dim isApprox As Boolean

            Select Case ComboBox1.SelectedIndex
                Case 0
                    a.Add(b)
                Case 1
                    a.Subtract(b)
                Case 2

                    a *= b

                    a.StripZeros()

                Case 3

                    If chkINT(b) Then
                        a.Divide(b(0), OALprecOnDIV)
                    Else
                        a = BigDivide(a, b)
                        a.StripZeros()
                        isApprox = True
                    End If

                Case 4
                    If chkINT(a) Then
                        x& = a.Divide(b(0), OALprecOnDIV)
                        a = New Hyper(0, 0)
                        a(0) = x
                    Else
                        MsgBox("The second operand is not an Int64.")
                    End If

                    'a.StripZeros()
                    ' a.Round(OALprec)
            End Select
            ' t = DateAndTime.Timer - t

            'comment out the following line to display Float64 in TextBox3
            GoTo skipDbl

            d# = 0
            s3$ = ""
            Try
                a.ExtractDouble(d)
                s3$ = d
            Catch ex As Exception
                s3$ = "out of rng"
            End Try
            TextBox3.Text = s3

skipDbl:

            If -a.GetBottExp > OALprec Then
                a.Round(OALprec)
            End If


            s2$ = ""

            If isApprox Then s2$ = "(approx.) "
            s2$ += olepšaj(a.ToString)

            If CheckBox1.Checked Then

                If TextBox2.TextLength Then lastTextInp = TextBox2.Text

                rtx.AppendText(" " & ComboBox1.Text & vbCrLf &
                    lastTextInp & " = " & vbCrLf & s2)

                rtx.SelectionStart = rtx.TextLength
                rtx.ScrollToCaret()
            End If

            TextBox1.Text = s2
            TextBox2.Text = ""
            izpis()
            'Label3.Text = "Operation took (s): " & t
            Exit Sub

npr:

            a = New Hyper(TextBox2.Text)
            TextBox2.Text = ""
            izpis()
            s4$ = olepšaj(a.ToString)
            TextBox1.Text = s4
            If CheckBox1.Checked Then
                rtx.AppendText(vbCrLf & s4)
                rtx.SelectionStart = rtx.TextLength
                rtx.ScrollToCaret()

            End If


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation)
        End Try

    End Sub

    Sub izpis()
        If IsNothing(a) Then
            Label1.Text = ""
            Label2.Text = ""
        Else
            Label1.Text = "Low Exp: " & a.GetBottExp & "        Buffer size: " & a.BufferSize
            Label2.Text = "High Exp: " & a.GetTopExp
        End If

    End Sub

    Function chkINT(h As Hyper) As Boolean
        If h.GetBottExp Or h.GetTopExp Then
            Return False
        Else
            Return True

        End If
    End Function


    Private Function BigDivide(dividend As Hyper, d As Hyper) As Hyper
        Return dividend * ReciprocalVal(d)
    End Function



    Private Function ReciprocalVal(d As Hyper) As Hyper


        precision% = XLdivPrec1         'number of 64-bit digits to extract. Must be larger than exponent range
        precision2% = XLdivPrec2 'may be zero


        Dim r As New Hyper(precision, 0) ' New Hyper(0, 0)
        Dim bp, r1 As Hyper


        hiExp% = d.FindHighExponent
        lowExp% = d.FindLowExponent '000000000000000000000000000000000000000000000 d.PartSize
        endpos% = precision + lowExp
        lowVal& = d(lowExp)
        '        d(lowestNZ)
        If lowVal And 1 = 0 Then
            ' Debug.WriteLine("even nr")
            'if the least significant bit is 0, then we can help ourselves with dividing/multiplying the dividend, and then the result, by 2^n.

            'lowVal = 1 Or lowVal
            d(lowExp - 1) = 1
            lowVal = 1
            lowExp -= 1
        End If
        mq& = GetMagicNr(lowVal)

        pos% = lowExp '
        d.Round(-pos)
        de% = pos
        d.PartSize = 0
        bp = New Hyper("1")
        pos1% = 0
mainloop:
        ' get the sequence which, when multiplied by divisor, nullifies itself

        r1 = New Hyper(pos1, pos1)
        'r1 = bp * mq
        'r1(pos) = bp(pos1) * mq


        'r1(pos) = bp(pos) * mq
        'r(pos) = r1(pos)
        r1(pos1) = bp(pos1) * mq






        '00000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000    
        'If 0 And 1 = 0 Then
        r(pos1) = r1(pos1)
        'r.Add(r1)
        'Else
        'r.Subtract(r1)
        'r(pos1) = r1(pos1)
        'End If

        bp -= d * r1
        bp.Negate()




        pos1 += 1
        pos += 1

        If pos > endpos Then GoTo nx
        'reciprocal values of large numbers tend to repeat at very large intervals, so we'll be satisfied with our precision                 

        GoTo mainloop


nx:




        r1 = r * d

        'l% = (hiExp - lowExp)
        'pos1 = r1.FindHighExponent - l - 2
        'pos1 = r.FindHighExponent

        'For i% = pos1 To r1.FindHighExponent
        'r1(i) = 0
        'Next


        hi% = r1.FindHighExponent
        ' r.Divide(r1(hi) * mq, 444) ' - r1.PartSize), 22)
        r.Divide(r1(hi), precision2) ' - r1.PartSize), 22)
        r.PartSize = hi + r1.PartSize + r.PartSize + de '.PartSize
        d.PartSize = -de
        Return r
    End Function

    Private Function GetMagicNr&(a&)


        ' Magic number or "reciprocal integer" - GET THE 64-BIT NUMBER WHICH, when multiplied by the lowest digit, gives 1 as the remainder of 2^64               
        ' only for odd numbers

        bt& = 1 'bit tester
        d& = a 'bit mask

        r& = 0 : i& = 0 : r0& = 0



        For i = 0 To 63

            If bt And r Then GoTo skip

            r += d
            r0 = r0 Or bt
skip:
            bt <<= 1 : d <<= 1
        Next

        Return r0


    End Function

    Private ms As New List(Of Hyper)

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Hyper.displayMode = Hyper.displayModeType.inTrueDecimal
        Hyper.ExtraDigitsForConversionFromDecimal = 22
        Hyper.QuotientPrecision = 3333

        ComboBox2.SelectedIndex = 0
        loadSettings()
        For i% = 0 To 5
            ms.Add(New Hyper(0, 0))
        Next


    End Sub

    Private Sub loadSettings()
        Hyper.ExtraDigitsForConversionFromDecimal = My.Settings.ConversionDecPrec
        Hyper.maxDigitsInString = My.Settings.DisplayDigitsDec
        Hyper.QuotientPrecision = My.Settings.QuotientPrec

        If My.Settings.ApplyOverallPrec Then
            OALprec% = My.Settings.OverallPrecision
        Else
            OALprec = 2000000
        End If

        OALprecOnDIV% = My.Settings.QuotientPrec
        XLdivPrec1% = My.Settings.DivPrec1
        XLdivPrec2% = My.Settings.DivPrec2



    End Sub

    Private Sub Form1_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown

        If e.Alt OrElse e.Control Then Exit Sub

        Dim unclear As Boolean = TextBox2.TextLength
        e.Handled = True
        e.SuppressKeyPress = True

        Select Case e.KeyCode
            Case Keys.Oemplus, Keys.Add
                If TextBox2.SelectionStart Then Button1.PerformClick()
                ComboBox1.SelectedIndex = 0
            Case Keys.Subtract, Keys.OemMinus

                With TextBox2

                    txt$ = TextBox2.Text

                    If .SelectionStart Then
                        If UCase(Mid(txt, .SelectionStart, 1)) = "E" Then
                            e.Handled = False
                            e.SuppressKeyPress = False
                            Exit Select
                        End If
                    End If
                End With

                Button1.PerformClick()
                ComboBox1.SelectedIndex = 1

            Case Keys.Multiply, e.Shift And Keys.D8
                If TextBox2.SelectionStart Then Button1.PerformClick()
                ComboBox1.SelectedIndex = 2
            Case Keys.Divide, 191

                If TextBox2.SelectionStart Then Button1.PerformClick()
                ComboBox1.SelectedIndex = 3

            Case Keys.M
                If TextBox2.SelectionStart Then
                    Button1.PerformClick()
                    ComboBox1.SelectedIndex = 4

                Else
                    e.Handled = False
                    e.SuppressKeyPress = False
                    Exit Select

                End If

            Case Keys.Escape
                If unclear Then
                    TextBox2.Text = ""
                Else
                    TextBox1.Clear()
                    a = Nothing
                    b = Nothing
                    izpis()
                End If
            Case Keys.Enter

                Button1.PerformClick()
                Exit Sub

            Case Keys.D0 To Keys.D9, Keys.NumPad0 To Keys.NumPad9, Keys.OemPeriod, Keys.Space,
Keys.Back, Keys.Delete, Keys.Up, Keys.Down, Keys.Left, Keys.Right,
Keys.Home, Keys.End, Keys.E

                e.Handled = False
                e.SuppressKeyPress = False

            Case Keys.Oemcomma
                SendKeys.Send(".")
                Exit Sub


        End Select


    End Sub

    
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If Settings.ShowDialog() = vbOK Then loadSettings()
    End Sub


    Private Sub MSbtn_Click(sender As Object, e As EventArgs) Handles MSbtn.Click
        ms(ComboBox2.SelectedIndex) = a.Clone
        TextBox2.Focus()
    End Sub

    Private Sub MRbtn_Click(sender As Object, e As EventArgs) Handles MRbtn.Click

        TextBox2.Text = ComboBox2.Text
        TextBox2.Focus()
    End Sub

    Private Sub MCbtn_Click(sender As Object, e As EventArgs) Handles MCbtn.Click
        ms(ComboBox2.SelectedIndex) = New Hyper(0, 0)
        TextBox2.Focus()
    End Sub

    
End Class
