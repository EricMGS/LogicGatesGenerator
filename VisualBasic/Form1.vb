Public Class Form1
	Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

	End Sub

	Private Sub gera_circuito()
		Select Case (cmb_a.Text)
			Case "A" : img_not_a.Image = Nothing
			Case Else : img_not_a.Image = VisualBasic.My.Resources.Resources._not
		End Select

		Select Case (cmb_b.Text)
			Case "B" : img_not_b.Image = Nothing
			Case Else : img_not_b.Image = VisualBasic.My.Resources.Resources._not
		End Select

		Select Case (cmb_c.Text)
			Case "C" : img_not_c.Image = Nothing
			Case Else : img_not_c.Image = VisualBasic.My.Resources.Resources._not
		End Select

		Select Case (cmb_d.Text)
			Case "D" : img_not_d.Image = Nothing
			Case Else : img_not_d.Image = VisualBasic.My.Resources.Resources._not
		End Select

		Select Case (cmb_exp1.Text)
			Case "and" : img_1.Image = VisualBasic.My.Resources.Resources._and
			Case "nand" : img_1.Image = VisualBasic.My.Resources.Resources.nand
			Case "or" : img_1.Image = VisualBasic.My.Resources.Resources._or
			Case "nor" : img_1.Image = VisualBasic.My.Resources.Resources.nor
			Case "xor" : img_1.Image = VisualBasic.My.Resources.Resources.xor3
			Case "xnor" : img_1.Image = VisualBasic.My.Resources.Resources.xnor
		End Select

		Select Case (cmb_exp2.Text)
			Case "and" : img_2.Image = VisualBasic.My.Resources.Resources._and
			Case "nand" : img_2.Image = VisualBasic.My.Resources.Resources.nand
			Case "or" : img_2.Image = VisualBasic.My.Resources.Resources._or
			Case "nor" : img_2.Image = VisualBasic.My.Resources.Resources.nor
			Case "xor" : img_2.Image = VisualBasic.My.Resources.Resources.xor3
			Case "xnor" : img_2.Image = VisualBasic.My.Resources.Resources.xnor
		End Select

		Select Case (cmb_exp3.Text)
			Case "and" : img_3.Image = VisualBasic.My.Resources.Resources._and
			Case "nand" : img_3.Image = VisualBasic.My.Resources.Resources.nand
			Case "or" : img_3.Image = VisualBasic.My.Resources.Resources._or
			Case "nor" : img_3.Image = VisualBasic.My.Resources.Resources.nor
			Case "xor" : img_3.Image = VisualBasic.My.Resources.Resources.xor3
			Case "xnor" : img_3.Image = VisualBasic.My.Resources.Resources.xnor
		End Select
	End Sub

	Private Sub gera_expressao()
		lbl_exp.Text = "(" & cmb_a.Text & " " & cmb_exp1.Text & " " & cmb_b.Text & ") " &
					   cmb_exp2.Text & " (" & cmb_c.Text & " " & cmb_exp3.Text & " " & cmb_d.Text & ")"

	End Sub

	Private Sub gera_tabela()
		Dim expressao As String

		expressao = exp()
		expressao = Replace(expressao, "A", "0")
		expressao = Replace(expressao, "B", "0")
		expressao = Replace(expressao, "C", "0")
		expressao = Replace(expressao, "D", "0")
		lbl_s1.Text = Eval(expressao)

		expressao = exp()
		expressao = Replace(expressao, "A", "0")
		expressao = Replace(expressao, "B", "0")
		expressao = Replace(expressao, "C", "0")
		expressao = Replace(expressao, "D", "1")
		lbl_s2.Text = Eval(expressao)

		expressao = exp()
		expressao = Replace(expressao, "A", "0")
		expressao = Replace(expressao, "B", "0")
		expressao = Replace(expressao, "C", "1")
		expressao = Replace(expressao, "D", "0")
		lbl_s3.Text = Eval(expressao)

		expressao = exp()
		expressao = Replace(expressao, "A", "0")
		expressao = Replace(expressao, "B", "0")
		expressao = Replace(expressao, "C", "1")
		expressao = Replace(expressao, "D", "1")
		lbl_s4.Text = Eval(expressao)

		expressao = exp()
		expressao = Replace(expressao, "A", "0")
		expressao = Replace(expressao, "B", "1")
		expressao = Replace(expressao, "C", "0")
		expressao = Replace(expressao, "D", "0")
		lbl_s5.Text = Eval(expressao)

		expressao = exp()
		expressao = Replace(expressao, "A", "0")
		expressao = Replace(expressao, "B", "1")
		expressao = Replace(expressao, "C", "0")
		expressao = Replace(expressao, "D", "1")
		lbl_s6.Text = Eval(expressao)

		expressao = exp()
		expressao = Replace(expressao, "A", "0")
		expressao = Replace(expressao, "B", "1")
		expressao = Replace(expressao, "C", "1")
		expressao = Replace(expressao, "D", "0")
		lbl_s7.Text = Eval(expressao)

		expressao = exp()
		expressao = Replace(expressao, "A", "0")
		expressao = Replace(expressao, "B", "1")
		expressao = Replace(expressao, "C", "1")
		expressao = Replace(expressao, "D", "1")
		lbl_s8.Text = Eval(expressao)

		expressao = exp()
		expressao = Replace(expressao, "A", "1")
		expressao = Replace(expressao, "B", "0")
		expressao = Replace(expressao, "C", "0")
		expressao = Replace(expressao, "D", "0")
		lbl_s9.Text = Eval(expressao)

		expressao = exp()
		expressao = Replace(expressao, "A", "1")
		expressao = Replace(expressao, "B", "0")
		expressao = Replace(expressao, "C", "0")
		expressao = Replace(expressao, "D", "1")
		lbl_s10.Text = Eval(expressao)

		expressao = exp()
		expressao = Replace(expressao, "A", "1")
		expressao = Replace(expressao, "B", "0")
		expressao = Replace(expressao, "C", "1")
		expressao = Replace(expressao, "D", "0")
		lbl_s11.Text = Eval(expressao)

		expressao = exp()
		expressao = Replace(expressao, "A", "1")
		expressao = Replace(expressao, "B", "0")
		expressao = Replace(expressao, "C", "1")
		expressao = Replace(expressao, "D", "1")
		lbl_s12.Text = Eval(expressao)

		expressao = exp()
		expressao = Replace(expressao, "A", "1")
		expressao = Replace(expressao, "B", "1")
		expressao = Replace(expressao, "C", "0")
		expressao = Replace(expressao, "D", "0")
		lbl_s13.Text = Eval(expressao)

		expressao = exp()
		expressao = Replace(expressao, "A", "1")
		expressao = Replace(expressao, "B", "1")
		expressao = Replace(expressao, "C", "0")
		expressao = Replace(expressao, "D", "1")
		lbl_s14.Text = Eval(expressao)

		expressao = exp()
		expressao = Replace(expressao, "A", "1")
		expressao = Replace(expressao, "B", "1")
		expressao = Replace(expressao, "C", "1")
		expressao = Replace(expressao, "D", "0")
		lbl_s15.Text = Eval(expressao)

		expressao = exp()
		expressao = Replace(expressao, "A", "1")
		expressao = Replace(expressao, "B", "1")
		expressao = Replace(expressao, "C", "1")
		expressao = Replace(expressao, "D", "1")
		lbl_s16.Text = Eval(expressao)

	End Sub

	Private Function exp() As String
		Dim str1 As String, str2 As String, str3 As String

		str1 = "(" & cmb_a.Text & " " & cmb_exp1.Text & " " & cmb_b.Text & ") "
		str2 = cmb_exp2.Text
		str3 = " (" & cmb_c.Text & " " & cmb_exp3.Text & " " & cmb_d.Text & ")"

		If cmb_exp1.Text = "nand" Then
			str1 = Replace(str1, "nand", "and")
			str1 = "not" & str1
		ElseIf cmb_exp1.Text = "nor" Then
			str1 = Replace(str1, "nor", "or")
			str1 = "not" & str1
		ElseIf cmb_exp1.Text = "xnor" Then
			str1 = Replace(str1, "xnor", "xor")
			str1 = "not" & str1
		End If

		If cmb_exp3.Text = "nand" Then
			str3 = Replace(str3, "nand", "and")
			str3 = " not" & str3
		ElseIf cmb_exp3.Text = "nor" Then
			str3 = Replace(str3, "nor", "or")
			str3 = " not" & str3
		ElseIf cmb_exp3.Text = "xnor" Then
			str3 = Replace(str3, "xnor", "xor")
			str3 = " not" & str3
		End If

		If cmb_exp2.Text = "nand" Then
			str2 = "and"
			exp = "not (" & str1 & str2 & str3 & ")"
		ElseIf cmb_exp2.Text = "nor" Then
			str2 = "or"
			exp = "not (" & str1 & str2 & str3 & ")"
		ElseIf cmb_exp2.Text = "xnor" Then
			str2 = "xor"
			exp = "not (" & str1 & str2 & str3 & ")"
		Else
			exp = str1 & str2 & str3
		End If


	End Function

	Private Function Eval(strEvalContent As String)
		Dim result As Integer

		With CreateObject("ScriptControl")
			.Language = "VBScript"
			strEvalContent = Replace(strEvalContent, "~", "not ")

			result = .Eval(strEvalContent)
			If result = -1 Then
				result = 1
			ElseIf result = -2 Then
				result = 0
			End If
			Eval = result
		End With
	End Function

	Private Sub btn_gerar_Click(sender As Object, e As EventArgs) Handles btn_gerar.Click
		gera_circuito()
		gera_expressao()
		gera_tabela()
	End Sub
End Class