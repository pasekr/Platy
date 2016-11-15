Public Class Hlavni
    Private Sub Hlavni_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'TODO: This line of code loads data into the 'DS_pers.PPV' table. You can move, or remove it, as needed.
        obnovSpojeni()
        Me.TA_ciselnik_TypPraxe.Fill(Me.DS_ciselnik_TypPraxe.TYPPRAXE)
        Me.TA_ciselnik_TypLhuty.Fill(Me.DS_ciselnik_TypLhuty.TYPLHUTY)
        Me.TA_ciselnik_VzdDosa.Fill(Me.DS_ciselnik_VzdDosa.VZD_DOSA)
        Me.TA_ciselnik_Duchody.Fill(Me.DS_ciselnik_Duchody.DUCHODY)
        Me.TA_ciselnik_ZdrPoj.Fill(Me.DS_ciselnik_ZdrPoj.ZDRPOJ)
        Me.TA_ciselnik_DruhPracRezim.Fill(Me.DS_ciselnik_DruhPracRezim.DRUHPRACREZIM)
        Me.TA_ciselnik_DruhPracDoba.Fill(Me.DS_ciselnik_DruhPracDoba.DRUHPRACDOBA)
        Me.TA_ciselnik_PPV.Fill(Me.DS_ciselnik_PPV.PPV)
        Me.TA_pers.Fill(Me.DS_pers.PERS)
        Me.TA_uvazky.Fill(Me.DS_uvazky.UVAZKY)
        Me.TA_skola.Fill(Me.DS_skola.SKOLA)
        Me.TA_praxe.Fill(Me.DS_praxe.PRAXE)
        Me.TA_lhuty.Fill(Me.DS_lhuty.LHUTY)

        TSB_pers.PerformClick()

    End Sub

    Private Sub DGV_pers_SelectionChanged(sender As Object, e As EventArgs) Handles DGV_pers.SelectionChanged
        If (BS_pers.Current("DAT_KONEC_PP").ToString <> "") Then
            DateTimePicker3.Format = DateTimePickerFormat.Short
        Else
            DateTimePicker3.CustomFormat = " "
            DateTimePicker3.Format = DateTimePickerFormat.Custom
        End If
        NastavFiltr(BS_pers.Current("ID").ToString)
    End Sub

    Private Sub NastavFiltr(str_Pers_ID As String)
        If TSB_praxe.Checked Then
            BS_praxe.Filter = "ID_PRAXE = " & str_Pers_ID
        End If
        If TSB_vzdelani.Checked Then
            BS_skola.Filter = "ID_SKOLA = " & str_Pers_ID
        End If
        If TSB_pers.Checked Then
            BS_uvazky.Filter = "ID_OSOBY = " & str_Pers_ID
        End If
        If TSB_lhuty.Checked Then
            BS_lhuty.Filter = SlozFiltrLhuty(BS_pers.Current("ID").ToString)
        End If
    End Sub

    Private Sub DateTimePicker3_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker3.ValueChanged
        If (DateTimePicker3.ToString <> "") Then
            DateTimePicker3.Format = DateTimePickerFormat.Short
        Else
            DateTimePicker3.CustomFormat = " "
            DateTimePicker3.Format = DateTimePickerFormat.Custom
        End If
    End Sub

    Private Sub TSB_vzdelani_Click(sender As Object, e As EventArgs) Handles TSB_vzdelani.Click

        BS_skola.Filter = "ID_SKOLA = " & BS_pers.Current("ID").ToString

        SC_hlavni.Visible = True
        DGV_vzdelani.Visible = True
        DGV_pers.Visible = True
        DGV_lhuty.Visible = False
        DGV_pers_ost.Visible = False
        DGV_praxe.Visible = False

        P_plat.Visible = False

        TC_Zamestnanci.Visible = False
        TS_levy_dolni.Visible = True
        TS_pravy_dolni.Visible = True

        TS_pravy_dolni_sep1.Visible = False
        TSB_lhuty_filtr.Visible = False

        'SC_hlavni.Panel2.Visible = True

        ReportViewer1.Visible = False

        TSB_tisk.Checked = False
        TSB_praxe.Checked = False
        TSB_lhuty.Checked = False
        TSB_plat.Checked = False
        TSB_pers.Checked = False
        TSB_pers_ost.Checked = False

        ZmenVelikostPanelu()
    End Sub

    Private Sub TSB_praxe_Click(sender As Object, e As EventArgs) Handles TSB_praxe.Click

        BS_praxe.Filter = "ID_PRAXE = " & BS_pers.Current("ID").ToString

        SC_hlavni.Visible = True
        DGV_praxe.Visible = True
        DGV_pers.Visible = True
        DGV_vzdelani.Visible = False
        DGV_lhuty.Visible = False
        DGV_pers_ost.Visible = False

        P_plat.Visible = False

        TC_Zamestnanci.Visible = False
        TS_levy_dolni.Visible = True
        TS_pravy_dolni.Visible = True

        TS_pravy_dolni_sep1.Visible = False
        TSB_lhuty_filtr.Visible = False

        'SC_hlavni.Panel2.Visible = True

        ReportViewer1.Visible = False

        TSB_tisk.Checked = False
        TSB_vzdelani.Checked = False
        TSB_lhuty.Checked = False
        TSB_plat.Checked = False
        TSB_pers.Checked = False
        TSB_pers_ost.Checked = False

        ZmenVelikostPanelu()
    End Sub

    Private Sub TSB_pers_Click(sender As Object, e As EventArgs) Handles TSB_pers.Click

        BS_uvazky.Filter = "ID_OSOBY = " & BS_pers.Current("ID").ToString

        SC_hlavni.Visible = True
        DGV_pers_ost.Visible = False
        DGV_pers.Visible = True
        DGV_praxe.Visible = False
        DGV_vzdelani.Visible = False
        DGV_lhuty.Visible = False

        P_plat.Visible = False

        TC_Zamestnanci.Visible = True
        TS_levy_dolni.Visible = True
        TS_pravy_dolni.Visible = False

        TS_pravy_dolni_sep1.Visible = False
        TSB_lhuty_filtr.Visible = False

        'SC_hlavni.Panel2.Visible = True

        ReportViewer1.Visible = False

        TSB_tisk.Checked = False
        TSB_vzdelani.Checked = False
        TSB_praxe.Checked = False
        TSB_lhuty.Checked = False
        TSB_plat.Checked = False
        TSB_pers_ost.Checked = False

        ZmenVelikostPanelu()
    End Sub

    Private Sub TSB_pers_ost_Click(sender As Object, e As EventArgs) Handles TSB_pers_ost.Click

        SC_hlavni.Visible = True
        DGV_pers_ost.Visible = True
        DGV_pers.Visible = True
        DGV_praxe.Visible = False
        DGV_vzdelani.Visible = False
        DGV_lhuty.Visible = False

        P_plat.Visible = False

        TC_Zamestnanci.Visible = False
        TS_levy_dolni.Visible = True
        TS_pravy_dolni.Visible = True

        TS_pravy_dolni_sep1.Visible = False
        TSB_lhuty_filtr.Visible = False

        'SC_hlavni.Panel2.Visible = True

        ReportViewer1.Visible = False

        TSB_tisk.Checked = False
        TSB_vzdelani.Checked = False
        TSB_praxe.Checked = False
        TSB_lhuty.Checked = False
        TSB_plat.Checked = False
        TSB_pers.Checked = False

        ZmenVelikostPanelu()
    End Sub

    Private Sub TSB_lhuty_Click(sender As Object, e As EventArgs) Handles TSB_lhuty.Click

        BS_lhuty.Filter = SlozFiltrLhuty(BS_pers.Current("ID").ToString)

        SC_hlavni.Visible = True
        DGV_lhuty.Visible = True
        DGV_pers.Visible = True
        DGV_pers_ost.Visible = False
        DGV_praxe.Visible = False
        DGV_vzdelani.Visible = False
        DGV_pers_ost.Visible = False

        P_plat.Visible = False

        TC_Zamestnanci.Visible = False
        TS_levy_dolni.Visible = True
        TS_pravy_dolni.Visible = True

        TS_pravy_dolni_sep1.Visible = True
        TSB_lhuty_filtr.Visible = True

        'SC_hlavni.Panel2.Visible = True

        ReportViewer1.Visible = False

        TSB_tisk.Checked = False
        TSB_vzdelani.Checked = False
        TSB_praxe.Checked = False
        TSB_plat.Checked = False
        TSB_pers_ost.Checked = False
        TSB_pers.Checked = False

        ZmenVelikostPanelu()

    End Sub

    Private Sub TSB_plat_Click(sender As Object, e As EventArgs) Handles TSB_plat.Click

        SC_hlavni.Visible = True
        DGV_pers.Visible = True
        DGV_pers_ost.Visible = False
        DGV_praxe.Visible = False
        DGV_vzdelani.Visible = False
        DGV_pers_ost.Visible = False
        DGV_lhuty.Visible = False

        P_plat.Visible = True

        TC_Zamestnanci.Visible = False
        TS_levy_dolni.Visible = True
        TS_pravy_dolni.Visible = True

        TS_pravy_dolni_sep1.Visible = False
        TSB_lhuty_filtr.Visible = False

        'SC_hlavni.Panel2.Visible = True

        ReportViewer1.Visible = False

        TSB_tisk.Checked = False
        TSB_vzdelani.Checked = False
        TSB_praxe.Checked = False
        TSB_lhuty.Checked = False
        TSB_pers_ost.Checked = False
        TSB_pers.Checked = False

        ZmenVelikostPanelu()
    End Sub

    Private Sub TSB_tisk_Click(sender As Object, e As EventArgs) Handles TSB_tisk.Click

        TSB_tisk.Checked = True
        'DGV_pers.Visible = False
        'TS_levy_dolni.Visible = False
        SC_hlavni.Visible = False

        TSB_pers.Checked = False
        TSB_pers_ost.Checked = False
        TSB_plat.Checked = False
        TSB_vzdelani.Checked = False
        TSB_praxe.Checked = False
        TSB_lhuty.Checked = False

        'ZmenVelikostPanelu()

        ReportViewer1.Visible = True
        Me.ReportViewer1.RefreshReport()
    End Sub

    Private Sub Hlavni_SizeChanged(sender As Object, e As EventArgs) Handles Me.SizeChanged
        ZmenVelikostPanelu()
    End Sub

    Private Sub ZmenVelikostPanelu()
        Dim min As Int32 = Me.SC_hlavni.Panel1MinSize
        Dim max As Int32 = Me.SC_hlavni.Width - Me.SC_hlavni.Panel2MinSize
        Dim distance As Int32

        If TSB_praxe.Checked Then
            distance = Me.Width - Me.Width * 2 / 3
            If min <= distance And distance <= max Then
                SC_hlavni.SplitterDistance = Me.Width - Me.Width * 2 / 3
            End If
        ElseIf TSB_vzdelani.Checked Then
            distance = Me.Width - Me.Width * 2 / 3
            If min <= distance And distance <= max Then
                SC_hlavni.SplitterDistance = Me.Width - Me.Width * 2 / 3
            End If
        ElseIf TSB_pers_ost.Checked Then
            distance = Me.Width - Me.Width * 1 / 2
            If min <= distance And distance <= max Then
                SC_hlavni.SplitterDistance = Me.Width - Me.Width * 1 / 2
            End If
        ElseIf TSB_lhuty.Checked Then
            distance = Me.Width - Me.Width * 2 / 3
            If min <= distance And distance <= max Then
                SC_hlavni.SplitterDistance = Me.Width - Me.Width * 2 / 3
            End If
        ElseIf TSB_plat.Checked Then
            distance = Me.Width - 450
            If min <= distance And distance <= max Then
                SC_hlavni.SplitterDistance = Me.Width - 450
            End If
        ElseIf TSB_tisk.Checked Then
            SC_hlavni.SplitterDistance = SC_hlavni.Width
        ElseIf TSB_pers.Checked Then
            distance = Me.Width - 450
            If min <= distance And distance <= max Then
                SC_hlavni.SplitterDistance = Me.Width - 450
            End If
        End If
    End Sub

    Private Sub TSB_lhuty_filtr_DropDownItemClicked(sender As Object, e As ToolStripItemClickedEventArgs) Handles TSB_lhuty_filtr.DropDownItemClicked
        DirectCast(e.ClickedItem, System.Windows.Forms.ToolStripMenuItem).Checked = Not DirectCast(e.ClickedItem, System.Windows.Forms.ToolStripMenuItem).Checked
        BS_lhuty.Filter = SlozFiltrLhuty(BS_pers.Current("ID").ToString)
    End Sub

    Private Function SlozFiltrLhuty(str_Pers_ID As String) As String
        SlozFiltrLhuty = ""
        For Each menu As System.Windows.Forms.ToolStripMenuItem In TSB_lhuty_filtr.DropDownItems
            If menu.Checked Then
                Select Case menu.Name.ToString
                    Case "TSMI_Zdrav"
                        SlozFiltrLhuty = SlozFiltrLhuty & IIf(Len(SlozFiltrLhuty) = 0, "KOD = 'ZDRAV'", " OR KOD = 'ZDRAV'")
                    Case "TSMI_BOZP"
                        SlozFiltrLhuty = SlozFiltrLhuty & IIf(Len(SlozFiltrLhuty) = 0, "KOD = 'BOZP'", " OR KOD = 'BOZP'")
                    Case "TSMI_SIPVZ"
                        SlozFiltrLhuty = SlozFiltrLhuty & IIf(Len(SlozFiltrLhuty) = 0, "KOD = 'SIPVZ'", " OR KOD = 'SIPVZ'")
                    Case "TSMI_MDRD"
                        SlozFiltrLhuty = SlozFiltrLhuty & IIf(Len(SlozFiltrLhuty) = 0, "KOD = 'MD/RD'", " OR KOD = 'MD/RD'")
                    Case "TSMI_Ostatni"
                        SlozFiltrLhuty = SlozFiltrLhuty & IIf(Len(SlozFiltrLhuty) = 0, "KOD = 'SKOL'", " OR KOD = 'SKOL'")
                End Select
            End If
        Next
        SlozFiltrLhuty = "ID_LHUTY = " & str_Pers_ID & IIf(Len(SlozFiltrLhuty) = 0, "", " AND (" & SlozFiltrLhuty & ")")
    End Function

    Private Sub obnovSpojeni()
        Dim con As String = My.Settings.persConnectionString
        Dim appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)
        con = con.Replace("{path}", appDataPath & "\Platy\Organizace\ZAKLADNI\Databaze\pers.mdb")
        My.Settings.Item("persConnectionString") = con
        con = My.Settings.praxeConnectionString
        con = con.Replace("{path}", appDataPath & "\Platy\Organizace\ZAKLADNI\Databaze\praxe.mdb")
        My.Settings.Item("praxeConnectionString") = con
        con = My.Settings.skolaConnectionString
        con = con.Replace("{path}", appDataPath & "\Platy\Organizace\ZAKLADNI\Databaze\skola.mdb")
        My.Settings.Item("skolaConnectionString") = con
        con = My.Settings.ciselnikConnectionString
        con = con.Replace("{path}", appDataPath & "\Platy\Organizace\ZAKLADNI\Databaze\ciselnik.mdb")
        My.Settings.Item("ciselnikConnectionString") = con
        con = My.Settings.lhutyConnectionString
        con = con.Replace("{path}", appDataPath & "\Platy\Organizace\ZAKLADNI\Databaze\lhuty.mdb")
        My.Settings.Item("lhutyConnectionString") = con
        con = My.Settings.uvazkyConnectionString
        con = con.Replace("{path}", appDataPath & "\Platy\Organizace\ZAKLADNI\Databaze\uvazky.mdb")
        My.Settings.Item("uvazkyConnectionString") = con
    End Sub

    Private Sub DGV_lhuty_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles DGV_lhuty.DataError

    End Sub

    Private Sub DGV_vzdelani_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles DGV_vzdelani.DataError

    End Sub

    Private Sub DGV_praxe_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles DGV_praxe.DataError

    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LL_Resk.LinkClicked
        System.Diagnostics.Process.Start("IExplore.exe", "http://resk.cz")
    End Sub

    Private Sub ToolStripButton4_Click(sender As Object, e As EventArgs) Handles ToolStripButton4.Click
        Dim psd As New PageSetupDialog
        psd.PageSettings = New System.Drawing.Printing.PageSettings
        psd.PrinterSettings = New System.Drawing.Printing.PrinterSettings
        Dim result As DialogResult = psd.ShowDialog()
        If (result = DialogResult.OK) Then
            Dim results() As Object = New Object() _
                {psd.PageSettings.Margins,
                 psd.PageSettings.PaperSize,
                 psd.PageSettings.Landscape,
                 psd.PrinterSettings.PrinterName,
                 psd.PrinterSettings.PrintRange}
           ' ListBox1.Items.AddRange(results)
        End If
    End Sub

    Private Sub TLB_Zamestnanec_novy_Click(sender As Object, e As EventArgs) Handles TLB_Zamestnanec_novy.Click
        Dim newRow = CType(Me.DS_pers.PERS.NewRow(), DS_pers.PERSRow)
        newRow.PPV = 1
        newRow.TYP_POZ = 0
        Me.DS_pers.PERS.Rows.Add(newRow)
        Me.TA_pers.Update(DS_pers.PERS)
        Me.DGV_pers.MultiSelect = False
        Me.DGV_pers.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        Me.DGV_pers.Rows(Me.DGV_pers.RowCount - 1).Selected = True
        Me.DGV_pers.CurrentCell = Me.DGV_pers.Rows(Me.DGV_pers.RowCount - 1).Cells(0)
        Me.DGV_pers.FirstDisplayedScrollingRowIndex = Me.DGV_pers.RowCount - 1
        NastavFiltr(BS_pers.Current("ID").ToString)
    End Sub

End Class

