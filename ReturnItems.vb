Imports System.Data.SqlClient
Public Class returnitems

    Public Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'TODO: This line of code loads data into the 'ShutterDataSet.rentreturn' table. You can move, or remove it, as needed.
        Me.RentreturnTableAdapter.Fill(Me.ShutterDataSet.rentreturn)

        Me.Refresh()
        DataGridView1.Refresh()
        
        Dim startdate As Date = DateAndTime.Today
        txtdtreturned.Text = startdate
        txtpamount.Text = 10
    End Sub
    Private Sub cmdCalculator_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        On Error GoTo Mali
        Shell("calc.exe", vbNormalFocus)
        Exit Sub
Mali:
        MsgBox("Calculator is not installed in your computer.", vbExclamation)
    End Sub

    Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click
        Me.Hide()
        frmMain.Show()
    End Sub

    Private Sub txtrefid_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtrefid.KeyPress
        
        ' Number validation 
        
        If (Microsoft.VisualBasic.Asc(e.KeyChar) < 48) _
           Or (Microsoft.VisualBasic.Asc(e.KeyChar) > 57) Then
            e.Handled = True
        End If
        If (Microsoft.VisualBasic.Asc(e.KeyChar) = 8) Then
            e.Handled = False
        End If

    End Sub

    Public Sub txtMovie_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtrefid.TextChanged
        
        ' Logic for the return of items by customer and updating the same in database
        
        If (txtrefid.Text <> "") Then
            Dim conn2 As New SqlConnection
            Dim comm As New SqlCommand
            Dim comm2 As New SqlCommand
            Dim comm3 As New SqlCommand
            Dim comm4 As New SqlCommand
            Dim comm5 As New SqlCommand
            Dim comm6 As New SqlCommand
            Dim dr As SqlDataReader
            Dim dr2 As SqlDataReader
            Dim dr3 As SqlDataReader
            Dim dr4 As SqlDataReader
            Dim dr5 As SqlDataReader
            Dim dr6 As SqlDataReader
            ' Dim count As Integer
            conn2.ConnectionString = "Data Source=MCA51\SQLEXPRESS;Initial Catalog=shutter;Integrated Security=True"
            conn2.Open()
            comm.Connection = conn2

            comm.CommandText = " select * from rentreturn where reftransno = '" & Val(txtrefid.Text) & "'"

            dr = comm.ExecuteReader
            If (dr.HasRows) Then
                While dr.Read()
                    lblitemid.Text = ("" & dr(1))
                    lblborrid.Text = ("" & dr(2))
                    lbldtborr.Text = ("" & dr(4))
                    lblduedate.Text = ("" & dr(5))
                    lblrentamt.Text = ("" & dr(8))
                    lblamount.Text = ("" & dr(7))
                    lblitemstatus.Text = ("" & dr(9))
                End While

                dr.Close()

                If txtrefid.Text = "" Then
                    lblcopies.Text = ""
                    lbldtborr.Text = ""
                    lblduedate.Text = ""
                    lblfid.Text = ""
                    lblformat.Text = ""
                    lblitemid.Text = ""
                    lblitemstatus.Text = ""
                    txtpamount.Text = ""
                    lblrentmovie.Text = ""
                End If

                comm2.Connection = conn2
                comm2.CommandText = "select * from items where itemid = '" & Val(lblitemid.Text) & "'"
                dr2 = comm2.ExecuteReader
                While dr2.Read
                    lblrentmovie.Text = ("" & dr2(1))
                End While
                dr2.Close()

                comm3.Connection = conn2
                comm3.CommandText = "select * from membership where membershipid = '" & Val(lblborrid.Text) & "'"
                dr3 = comm3.ExecuteReader
                While dr3.Read
                    lblborrname.Text = ("" & dr3(2))
                End While
                dr3.Close()

                comm4.Connection = conn2
                comm4.CommandText = "select * from items where itemid = '" & Val(lblitemid.Text) & "'"
                dr4 = comm4.ExecuteReader
                While dr4.Read
                    lblfid.Text = ("" & dr4(9))
                End While
                dr4.Close()

                comm5.Connection = conn2
                comm5.CommandText = "select * from format where fid = '" & Val(lblfid.Text) & "'"
                dr5 = comm5.ExecuteReader
                While dr5.Read
                    lblformat.Text = ("" & dr5(1))
                    lblamount.Text = ("" & dr5(3))
                End While
                dr5.Close()

                comm6.Connection = conn2
                comm6.CommandText = "select noofcopies from itemcopies where itemidno = '" & Val(lblitemid.Text) & "'"
                dr6 = comm6.ExecuteReader
                While dr6.Read
                    lblcopies.Text = ("" & dr6(0))
                End While
                dr6.Close()

                Dim dtStartDate As Date
                Dim tsTimeSpan As TimeSpan
                Dim iNumberOfDays As Integer
                '  Try


                dtStartDate = lblduedate.Text
                tsTimeSpan = Now.Subtract(dtStartDate)

                iNumberOfDays = tsTimeSpan.Days

                txtdayspenalty.Text = iNumberOfDays
                If txtdayspenalty.Text <= 0 Then
                    txtdayspenalty.Text = 0
                    lblpenaltyamt.Text = 0
                Else

                    lblpenaltyamt.Text = Val(txtdayspenalty.Text) * Val(txtpamount.Text)
                    lblamount.Text = Val(lblrentamt.Text) + Val(lblpenaltyamt.Text)

                End If

                comm.ExecuteNonQuery()
                comm.CommandText = "commit"

                conn2.Close()

                ' Else
                '    MsgBox("No record Found.")
                '     clear()
            End If
        End If

    End Sub
    Public Sub cmdClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClear.Click
        txtrefid.Text = ""
        lblpenaltyamt.Text = ""
        txtdayspenalty.Text = ""
        lblformat.Text = ""
        lblborrname.Text = ""
        'lblstatus.Text = ""
        lblrentmovie.Text = ""
        lblitemstatus.Text = ""
        lblitemid.Text = ""
        lblduedate.Text = ""
        lbldtborr.Text = ""
        txtdtreturned.Text = ""
    End Sub
    Private Sub clear()
        txtrefid.Text = ""
        lblpenaltyamt.Text = ""
        txtdayspenalty.Text = ""
        lblformat.Text = ""
        lblborrname.Text = ""
        lblrentmovie.Text = ""
        lblitemstatus.Text = ""
        lblitemid.Text = ""
        lblduedate.Text = ""
        lbldtborr.Text = ""
        lblamount.Text = ""
        lblrentamt.Text = ""
        lblcopies.Text = ""
        lblborrid.Text = ""
    End Sub

    Private Sub cmdReturn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReturn.Click
        Dim check As Integer
        Dim conn As SqlConnection
        Dim conn4 As New SqlConnection
        Dim conn8 As New SqlConnection
        Dim cmdretu As New SqlCommand
        Dim ma As Integer
        Dim cmdmember As New SqlCommand
        Dim damember As New SqlDataAdapter
        Dim dsmember As New DataSet
        Dim cmdupdt As SqlCommand
        Dim cmdstat As SqlCommand

        conn4.ConnectionString = "Data Source=MCA51\SQLEXPRESS;Initial Catalog=shutter;Integrated Security=True"
        conn8.ConnectionString = "Data Source=MCA51\SQLEXPRESS;Initial Catalog=shutter;Integrated Security=True"

        If txtrefid.Text <> "" Then

            If MsgBox("Are you sure to return the movie with transaction no.  : " & txtrefid.Text & " ?", MsgBoxStyle.OkCancel, "Delete confirm") = MsgBoxResult.Cancel Then
                ' do nothing
            Else
                conn = GetConnect()
                Try
                    conn.Open()
                    cmdmember = conn.CreateCommand
                    cmdmember.CommandText = "DELETE FROM rentreturn WHERE reftransno ='" & Trim(txtrefid.Text) & "'"
                    check = cmdmember.ExecuteReader.RecordsAffected

                    If check > 0 Then

                        MsgBox("Movie with transaction ID " & Trim(txtrefid.Text) & " returned sucessfully", MsgBoxStyle.OkOnly, "Info Delete Student")

                        lblcopies.Text = Val(lblcopies.Text) + 1
                        conn4.Open()
                        cmdupdt = New SqlCommand("Update itemcopies set noofcopies = '" & lblcopies.Text & "' where itemidno = '" & Val(lblitemid.Text) & "'", conn4)
                        cmdstat = New SqlCommand("Update items set status = 'available' where itemid = '" & Val(lblitemid.Text) & "'", conn4)
                        cmdupdt.ExecuteNonQuery()
                        cmdstat.ExecuteNonQuery()
                        conn4.Close()
                        conn8.Open()
                        cmdretu = New SqlCommand("Insert into [return]([reftransnos],[itemidno],[title],[amount],[dateborr],[dtreturned],[noofdayspenalty],[totalamt],[penaltyamount],[format]) values ('" & txtrefid.Text & "','" & lblitemid.Text & "','" & lblrentmovie.Text & "','" & lblrentamt.Text & "','" & lbldtborr.Text & "','" & txtdtreturned.Text & "','" & txtdayspenalty.Text & "','" & lblamount.Text & "','" & lblpenaltyamt.Text & "','" & lblformat.Text & "')", conn8)
                        ma = cmdretu.ExecuteNonQuery()
                        conn8.Close()
                        clear()

                    Else
                        MsgBox("Movie Transaction With Id " & Trim(txtrefid.Text) & " failed to accept ", MsgBoxStyle.OkOnly, "Info Delete Student")
                    End If

                    Refresh_Form()

                    conn.Close()



                Catch ex As Exception

                    MsgBox("Error: " & ex.Source & ": " & ex.Message, MsgBoxStyle.OkOnly, "Connection Error !!")

                End Try

            End If
        Else
            MsgBox("Fill Id Movie on Id textbox which movie to be returned!!", MsgBoxStyle.OkOnly, "Info Data")

        End If
    End Sub

    Private Sub Refresh_Form()
        Dim conn As SqlConnection
        Dim cmdStudent As New SqlCommand
        Dim daStudent As New SqlDataAdapter
        Dim dsStudent As New DataSet
        Dim dtStudent As New DataTable
        conn = GetConnect()
        Try
            cmdStudent = conn.CreateCommand
            cmdStudent.CommandText = "SELECT * FROM rentreturn"
            daStudent.SelectCommand = cmdStudent
            daStudent.Fill(dsStudent, "rentreturn")
            DataGridView1.DataSource = dsStudent
            DataGridView1.DataMember = "rentreturn"
            DataGridView1.ReadOnly = True
        Catch ex As Exception
            MsgBox("Error: " & ex.Source & ": " & ex.Message, MsgBoxStyle.OKOnly, "Connection Error !!")
        End Try
    End Sub
  
End Class
