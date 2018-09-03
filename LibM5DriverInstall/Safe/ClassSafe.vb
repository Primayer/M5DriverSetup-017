Imports System.Windows.Forms

Public Class Safe

    Class TreeView
        Class ClassTreeData
            Public Nodo As TreeNode
        End Class




        Private Delegate Sub ExpandAllDel(ByVal t As System.Windows.Forms.TreeView)
        Public Shared Sub ExpandAll(ByVal t As System.Windows.Forms.TreeView)
            Try
                If t IsNot Nothing Then
                    If t.InvokeRequired Then
                        t.Invoke(New Safe.TreeView.ExpandAllDel(AddressOf Safe.TreeView.ExpandAll), t)
                    Else
                        t.ExpandAll()
                    End If
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message + vbCrLf + ex.StackTrace, "Safe.TreeView.ExpandAll", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub


        Private Delegate Sub NodesClearDel(ByVal t As System.Windows.Forms.TreeView)
        Public Shared Sub NodesClear(ByVal t As System.Windows.Forms.TreeView)
            Try
                If t IsNot Nothing Then
                    If t.InvokeRequired Then
                        t.Invoke(New Safe.TreeView.NodesClearDel(AddressOf Safe.TreeView.NodesClear), t)
                    Else
                        t.Nodes.Clear()
                    End If
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message + vbCrLf + ex.StackTrace, "Safe.TreeView.NodesClear", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        Private Delegate Sub NodesAddDel(ByVal s As String, ByVal tag As String, ByVal nodoPadre As System.Windows.Forms.TreeNode, ByRef treeData As ClassTreeData)
        Public Shared Sub NodesAdd(ByVal text As String, ByVal tag As String, ByVal nodoPadre As System.Windows.Forms.TreeNode, ByRef treeData As ClassTreeData)
            Try
                If nodoPadre IsNot Nothing Then
                    If nodoPadre.TreeView.InvokeRequired Then
                        nodoPadre.TreeView.Invoke(New Safe.TreeView.NodesAddDel(AddressOf Safe.TreeView.NodesAdd), text, tag, nodoPadre, treeData)
                    Else
                        treeData.Nodo = nodoPadre.Nodes.Add(text)
                        treeData.Nodo.Tag = tag
                    End If
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message + vbCrLf + ex.StackTrace, "Safe.TreeView.NodesAdd", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        Private Delegate Sub TreeNodesClearDel(t As System.Windows.Forms.TreeView, ByVal n As System.Windows.Forms.TreeNode)
        Public Shared Sub TreeNodesClear(t As System.Windows.Forms.TreeView, ByVal n As System.Windows.Forms.TreeNode)
            Try
                If t IsNot Nothing AndAlso n IsNot Nothing Then
                    If t.InvokeRequired Then
                        t.Invoke(New Safe.TreeView.TreeNodesClearDel(AddressOf Safe.TreeView.TreeNodesClear), t, n)
                    Else
                        n.Nodes.Clear()
                    End If
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message + vbCrLf + ex.StackTrace, "Safe.TreeView.TreeNodesClear", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        Private Delegate Sub TreeAddImgDel(ByVal text As String, ByVal treeW As System.Windows.Forms.TreeView, ByVal ImgIndex As Integer, ByRef treeData As ClassTreeData)
        Public Shared Sub TreeAddImg(ByVal text As String, ByVal treeW As System.Windows.Forms.TreeView, ByVal ImgIndex As Integer, ByRef treeData As ClassTreeData)
            Try
                If treeW IsNot Nothing Then
                    If treeW.InvokeRequired Then
                        treeW.Invoke(New Safe.TreeView.TreeAddImgDel(AddressOf Safe.TreeView.TreeAddImg), text, treeW, ImgIndex, treeData)
                    Else
                        treeData.Nodo = treeW.Nodes.Add("", text, ImgIndex, ImgIndex)
                    End If
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message + vbCrLf + ex.StackTrace, "Safe.TreeView.TreeAddImg", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        Private Delegate Sub NodesAddImgDel(key As String, ByVal text As String, ByVal nodoPadre As System.Windows.Forms.TreeNode, ByVal ImgIndex As Integer, ByRef treeData As ClassTreeData)
        Public Shared Sub NodesAddImg(key As String, ByVal text As String, ByVal nodoPadre As System.Windows.Forms.TreeNode, ByVal ImgIndex As Integer, ByRef treeData As ClassTreeData)
            Try
                If nodoPadre IsNot Nothing Then
                    If nodoPadre.TreeView.InvokeRequired Then
                        nodoPadre.TreeView.Invoke(New Safe.TreeView.NodesAddImgDel(AddressOf Safe.TreeView.NodesAddImg), key, text, nodoPadre, ImgIndex, treeData)
                    Else
                        treeData.Nodo = nodoPadre.Nodes.Add(key, text, ImgIndex, ImgIndex)
                    End If
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message + vbCrLf + ex.StackTrace, "Safe.TreeView.NodesAddImg", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        Private Delegate Sub TreeNodesAddDel(ByVal s As String, ByVal tag As String, ByVal tree As System.Windows.Forms.TreeView, ByRef treeData As ClassTreeData)
        Public Shared Sub TreeNodesAdd(ByVal text As String, ByVal tag As String, ByVal tree As System.Windows.Forms.TreeView, ByRef treeData As ClassTreeData)
            Try
                If tree IsNot Nothing Then
                    If tree.InvokeRequired Then
                        tree.Invoke(New Safe.TreeView.TreeNodesAddDel(AddressOf Safe.TreeView.TreeNodesAdd), text, tag, tree, treeData)
                    Else
                        treeData.Nodo = tree.Nodes.Add(text)
                        treeData.Nodo.Tag = tag
                    End If
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message + vbCrLf + ex.StackTrace, "Safe.TreeView.NodesAdd", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        Private Delegate Sub NodeTagDel(ByVal tag As String, ByVal n As System.Windows.Forms.TreeNode)
        Public Shared Sub NodeTag(ByVal tag As String, ByVal n As System.Windows.Forms.TreeNode)
            Try
                If n IsNot Nothing Then
                    If n.TreeView.InvokeRequired Then
                        n.TreeView.Invoke(New Safe.TreeView.NodeTagDel(AddressOf Safe.TreeView.NodeTag), tag, n)
                    Else
                        n.Tag = tag
                    End If
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message + vbCrLf + ex.StackTrace, "Safe.TreeView.NodeTag", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        Private Delegate Sub ImageIndexDel(ByVal n As System.Windows.Forms.TreeNode, ByVal index As Integer, ByVal selectedImageindex As Integer)
        Public Shared Sub ImageIndex(ByVal n As System.Windows.Forms.TreeNode, ByVal imageIndex As Integer, ByVal selectedImageindex As Integer)
            Try
                If n IsNot Nothing Then
                    If n.TreeView.InvokeRequired Then
                        n.TreeView.Invoke(New Safe.TreeView.ImageIndexDel(AddressOf Safe.TreeView.ImageIndex), n, imageIndex, selectedImageindex)
                    Else
                        n.ImageIndex = imageIndex
                        n.SelectedImageIndex = selectedImageindex
                    End If
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message + vbCrLf + ex.StackTrace, "Safe.TreeView.ImageIndex", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub
    End Class


    Class Tabpage
        Private Delegate Sub RemoveTabDel(ByVal tabctrl As System.Windows.Forms.TabControl, ByVal tabToRemove As System.Windows.Forms.TabPage)
        Public Shared Sub RemoveTab(ByVal tabctrl As System.Windows.Forms.TabControl, ByVal tabToRemove As System.Windows.Forms.TabPage)
            Try
                If tabctrl IsNot Nothing Then
                    If tabctrl.InvokeRequired Then
                        tabctrl.Invoke(New Safe.Tabpage.RemoveTabDel(AddressOf Safe.Tabpage.RemoveTab), tabctrl, tabToRemove)
                    Else
                        tabctrl.TabPages.Remove(tabToRemove)
                    End If
                End If

            Catch ex As Exception
            End Try
        End Sub

        Private Delegate Sub InsertTabDel(ByVal tabctrl As System.Windows.Forms.TabControl, ByVal tabToRemove As System.Windows.Forms.TabPage, ByVal indice As Integer)
        Public Shared Sub InsertTab(ByVal tabctrl As System.Windows.Forms.TabControl, ByVal tabToInsert As System.Windows.Forms.TabPage, ByVal indice As Integer)
            Try
                If tabctrl IsNot Nothing Then
                    If tabctrl.InvokeRequired Then
                        tabctrl.Invoke(New Safe.Tabpage.InsertTabDel(AddressOf Safe.Tabpage.InsertTab), tabctrl, tabToInsert, indice)
                    Else
                        tabctrl.TabPages.Insert(indice, tabToInsert)
                    End If
                End If

            Catch ex As Exception
            End Try
        End Sub
    End Class

    Class DataGridView
        Private Delegate Sub ClearSelectionDel(ByVal dgw As System.Windows.Forms.DataGridView)
        Public Shared Sub ClearSelection(ByVal dgw As System.Windows.Forms.DataGridView)
            Try
                If dgw IsNot Nothing Then
                    If dgw.InvokeRequired Then
                        dgw.Invoke(New Safe.DataGridView.ClearSelectionDel(AddressOf Safe.DataGridView.ClearSelection), dgw)
                    Else
                        dgw.ClearSelection()
                    End If
                End If

            Catch ex As Exception
            End Try
        End Sub



        Private Delegate Sub SelectRigaDel(ByVal dgw As System.Windows.Forms.DataGridView, ByVal riga0 As Integer)
        Public Shared Sub SelectRiga(ByVal dgw As System.Windows.Forms.DataGridView, ByVal riga0 As Integer)
            Try
                If dgw IsNot Nothing Then
                    If dgw.InvokeRequired Then
                        dgw.Invoke(New Safe.DataGridView.SelectRigaDel(AddressOf Safe.DataGridView.SelectRiga), dgw, riga0)
                    Else
                        dgw.Rows(riga0).Selected = True
                    End If
                End If

            Catch ex As Exception
            End Try
        End Sub


    End Class

    Class ListBox

        Private Delegate Sub AddListboxsafeDel(ByVal s As String, ByVal l As System.Windows.Forms.ListBox)
        Public Shared Sub AddListboxsafe(ByVal s As String, ByVal l As System.Windows.Forms.ListBox)
            Try
                If l IsNot Nothing Then
                    If l.InvokeRequired Then
                        l.Invoke(New Safe.ListBox.AddListboxsafeDel(AddressOf Safe.ListBox.AddListboxsafe), s, l)
                    Else
                        l.Items.Add(s)
                    End If
                End If

            Catch ex As Exception

            End Try
        End Sub

        Private Delegate Sub AddRangeListboxsafeDel(ByVal stringhe() As String, ByVal ClearBefore As Boolean, ByVal listbox As System.Windows.Forms.ListBox)
        Public Shared Sub AddRangeListboxsafe(ByVal s() As String, ByVal ClearBefore As Boolean, ByVal l As System.Windows.Forms.ListBox)
            Try
                If l.InvokeRequired Then
                    l.Invoke(New Safe.ListBox.AddRangeListboxsafeDel(AddressOf Safe.ListBox.AddRangeListboxsafe), s, ClearBefore, l)
                Else
                    If ClearBefore Then
                        l.Items.Clear()
                    End If
                    l.Items.AddRange(s)
                End If
            Catch ex As Exception

            End Try
        End Sub

        Private Delegate Sub ClearDel(ByVal l As System.Windows.Forms.ListBox)
        Public Shared Sub Clear(ByVal l As System.Windows.Forms.ListBox)
            Try
                If l IsNot Nothing Then
                    If l.InvokeRequired Then
                        l.Invoke(New Safe.ListBox.ClearDel(AddressOf Safe.ListBox.Clear), l)
                    Else
                        l.Items.Clear()
                    End If
                End If
            Catch ex As Exception

            End Try
        End Sub

        Private Delegate Sub ReplaceLastDel(ByVal s As String, ByVal l As System.Windows.Forms.ListBox)
        Public Shared Sub ReplaceLast(ByVal s As String, ByVal l As System.Windows.Forms.ListBox)
            Try
                If l.InvokeRequired Then
                    l.Invoke(New ReplaceLastDel(AddressOf ReplaceLast), s, l)
                Else
                    l.Items(l.Items.Count - 1) = s
                End If
            Catch ex As Exception

            End Try
        End Sub

    End Class



    Class ProgressBar
        Private Delegate Sub ProgBarInitDel(ByVal max As Integer, ByVal c As System.Windows.Forms.ProgressBar)
        Public Shared Sub ProgBarInit(ByVal max As Integer, ByVal c As System.Windows.Forms.ProgressBar)
            Try
                If c IsNot Nothing Then

                    If c.InvokeRequired Then
                        c.Invoke(New ProgBarInitDel(AddressOf ProgBarInit), max, c)
                    Else
                        c.Value = 0
                        c.Maximum = max
                        c.Step = 1
                    End If

                End If
            Catch ex As Exception

            End Try
        End Sub

        Private Delegate Sub ProgBarValDel(ByVal val As Integer, ByVal c As System.Windows.Forms.ProgressBar)
        Public Shared Sub ProgBarVal(ByVal val As Integer, ByVal c As System.Windows.Forms.ProgressBar)
            Try
                If c IsNot Nothing Then

                    If c.InvokeRequired Then
                        c.Invoke(New ProgBarValDel(AddressOf ProgBarVal), val, c)
                    Else
                        c.Value = val

                    End If

                End If
            Catch ex As Exception

            End Try
        End Sub
    End Class

    Class Textbox
        Private Delegate Sub AppendTextDel(ByVal s As String, ByVal txt As System.Windows.Forms.TextBox)
        Public Shared Sub AppendText(ByVal s As String, ByVal txt As System.Windows.Forms.TextBox)
            Try
                If txt IsNot Nothing Then
                    If txt.InvokeRequired Then
                        txt.BeginInvoke(New Safe.Textbox.AppendTextDel(AddressOf Safe.Textbox.AppendText), s, txt)
                    Else
                        txt.AppendText(s)
                    End If
                End If
            Catch ex As Exception

            End Try
        End Sub

    End Class

    Class Control

        Private Delegate Sub ColorDel(ByVal c As System.Windows.Forms.Control, ByVal fore As System.Drawing.Color, ByVal back As System.Drawing.Color)
        Public Shared Sub Color(ByVal c As System.Windows.Forms.Control, ByVal fore As System.Drawing.Color, ByVal back As System.Drawing.Color)
            Try
                If c.InvokeRequired Then
                    c.Invoke(New Safe.Control.ColorDel(AddressOf Safe.Control.Color), c, fore, back)
                Else
                    c.ForeColor = fore
                    c.BackColor = back
                End If
            Catch ex As Exception

            End Try
        End Sub

        Private Delegate Sub FontDel(ByVal c As System.Windows.Forms.Control, ByVal fontName As String, ByVal size As Integer, ByVal style As System.Drawing.FontStyle)
        Public Shared Sub Font(ByVal c As System.Windows.Forms.Control, ByVal fontName As String, ByVal size As Integer, ByVal style As System.Drawing.FontStyle)
            Try
                If c.InvokeRequired Then
                    c.Invoke(New Safe.Control.FontDel(AddressOf Safe.Control.Font), c, fontName, size)
                Else
                    c.Font = New Font(fontName, size, style)

                End If
            Catch ex As Exception

            End Try
        End Sub

        Private Delegate Sub TextDel(ByVal s As String, ByVal c As System.Windows.Forms.Control)
        Public Shared Sub Text(ByVal s As String, ByVal c As System.Windows.Forms.Control)
            Try
                If c IsNot Nothing Then
                    If c.InvokeRequired Then
                        c.Invoke(New Safe.Control.TextDel(AddressOf Safe.Control.Text), s, c)
                    Else
                        c.Text = s
                    End If
                End If
            Catch ex As Exception

            End Try
        End Sub

        Private Delegate Sub BringToFrontDel(ByVal c As System.Windows.Forms.Control)
        Public Shared Sub BringToFront(ByVal c As System.Windows.Forms.Control)
            Try
                If c IsNot Nothing Then
                    If c.InvokeRequired Then
                        c.Invoke(New Safe.Control.BringToFrontDel(AddressOf Safe.Control.BringToFront), c)
                    Else
                        c.BringToFront()
                    End If
                End If
            Catch ex As Exception

            End Try
        End Sub

        Private Delegate Sub VisibleDel(ByVal OnOff As Boolean, ByVal c As System.Windows.Forms.Control)
        Public Shared Sub Visible(ByVal OnOff As Boolean, ByVal c As System.Windows.Forms.Control)
            Try
                If c.InvokeRequired Then
                    c.Invoke(New Safe.Control.VisibleDel(AddressOf Safe.Control.Visible), OnOff, c)
                Else
                    c.Visible = OnOff
                End If
            Catch ex As Exception

            End Try
        End Sub

        Private Delegate Sub EnabledDel(ByVal OnOff As Boolean, ByVal c As System.Windows.Forms.Control)
        Public Shared Sub Enabled(ByVal OnOff As Boolean, ByVal c As System.Windows.Forms.Control)
            Try
                If c IsNot Nothing Then
                    If c.InvokeRequired Then
                        c.Invoke(New Safe.Control.EnabledDel(AddressOf Safe.Control.Enabled), OnOff, c)
                    Else
                        c.Enabled = OnOff
                    End If
                End If
            Catch ex As Exception

            End Try
        End Sub

        Private Delegate Sub RefreshDel(ByVal c As System.Windows.Forms.Control)
        Public Shared Sub Refresh(ByVal c As System.Windows.Forms.Control)
            Try
                If c IsNot Nothing Then
                    If c.InvokeRequired Then
                        c.Invoke(New Safe.Control.RefreshDel(AddressOf Safe.Control.Refresh), c)
                    Else
                        c.Refresh()
                    End If
                End If
            Catch ex As Exception

            End Try
        End Sub

        Private Delegate Sub LocationDel(ByVal loc As System.Drawing.Point, ByVal c As System.Windows.Forms.Control)
        Public Shared Sub Location(ByVal loc As System.Drawing.Point, ByVal c As System.Windows.Forms.Control)
            Try
                If c.InvokeRequired Then
                    c.Invoke(New Safe.Control.LocationDel(AddressOf Safe.Control.Location), loc, c)
                Else
                    c.Location = loc
                End If
            Catch ex As Exception

            End Try
        End Sub


        Private Delegate Sub SizeDel(ByVal Size As System.Drawing.Size, ByVal c As System.Windows.Forms.Control)
        Public Shared Sub Size(ByVal Size As System.Drawing.Size, ByVal c As System.Windows.Forms.Control)
            Try
                If c.InvokeRequired Then
                    c.Invoke(New Safe.Control.SizeDel(AddressOf Safe.Control.Size), Size, c)
                Else
                    c.Size = Size
                End If
            Catch ex As Exception

            End Try
        End Sub

    End Class


    Private Delegate Sub CheckBoxSetDel(ByVal cheked As Boolean, ByVal c As CheckBox)
    Public Shared Sub CheckBoxSet(ByVal cheked As Boolean, ByVal c As CheckBox)
        Try
            If c.InvokeRequired Then
                c.Invoke(New CheckBoxSetDel(AddressOf CheckBoxSet), cheked, c)
            Else
                c.Checked = cheked
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Delegate Function CheckBoxGetDel(ByVal c As CheckBox) As Boolean
    Public Shared Function CheckBoxGet(ByVal c As CheckBox) As Boolean
        Try
            If c.InvokeRequired Then
                c.Invoke(New CheckBoxGetDel(AddressOf CheckBoxGet), c)
            Else
                Return c.Checked
            End If

        Catch ex As Exception

        End Try
    End Function




    Private Delegate Sub ToolStripMenuItemEnableDel(ByVal OnOff As Boolean, ByVal ToolTsm As System.Windows.Forms.ToolStripMenuItem, ByVal f As Form)
    Private Shared Sub ToolStripMenuItemEnableSub(ByVal OnOff As Boolean, ByVal ToolTsm As System.Windows.Forms.ToolStripMenuItem, ByVal f As Form)
        Try
            ToolTsm.Enabled = OnOff
        Catch ex As Exception

        End Try
    End Sub
    Public Shared Sub TooStripEnabled(ByVal OnOff As Boolean, ByVal ToolTsm As System.Windows.Forms.ToolStripMenuItem, ByVal f As System.Windows.Forms.Form)
        Try
            f.Invoke(New ToolStripMenuItemEnableDel(AddressOf ToolStripMenuItemEnableSub), OnOff, ToolTsm, f)
        Catch ex As Exception
        End Try
    End Sub


    Class Form
        Private Delegate Sub SetCursorDel(ByVal c As Cursor, ByVal f As System.Windows.Forms.Form)
        Public Shared Sub SetCursor(ByVal c As Cursor, ByVal f As System.Windows.Forms.Form)
            Try
                If f IsNot Nothing Then
                    If f.InvokeRequired Then
                        f.Invoke(New Safe.Form.SetCursorDel(AddressOf Safe.Form.SetCursor), c, f)
                    Else
                        f.Cursor = c
                    End If
                End If
            Catch ex As Exception

            End Try
        End Sub

        Private Delegate Sub HideDel(ByVal f As System.Windows.Forms.Form)
        Public Shared Sub Hide(ByVal f As System.Windows.Forms.Form)
            Try
                If f.InvokeRequired Then
                    f.Invoke(New Safe.Form.HideDel(AddressOf Safe.Form.Hide), f)
                Else
                    f.Hide()
                End If
            Catch ex As Exception

            End Try
        End Sub

        Private Delegate Sub CloseDel(ByVal f As System.Windows.Forms.Form)
        Public Shared Sub Close(ByVal f As System.Windows.Forms.Form)
            Try
                If f.InvokeRequired Then
                    f.Invoke(New Safe.Form.CloseDel(AddressOf Safe.Form.Close), f)
                Else
                    f.Close()
                End If
            Catch ex As Exception

            End Try
        End Sub
    End Class
End Class
