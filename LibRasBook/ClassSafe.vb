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

        Private Delegate Sub TreeAddImgDel(ByVal text As String, ByVal treeW As System.Windows.Forms.TreeView, ByVal ImgIndex As Integer, ByRef treeData As ClassTreeData)
        Public Shared Sub TreeAddImg(ByVal text As String, ByVal treeW As System.Windows.Forms.TreeView, ByVal ImgIndex As Integer, ByRef treeData As ClassTreeData)
            Try
                If treeW IsNot Nothing Then
                    If treeW.InvokeRequired Then
                        treeW.Invoke(New Safe.TreeView.TreeAddImgDel(AddressOf Safe.TreeView.TreeAddImg), text, treeW, ImgIndex, treeData)
                    Else
                        treeData.Nodo = treeW.Nodes.Add("", text, ImgIndex, ImgIndex)
                        'treeData.Nodo.NodeFont = New Font(treeW.Font, FontStyle.Bold)
                    End If
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message + vbCrLf + ex.StackTrace, "Safe.TreeView.TreeAddImg", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        Private Delegate Sub NodesAddImgDel(ByVal text As String, ByVal nodoPadre As System.Windows.Forms.TreeNode, ByVal ImgIndex As Integer, ByRef treeData As ClassTreeData)
        Public Shared Sub NodesAddImg(ByVal text As String, ByVal nodoPadre As System.Windows.Forms.TreeNode, ByVal ImgIndex As Integer, ByRef treeData As ClassTreeData)
            Try
                If nodoPadre IsNot Nothing Then
                    If nodoPadre.TreeView.InvokeRequired Then
                        nodoPadre.TreeView.Invoke(New Safe.TreeView.NodesAddImgDel(AddressOf Safe.TreeView.NodesAddImg), text, nodoPadre, ImgIndex, treeData)
                    Else
                        treeData.Nodo = nodoPadre.Nodes.Add("", text, ImgIndex, ImgIndex)
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

End Class
