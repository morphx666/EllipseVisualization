Public Class FormMain
    Public Enum OverSections
        None
        EllipseTop
        EllipseLeft
        EllipseRight
        EllipseBottom
        EccentricPoint
    End Enum
    Private overSection As OverSections = OverSections.None

    Private w As Single ' Ellipse Width
    Private h As Single ' Ellipse Height
    Private ecp As PointF ' Ellipse Center Point
    Private ep As PointF ' Eccentric Point

    Private mouseOrigin As PointF
    Private mouseIsDown As Boolean

    Private Const finalAngle As Double = Math.PI / 2
    Private angle As Double

    Private Sub FormMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.SetStyle(ControlStyles.AllPaintingInWmPaint, True)
        Me.SetStyle(ControlStyles.OptimizedDoubleBuffer, True)
        Me.SetStyle(ControlStyles.ResizeRedraw, True)
        Me.SetStyle(ControlStyles.UserPaint, True)

        w = Math.Min(Me.DisplayRectangle.Width, Me.DisplayRectangle.Height) - 150
        h = w / 1

        ep = New PointF(w / 2 - 80, 0)
        SetEllipseCenterPoint()

        AddHandler Me.MouseMove, Sub(s1 As Object, e1 As MouseEventArgs) HandleMouseMove(e1)

        AddHandler Me.MouseDown, Sub(s1 As Object, e1 As MouseEventArgs)
                                     If overSection <> OverSections.None AndAlso e1.Button = MouseButtons.Left Then
                                         mouseOrigin = e1.Location
                                         mouseIsDown = True
                                     End If
                                 End Sub

        AddHandler Me.MouseUp, Sub() mouseIsDown = False

        Task.Run(Sub()
                     Threading.Thread.Sleep(400)
                     Do
                         angle += 0.01
                         Me.Invalidate()
                         Threading.Thread.Sleep(10)
                     Loop Until angle >= finalAngle
                     angle = finalAngle
                 End Sub)
    End Sub

    Private Sub SetEllipseCenterPoint()
        ecp = New PointF(Me.DisplayRectangle.Width / 2, Me.DisplayRectangle.Height / 2)
    End Sub

    Private Sub HandleMouseMove(e As MouseEventArgs)
        If mouseIsDown Then
            Select Case overSection
                Case OverSections.EccentricPoint
                    ep.X += e.X - mouseOrigin.X
                    ep.Y += e.Y - mouseOrigin.Y
                Case OverSections.EllipseLeft : w -= e.X - mouseOrigin.X
                Case OverSections.EllipseRight : w += e.X - mouseOrigin.X
                Case OverSections.EllipseTop : h -= e.Y - mouseOrigin.Y
                Case OverSections.EllipseBottom : h += e.Y - mouseOrigin.Y
            End Select
            mouseOrigin = e.Location
            Me.Invalidate()
        Else
            If IsInsideElipse(e.Location, New PointF(ep.X + Me.DisplayRectangle.Width / 2, ep.Y + Me.DisplayRectangle.Height / 2), 10, 10) Then
                overSection = OverSections.EccentricPoint
                Me.Cursor = Cursors.Hand
            ElseIf IsOverElipse(e.Location, ecp, w, h, 0.1) Then
                overSection = OverSections.None
                If Math.Abs(ecp.X - e.X) <= 20 Then
                    If Math.Abs(ecp.Y - h / 2 - e.Y) <= 20 Then
                        overSection = OverSections.EllipseTop
                        Me.Cursor = Cursors.SizeNS
                    ElseIf Math.Abs(ecp.Y + h / 2 - e.Y) <= 20 Then
                        overSection = OverSections.EllipseBottom
                        Me.Cursor = Cursors.SizeNS
                    End If
                ElseIf Math.Abs(ecp.Y - e.Y) <= 20 Then
                    If Math.Abs(ecp.X - w / 2 - e.X) <= 20 Then
                        overSection = OverSections.EllipseLeft
                        Me.Cursor = Cursors.SizeWE
                    ElseIf Math.Abs(ecp.X + w / 2 - e.X) <= 20 Then
                        overSection = OverSections.EllipseRight
                        Me.Cursor = Cursors.SizeWE
                    End If
                End If
            Else
                overSection = OverSections.None
            End If
            If overSection = OverSections.None Then Me.Cursor = Cursors.Default
            'Debug.WriteLine(overSection)
        End If
    End Sub

    Private Sub FormMain_Paint(sender As Object, e As PaintEventArgs) Handles Me.Paint
        Dim g As Graphics = e.Graphics
        g.Clear(Color.Black)

        SetEllipseCenterPoint()
        g.TranslateTransform(ecp.X, ecp.Y)

        Dim p1 As PointF
        Dim p2 As PointF
        Dim cp As PointF
        Using p As New Pen(Color.FromArgb(128, Color.Gray))
            For a As Single = 0 To Math.PI * 2 Step 0.05
                p2.X = w / 2 * Math.Cos(a)
                p2.Y = h / 2 * Math.Sin(a)
                g.DrawLine(p, ep, p2)

                cp = GetCenter(ep, p2)
                p1 = RotateLineAtPoint(ep, cp, angle)
                p2 = RotateLineAtPoint(p2, cp, angle)
                g.DrawLine(Pens.White, p1, p2)
            Next
        End Using

        Using p As New Pen(Color.Cyan, 2)
            DrawEllipseFromCenter(g, p, 0, 0, w, h)
        End Using
        DrawEllipseFromCenter(g, Pens.Yellow, ep.X, ep.Y, 10, 10, True)

        g.ResetTransform()
        g.DrawString($"Use mouse to resize Ellipse{Environment.NewLine}and move point", Me.Font, Brushes.Gray, 10, 10)
    End Sub

    Private Function GetCenter(p1 As PointF, p2 As PointF) As PointF
        Return New PointF(p1.X + (p2.X - p1.X) / 2, p1.Y + (p2.Y - p1.Y) / 2)
    End Function

    Private Function RotateLineAtPoint(p1 As PointF, p2 As PointF, angle As Double) As PointF
        Dim p3 As PointF = New Point(p2.X - p1.X, p2.Y - p1.Y)
        p3 = RotateLineAtOrigin(p3, angle)
        Return New PointF(p2.X + p3.X, p2.Y + p3.Y)
    End Function

    Private Function RotateLineAtOrigin(p As PointF, angle As Double) As PointF
        Return New Point(
            p.X * Math.Cos(angle) - p.Y * Math.Sin(angle),
            p.X * Math.Sin(angle) + p.Y * Math.Cos(angle)
        )
    End Function

    Private Sub DrawEllipseFromCenter(g As Graphics, p As Pen, x As Single, y As Single, w As Single, h As Single, Optional fill As Boolean = False)
        If fill Then
            Using b As New SolidBrush(p.Color)
                g.FillEllipse(b, x - w \ 2, y - h \ 2, w, h)
            End Using
        Else
            g.DrawEllipse(p, x - w \ 2, y - h \ 2, w, h)
        End If
    End Sub

    Private Function IsInsideElipse(p As PointF, cp As PointF, w As Single, h As Single) As Boolean
        Return (p.X - cp.X) ^ 2 / (w / 2) ^ 2 + (p.Y - cp.Y) ^ 2 / (h / 2) ^ 2 < 1.0
    End Function

    Private Function IsOverElipse(p As PointF, cp As PointF, w As Single, h As Single, precision As Double) As Boolean
        Dim v As Double = (p.X - cp.X) ^ 2 / (w / 2) ^ 2 + (p.Y - cp.Y) ^ 2 / (h / 2) ^ 2
        Return v >= 1.0 - precision AndAlso v <= 1.0 + precision
    End Function
End Class
