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
                Case OverSections.EccentricPoint : ep = ep.Plus(e.Location.Minus(mouseOrigin))
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
        Dim a As Single

        Using p As New Pen(Color.FromArgb(128, Color.Gray))
            For a = 0 To Math.PI * 2 Step Math.PI / 60
                p2.X = w / 2 * Math.Cos(a)
                p2.Y = -h / 2 * Math.Sin(a)
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

        ' TODO: generate the inner ellipse's formula and draw it.
        '           I guess I'll have to re-watch the video...
        'p1 = New PointF(0, 0)
        'g.DrawLine(Pens.Red, p1, ep)
        'a = Math.Atan(p1.Slope(ep))
        'cp = GetCenter(p1, ep)

        'DrawEllipseFromCenter(g, Pens.Red, 0, 0, 10, 10, True)

        'g.ResetTransform()
        'g.TranslateTransform(ecp.X + cp.X, ecp.Y + cp.Y)
        'g.RotateTransform(a * 180 / Math.PI)
        'Using p As New Pen(Color.Magenta, 2)
        '    DrawEllipseFromCenter(g, p, 0, 0, ecp.X / 2, ecp.Y / 2)
        'End Using

        g.ResetTransform()
        g.DrawString($"Use mouse to resize Ellipse{Environment.NewLine}and move point", Me.Font, Brushes.Gray, 10, 10)
    End Sub

    ' http://geomalgorithms.com/a05-_intersect-1.html#intersect2D_2Segments()
    Private Function GetIntersection(s10 As PointF, s11 As PointF, s20 As PointF, s21 As PointF) As PointF
        Dim u As PointF = s11.Minus(s10)
        Dim v As PointF = s21.Minus(s20)
        Dim w As PointF = s10.Minus(s20)
        Dim d As Single = u.Perp(v)

        If Math.Abs(d) < 0.001 Then ' The lines are parallel
            If u.Perp(w) <> 0 OrElse v.Perp(w) <> 0 Then Return Point.Empty

            Dim du As Single = u.Dot(u)
            Dim dv As Single = v.Dot(v)

            If du = 0 OrElse dv = 0 Then
                If Not s10.Equals(s20) Then Return PointF.Empty
                Return s10
            End If

            If du = 0 Then
                If Not s10.InLine(s20, s21) Then Return Point.Empty
                Return s10
            End If
            If dv = 0 Then
                If Not s20.InLine(s10, s11) Then Return Point.Empty
                Return s20
            End If

            Dim t0 As Single
            Dim t1 As Single
            Dim w2 As PointF = s11.Minus(s20)
            If v.X <> 0 Then
                t0 = w.X / v.X
                t1 = w2.X / v.X
            Else
                t0 = w.Y / v.Y
                t1 = w2.Y / v.Y
            End If
            If t0 > t1 Then
                Dim t As Single = t0
                t0 = t1
                t1 = t0
            End If
            If t0 > 1 OrElse t1 < 0 Then Return Point.Empty
            t0 = If(t0 < 0, 0, t0)
            t1 = If(t1 > 1, 1, t1)
            If t0 = t1 Then Return s20.Plus(v.Mult(t0))

            s20.Plus(v.Mult(t0))
        End If

        Dim si As Single = v.Perp(w) / d
        If si < 0 OrElse si > 1 Then Return Point.Empty

        Dim ti As Single = u.Perp(w) / d
        If ti < 0 OrElse ti > 1 Then Return Point.Empty

        Return s10.Plus(u.Mult(si))
    End Function

    Private Function GetCenter(p1 As PointF, p2 As PointF) As PointF
        Return New PointF(p1.X + (p2.X - p1.X) / 2, p1.Y + (p2.Y - p1.Y) / 2)
    End Function

    Private Function RotateLineAtPoint(p1 As PointF, p2 As PointF, angle As Double) As PointF
        Dim p3 As PointF = p1.Minus(p2)
        p3 = RotateLineAtOrigin(p3, angle)
        Return p2.Plus(p3)
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
