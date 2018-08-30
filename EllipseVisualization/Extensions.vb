Imports System.Runtime.CompilerServices

Module Extensions
    <Extension()>
    Public Function Plus(p1 As PointF, p2 As PointF) As PointF
        Return New PointF(p1.X + p2.X, p1.Y + p2.Y)
    End Function

    <Extension()>
    Public Function Minus(p1 As Point, p2 As PointF) As PointF
        Return New Point(p1.X - p2.X, p1.Y - p2.Y)
    End Function

    <Extension()>
    Public Function Minus(p1 As PointF, p2 As PointF) As PointF
        Return New PointF(p1.X - p2.X, p1.Y - p2.Y)
    End Function

    <Extension()>
    Public Function Mult(p1 As PointF, v As Single) As PointF
        Return New PointF(p1.X * v, p1.Y * v)
    End Function

    <Extension()>
    Public Function Distance(p1 As PointF, p2 As PointF) As Single
        Return Math.Sqrt((p2.X - p1.X) ^ 2 + (p2.Y - p1.Y) ^ 2)
    End Function

    <Extension()>
    Public Function Slope(p1 As PointF, p2 As PointF) As Single
        Return (p2.Y - p1.Y) / (p2.X - p1.X)
    End Function

    <Extension()>
    Public Function IsEqual(p1 As PointF, p2 As PointF) As Boolean
        Return p1.X = p2.X AndAlso p1.Y = p2.Y
    End Function

    <Extension()>
    Public Function InLine(p As PointF, p1 As PointF, p2 As PointF) As Boolean
        If p1.X <> p2.X Then
            If p1.X <= p.X AndAlso p.X <= p2.X Then Return True
            If p1.X >= p.X AndAlso p.X >= p2.X Then Return True
        Else
            If p1.Y <= p.Y AndAlso p.Y < p2.Y Then Return True
            If p1.Y >= p.Y AndAlso p.Y >= p2.Y Then Return True
        End If

        Return False
    End Function

    <Extension()>
    Public Function Dot(p1 As PointF, p2 As PointF) As Single
        Return p1.X * p2.X + p1.Y * p2.Y
    End Function

    <Extension()>
    Public Function Perp(p1 As PointF, p2 As PointF) As Single
        Return p1.X * p2.Y + p1.Y * p2.X
    End Function
End Module
