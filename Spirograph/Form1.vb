Module MyUtil
    Sub MySleep(numSecs As Integer)
        Dim stopTime = DateAdd("s", numSecs, Now())
        Do Until (Now() > stopTime)
        Loop
    End Sub
End Module

Public Class Form1
    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        'Draws a line on picturebox control
        '
        'Put a picturebox on the form named pic
        Dim ring_radius As Int32
        Dim planet_radius As Int32
        Dim planet_arm As Int32
        Dim orbit_radius As Int32
        Dim rotation_count As Double
        Dim step_angle As Double
        Dim rad_angle As Double
        Dim width As Int32
        Dim height As Int32
        Dim origin_offset_x As Int32
        Dim origin_offset_y As Int32

        Dim planet_ring_ratio As Double
        Dim pen_x_start As Int32
        Dim pen_y_start As Int32
        Dim pen_x_end As Int32
        Dim pen_y_end As Int32
        Dim orbit_x_start As Int32
        Dim orbit_x As Int32
        Dim orbit_y_start As Int32
        Dim orbit_y As Int32
        Dim orbit_angle As Double
        Dim planet_arm_angle As Double
        Dim loop_counter As Int32
        Dim total_loops As Int32

        Dim adjusted_x_start As Int32
        Dim adjusted_y_start As Int32
        Dim adjusted_x_end As Int32
        Dim adjusted_y_end As Int32

        'Pass these in from message box
        ring_radius = 300
        planet_radius = 49
        planet_arm = 60
        rotation_count = 100
        'step_angle = 0.25
        step_angle = 2

        'The rest is calculated

        width = PictureBox1.Width
        height = PictureBox1.Height
        origin_offset_x = width / 2
        origin_offset_y = height / 2
        orbit_radius = ring_radius - planet_radius
        planet_ring_ratio = ring_radius / planet_radius
        pen_x_start = planet_arm
        pen_y_start = 0
        orbit_x_start = orbit_radius
        orbit_y_start = 0
        total_loops = rotation_count * (360.0 / step_angle)
        loop_counter = 0

        'Eash Draw initialization
        Dim bitm As Bitmap = New Bitmap(PictureBox1.Width, PictureBox1.Height)
        Dim g As Graphics = Graphics.FromImage(bitm)
        Dim myPen As Pen = New Pen(Color.Blue, 3)
        Dim myPoints As New List(Of Point)

        'Loop
        While loop_counter < total_loops
            For i As Integer = 1 To 4
                'Find Orbit x,y (origin offset not added until drawing)
                orbit_angle = orbit_angle + step_angle
                rad_angle = Math.PI * orbit_angle / 180
                orbit_x = Math.Cos(rad_angle) * orbit_radius
                orbit_y = Math.Sin(rad_angle) * orbit_radius
                'Find Planet Arm / Pen position (origin offset and orbit offset not added until drawing)
                planet_arm_angle = (0 - (rad_angle * planet_ring_ratio)) 'if weird do math in deg not rad
                pen_x_end = Math.Cos(planet_arm_angle) * planet_arm
                pen_y_end = Math.Sin(planet_arm_angle) * planet_arm

                'Drawing Stuff

                adjusted_x_start = origin_offset_x + orbit_x + pen_x_start
                adjusted_x_end = origin_offset_x + orbit_x + pen_x_end
                adjusted_y_start = origin_offset_y + orbit_y + pen_y_start
                adjusted_y_end = origin_offset_y + orbit_y + pen_y_end

                myPoints.Add(New Point(adjusted_x_start, adjusted_y_start))

                pen_x_start = pen_x_end
                pen_y_start = pen_y_end

                myPen.Color = Color.Blue
                myPen.Width = 3
                'g.DrawLine(myPen, adjusted_x_start, adjusted_y_start, adjusted_x_start + 1, adjusted_y_start + 1)
                g.DrawEllipse(myPen, adjusted_x_start, adjusted_y_start, 3, 3)
            Next

            myPen.Color = Color.Pink
            myPen.Width = 2
            g.DrawBezier(myPen, myPoints(0), myPoints(1), myPoints(2), myPoints(3))

            PictureBox1.Image = bitm
            PictureBox1.Refresh()

            myPoints.RemoveRange(0, 3)
        End While

    End Sub


End Class
