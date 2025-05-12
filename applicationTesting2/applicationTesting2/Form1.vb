Imports System.ComponentModel.DataAnnotations
Imports System.Drawing.Text
Imports System.IO
Imports System.IO.Directory
Imports System.Reflection.Metadata
Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.JavaScript.JSType
Imports System.Security.Cryptography.Xml
Imports System.Windows.Forms.Design
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.TaskbarClock
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.ToolTip

Imports System.Windows.Forms.DataVisualization.Charting
Imports System.Threading.Channels

Public Class Form1
    Private DataJSON = New Data
    Private data = New Dictionary(Of Object, Object)
    Private id As String
    Private newWindow = False

    Private selection As TextBox
    Private units As ListBox
    Private selection1 As TextBox
    Private units1 As ListBox
    Private selection2 As TextBox
    Private units2 As ListBox
    Private estimated_diameter As Label

    Private orbital_data As Panel
    Private morestats As Button
    Private close_approach_data As Panel
    Private footer As Panel
    Private series As Series
    Private chart As New Chart()
    Private series1 As Series
    Private chart1 As New Chart()


    Private moreStatsExpand As Boolean = False

    Private pfc As New PrivateFontCollection()
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.AutoScroll = True
        Me.Text = "NASA API Viewer"
        Me.BackColor = Color.Gray
        TextBox1.Location = New Point((Me.Width / 2) - (TextBox1.Width / 2), (Me.Height / 2) - 52 - (TextBox1.Height / 2))   'offset of 47 px for height
        Button1.Location = New Point((Me.Width / 2) - (Button1.Width / 2), (Me.Height / 2) - 42 + (Button1.Height / 2))
        'MsgBox("height = " & Me.Height & vbCrLf & vbCrLf & "desktopbound height = " & Me.DisplayRectangle.Height)    'actual height of application
        'MsgBox("width = " & Me.Width & vbCrLf & vbCrLf & "desktopbound width = " & Me.DisplayRectangle.Width)        'actual width of applcation for debugging
        pfc.AddFontFile(IO.Path.Combine(Directory.GetParent(Application.StartupPath).Parent.Parent.Parent.FullName, "BlackOpsOne-Regular.ttf"))     'import font
    End Sub

    Private Sub Form1_Resize(sender As Object, e As EventArgs) Handles MyBase.Resize                'forces the form to the center of the screen on form resize (depriciated probably)
        Dim screenCenter As Point = Screen.PrimaryScreen.WorkingArea.Location + Screen.PrimaryScreen.WorkingArea.Size / 2
        Dim formCenter As Point = Location + Size / 2
        Location = New Point(screenCenter.X - formCenter.X + (Size.Width \ 2), screenCenter.Y - formCenter.Y + (Size.Height \ 2))
    End Sub

    Private Sub button1_click(sender As Object, e As EventArgs) Handles Button1.Click   'on click submit the inputted ID with error handling for blank input
        If TextBox1.Text = "" Then
            TextBox1.Text = 0
        End If
        id = TextBox1.Text

        createDataWindow()
    End Sub

    Private Sub createDataWindow()
        data = DataJSON.parseData(id)       'runs the parser and assigns the value

        Try
            data.item("links")      'checks if the form generated properly

            Me.Controls.Clear()     'clears previous frame, resizes from and relocated frame
            Me.Height = 900
            Me.Width = 1200

            Location = New Point(Location.X - 300, Location.Y - 200)

            initializeNewWindow()   'initializes elements of new window
        Catch ex As Exception
            MsgBox("Invalid Id.")   'error handling
            TextBox1.Text = ""
        End Try
    End Sub

    Private Sub initializeNewWindow()       'initializes all elements present in new window and sets a variable to true for other code
        initializeHelp()
        initializeName()
        initializeId()
        initializeJPLURL()
        initializeAbsMagnitude()
        initializeIsHazardous()
        initializeSentryObj()

        initializeEstDiameter()
        initializeOrbitalData()
        initializeCloseApproachData()

        initializeFooter()
        newWindow = True
    End Sub

    Private Sub initializeHelp()
        Dim help As New Label
        With help
            .Size = New Point(600, 40)
            .Location = New Point(Me.Width / 2 - .Width / 2, 107)
            .Font = New Font(pfc.Families(0), 16, FontStyle.Regular)
            .Text = "Hover over data for more details."
            .TextAlign = ContentAlignment.MiddleCenter

        End With
        Me.Controls.Add(help)
    End Sub

    Private Sub initializeName()
        Dim name As New Label
        With name
            .Size = New Point(600, 40)
            .Location = New Point(Me.Width / 2 - .Width / 2, 147)
            .Font = New Font(pfc.Families(0), 16, FontStyle.Regular)
            .Text = "Name: " & data.item("name")
        End With
        Me.Controls.Add(name)

        Dim ToolTip As New ToolTip
        ToolTip.ShowAlways = True
        ToolTip.SetToolTip(name, "The official name of the asteroid.")
    End Sub

    Private Sub initializeId()
        Dim id As New Label
        With id
            .Size = New Point(600, 40)
            .Location = New Point(Me.Width / 2 - .Width / 2, 187)
            .Font = New Font(pfc.Families(0), 16, FontStyle.Regular)
            .Text = "Id: " & data.item("id")
        End With
        Me.Controls.Add(id)

        Dim ToolTip As New ToolTip
        ToolTip.ShowAlways = True
        ToolTip.SetToolTip(id, "The unique identifier for the asteroid.")
    End Sub

    Private Sub initializeAbsMagnitude()
        Dim absolute_magnitude_h As New Label
        With absolute_magnitude_h
            .Size = New Point(600, 40)
            .Location = New Point(Me.Width / 2 - .Width / 2, 227)
            .Font = New Font(pfc.Families(0), 16, FontStyle.Regular)
            .Text = "Absolute Magnitude H: " & data.item("absolute_magnitude_h")
        End With
        Me.Controls.Add(absolute_magnitude_h)

        Dim ToolTip As New ToolTip
        ToolTip.ShowAlways = True
        ToolTip.SetToolTip(absolute_magnitude_h, "A measure of the asteroid's brightness from a standard distance.")
    End Sub

    Private Sub initializeIsHazardous()
        Dim is_potentially_hazardous_asteroid As New Label
        With is_potentially_hazardous_asteroid
            .Size = New Point(600, 40)
            .Location = New Point(Me.Width / 2 - .Width / 2, 267)
            .Font = New Font(pfc.Families(0), 16, FontStyle.Regular)
            If data.item("is_potentially_hazardous_asteroid") = "False" Then    'boolean handling
                .Text = "Potentially Hazardous: No"

            ElseIf data.item("is_potentially_hazardous_asteroid") = "True" Then
                .Text = "Potentially Hazardous: Yes"
            End If
        End With
        Me.Controls.Add(is_potentially_hazardous_asteroid)

        Dim ToolTip As New ToolTip
        ToolTip.ShowAlways = True
        ToolTip.SetToolTip(is_potentially_hazardous_asteroid, "Indicates whether the asteroid is potentially hazardous to Earth.")
    End Sub

    Private Sub initializeSentryObj()
        Dim is_sentry_object As New Label
        With is_sentry_object
            .Size = New Point(600, 40)
            .Location = New Point(Me.Width / 2 - .Width / 2, 307)
            .Font = New Font(pfc.Families(0), 16, FontStyle.Regular)
            If data.item("is_sentry_object") = "False" Then
                .Text = "Sentry Object: No"

            ElseIf data.item("is_sentry_object") = "True" Then
                .Text = "Sentry Object: Yes"
            End If
        End With
        Me.Controls.Add(is_sentry_object)

        Dim ToolTip As New ToolTip
        ToolTip.ShowAlways = True
        ToolTip.SetToolTip(is_sentry_object, "Indicates whether the asteroid is a Sentry object (an object monitored for potential future impact with Earth).")
    End Sub

    Private Sub initializeJPLURL()
        Dim nasa_jpl_url As New LinkLabel
        AddHandler nasa_jpl_url.LinkClicked, AddressOf nasa_jpl_url_LinkClicked
        With nasa_jpl_url
            .Size = New Point(600, 40)
            .Location = New Point(Me.Width / 2 - .Width / 2, 347)
            .Font = New Font(pfc.Families(0), 16, FontStyle.Regular)
            .Text = "NASA JPL URL: visit"
            .Links.Add(14, 5, data.item("nasa_jpl_url"))
        End With
        Me.Controls.Add(nasa_jpl_url)

        Dim ToolTip As New ToolTip
        ToolTip.ShowAlways = True
        ToolTip.SetToolTip(nasa_jpl_url, "A link to the JPL (Jet Propulsion Laboratory) tool for further information on the asteroid.")
    End Sub

    Private Sub initializeEstDiameter()
        Dim est_diameter_panel = New Panel
        With est_diameter_panel
            .Size = New Point(500, 200)
            .Location = New Point(50, 507)
            .Font = New Font(pfc.Families(0), 16, FontStyle.Regular)
        End With

        initializeDropDown(New List(Of String)({"kilometers", "meters", "miles", "feet"}), est_diameter_panel, 320, 0)


        estimated_diameter = New Label
        With estimated_diameter

            .Size = New Point(500, 120)
            .Font = New Font(pfc.Families(0), 16, FontStyle.Regular)
            .Text = "Estimated Diameter: " &
            vbCrLf & "Est. Min: " & data.item("estimated_diameter").item(selection.Text.ToLower).item("estimated_diameter_min") &
            vbCrLf & "Est. Max: " & data.item("estimated_diameter").item(selection.Text.ToLower).item("estimated_diameter_max")
        End With

        Me.Controls.Add(est_diameter_panel)
        est_diameter_panel.Controls.Add(estimated_diameter)

        Dim ToolTip As New ToolTip
        ToolTip.ShowAlways = True
        ToolTip.SetToolTip(estimated_diameter, "The asteroid's estimated diameter, given as a minimum and a maximum.")
    End Sub

    Private Sub initializeDropDown(ByVal liist As List(Of String), ByRef panel As Panel, x As Integer, y As Integer)
        selection = New TextBox
        With selection
            .Location = New Point(x, y)
            .Size = New Point(140, 40)
            .Font = New Font(pfc.Families(0), 15.5, FontStyle.Regular)
            .ReadOnly = True
            .Text = liist(0)
        End With

        Dim dropdownbutton As New Button
        With dropdownbutton
            .Location = New Point(x + 140, y)
            .Size = New Point(40, 40)
            .Font = New Font(pfc.Families(0), 14, FontStyle.Regular)
            .Image = Image.FromFile(IO.Path.Combine(Directory.GetParent(Application.StartupPath).Parent.Parent.Parent.FullName, "dropdownarrow.png"))
            .FlatStyle = FlatStyle.Flat
        End With

        units = New ListBox
        With units
            .Location = New Point(x - 1, y + 47)
            .Size = New Size(180, 200)
            For Each l In liist             'units into listbox
                .Items.Add(l)
            Next
            .Font = New Font(pfc.Families(0), 16, FontStyle.Regular)
            .Show()
            .Hide()
        End With
        panel.Controls.Add(selection)
        panel.Controls.Add(dropdownbutton)
        panel.Controls.Add(units)

        AddHandler dropdownbutton.Click, AddressOf dropdownbutton_click
        AddHandler units.SelectedIndexChanged, AddressOf units_selectedIndexChanged
    End Sub

    Private Sub initializeOrbitalData()
        orbital_data = New Panel
        With orbital_data

            .Size = New Point(500, 440)
            .Location = New Point(600, 507)
            .Font = New Font(pfc.Families(0), 16, FontStyle.Regular)
        End With

        Dim orbital_data_header As New Label
        With orbital_data_header
            .Size = New Point(500, 40)
            .Location = New Point(0, 0)
            .Font = New Font(pfc.Families(0), 16, FontStyle.Regular)
            .Text = "Orbital Data:"
            .TextAlign = ContentAlignment.MiddleCenter
        End With

        Dim orbit_id As New Label
        With orbit_id
            .Size = New Point(500, 40)
            .Location = New Point(0, 40)
            .Font = New Font(pfc.Families(0), 16, FontStyle.Regular)
            .Text = "Orbit Id: " & data.item("orbital_data").item("orbit_id")
        End With

        Dim orbit_determination_date As New Label
        With orbit_determination_date
            .Size = New Point(500, 80)
            .Location = New Point(0, 80)
            .Font = New Font(pfc.Families(0), 16, FontStyle.Regular)
            .Text = "Orbit Determination Date: " & vbCrLf & data.item("orbital_data").item("orbit_determination_date")
        End With

        Dim first_observation_date As New Label
        With first_observation_date
            .Size = New Point(500, 40)
            .Location = New Point(0, 160)
            .Font = New Font(pfc.Families(0), 16, FontStyle.Regular)
            .Text = "First Observed: " & data.item("orbital_data").item("first_observation_date")
        End With

        Dim last_observation_date As New Label
        With last_observation_date
            .Size = New Point(500, 40)
            .Location = New Point(0, 200)
            .Font = New Font(pfc.Families(0), 16, FontStyle.Regular)
            .Text = "Last Obeserved: " & data.item("orbital_data").item("last_observation_date")
        End With

        Dim data_arc_in_days As New Label
        With data_arc_in_days
            .Size = New Point(500, 40)
            .Location = New Point(0, 240)
            .Font = New Font(pfc.Families(0), 16, FontStyle.Regular)
            .Text = "Days Observed: " & data.item("orbital_data").item("data_arc_in_days")
        End With

        Dim observations_used As New Label
        With observations_used
            .Size = New Point(500, 40)
            .Location = New Point(0, 280)
            .Font = New Font(pfc.Families(0), 16, FontStyle.Regular)
            .Text = "Observations Used: " & data.item("orbital_data").item("observations_used")
        End With

        Dim orbital_period As New Label
        With orbital_period
            .Size = New Point(500, 80)
            .Location = New Point(0, 320)
            .Font = New Font(pfc.Families(0), 16, FontStyle.Regular)
            .Text = "Orbital Period (Days): " & vbCrLf & data.item("orbital_data").item("orbital_period")
        End With

        morestats = New Button
        With morestats
            .Size = New Size(280, 40)
            .Location = New Point(110, 400)
            .Font = New Font(pfc.Families(0), 15, FontStyle.Regular)
            .Text = "Stats For Nerds"
            .Image = Image.FromFile(IO.Path.Combine(Directory.GetParent(Application.StartupPath).Parent.Parent.Parent.FullName, "dropdownarrow.png"))
            .TextAlign = ContentAlignment.MiddleLeft
            .ImageAlign = ContentAlignment.MiddleRight
            .FlatStyle = FlatStyle.Flat
            .FlatAppearance.BorderSize = 0
        End With
        AddHandler morestats.Click, AddressOf morestats_click

        Dim orbit_uncertainty = New Label
        With orbit_uncertainty
            .Size = New Point(500, 40)
            .Location = New Point(0, 440)
            .Font = New Font(pfc.Families(0), 16, FontStyle.Regular)
            .Text = "Orbit Uncertainty: " & data.item("orbital_data").item("orbit_uncertainty")
        End With

        Dim minimum_orbit_intersection = New Label
        With minimum_orbit_intersection
            .Size = New Point(500, 80)
            .Location = New Point(0, 480)
            .Font = New Font(pfc.Families(0), 16, FontStyle.Regular)
            .Text = "Minimuim Orbit Intersection: " & vbCrLf & data.item("orbital_data").item("minimum_orbit_intersection")
        End With

        Dim jupiter_tisserand_invariant = New Label
        With jupiter_tisserand_invariant
            .Size = New Point(500, 80)
            .Location = New Point(0, 560)
            .Font = New Font(pfc.Families(0), 16, FontStyle.Regular)
            .Text = "Jupiter Tisserand Invariant: " & vbCrLf & data.item("orbital_data").item("jupiter_tisserand_invariant")
        End With

        Dim epoch_osculation = New Label
        With epoch_osculation
            .Size = New Point(500, 80)
            .Location = New Point(0, 640)
            .Font = New Font(pfc.Families(0), 16, FontStyle.Regular)
            .Text = "Epoch Osculation: " & vbCrLf & data.item("orbital_data").item("epoch_osculation")
        End With

        Dim eccentricity = New Label
        With eccentricity
            .Size = New Point(500, 80)
            .Location = New Point(0, 720)
            .Font = New Font(pfc.Families(0), 16, FontStyle.Regular)
            .Text = "Eccentricity: " & data.item("orbital_data").item("eccentricity")
        End With

        Dim semi_major_axis = New Label
        With semi_major_axis
            .Size = New Point(500, 80)
            .Location = New Point(0, 800)
            .Font = New Font(pfc.Families(0), 16, FontStyle.Regular)
            .Text = "Semi Major Axis: " & vbCrLf & data.item("orbital_data").item("semi_major_axis")
        End With

        Dim inclination = New Label
        With inclination
            .Size = New Point(500, 40)
            .Location = New Point(0, 880)
            .Font = New Font(pfc.Families(0), 16, FontStyle.Regular)
            .Text = "Inclination: " & data.item("orbital_data").item("inclination")
        End With

        Dim ascending_node_longitude = New Label
        With ascending_node_longitude
            .Size = New Point(500, 80)
            .Location = New Point(0, 920)
            .Font = New Font(pfc.Families(0), 16, FontStyle.Regular)
            .Text = "Ascending Node Longitude: " & vbCrLf & data.item("orbital_data").item("ascending_node_longitude")
        End With

        Dim perihelion_distance = New Label
        With perihelion_distance
            .Size = New Point(500, 80)
            .Location = New Point(0, 1000)
            .Font = New Font(pfc.Families(0), 16, FontStyle.Regular)
            .Text = "Perihelion Distance: " & vbCrLf & data.item("orbital_data").item("perihelion_distance")
        End With

        Dim perihelion_argument = New Label
        With perihelion_argument
            .Size = New Point(500, 80)
            .Location = New Point(0, 1080)
            .Font = New Font(pfc.Families(0), 16, FontStyle.Regular)
            .Text = "Perihelion Argument: " & vbCrLf & data.item("orbital_data").item("perihelion_argument")
        End With

        Dim aphelion_distance = New Label
        With aphelion_distance
            .Size = New Point(500, 80)
            .Location = New Point(0, 1160)
            .Font = New Font(pfc.Families(0), 16, FontStyle.Regular)
            .Text = "Aphelion Distance: " & vbCrLf & data.item("orbital_data").item("aphelion_distance")
        End With

        Dim perihelion_time = New Label
        With perihelion_time
            .Size = New Point(500, 80)
            .Location = New Point(0, 1240)
            .Font = New Font(pfc.Families(0), 16, FontStyle.Regular)
            .Text = "Perihelion Time: " & vbCrLf & data.item("orbital_data").item("perihelion_time")
        End With

        Dim mean_anomaly = New Label
        With mean_anomaly
            .Size = New Point(500, 80)
            .Location = New Point(0, 1320)
            .Font = New Font(pfc.Families(0), 16, FontStyle.Regular)
            .Text = "Mean Anomaly: " & vbCrLf & data.item("orbital_data").item("mean_anomaly")
        End With

        Dim mean_motion = New Label
        With mean_motion
            .Size = New Point(500, 80)
            .Location = New Point(0, 1400)
            .Font = New Font(pfc.Families(0), 16, FontStyle.Regular)
            .Text = "Mean Motion: " & vbCrLf & data.item("orbital_data").item("mean_motion")
        End With

        Dim equinox = New Label
        With equinox
            .Size = New Point(500, 40)
            .Location = New Point(0, 1480)
            .Font = New Font(pfc.Families(0), 16, FontStyle.Regular)
            .Text = "Equinox: " & data.item("orbital_data").item("equinox")
        End With

        Dim orbit_class_type = New Label
        With orbit_class_type
            .Size = New Point(500, 40)
            .Location = New Point(0, 1520)
            .Font = New Font(pfc.Families(0), 16, FontStyle.Regular)
            .Text = "Orbit Class Type: " & data.item("orbital_data").item("orbit_class").item("orbit_class_type")
        End With

        Dim orbit_class_description = New Label
        With orbit_class_description
            .Size = New Point(500, 80)
            .Location = New Point(0, 1560)
            .Font = New Font(pfc.Families(0), 16, FontStyle.Regular)
            .Text = "Orbit Class Description: " & vbCrLf & data.item("orbital_data").item("orbit_class").item("orbit_class_description")
        End With

        Dim orbit_class_range = New Label
        With orbit_class_range
            .Size = New Point(500, 80)
            .Location = New Point(0, 1640)
            .Font = New Font(pfc.Families(0), 16, FontStyle.Regular)
            .Text = "Orbit Class Range: " & vbCrLf & data.item("orbital_data").item("orbit_class").item("orbit_class_range")
        End With



        orbital_data.Controls.Add(orbital_data_header)
        orbital_data.Controls.Add(orbit_id)
        orbital_data.Controls.Add(orbit_determination_date)
        orbital_data.Controls.Add(first_observation_date)
        orbital_data.Controls.Add(last_observation_date)
        orbital_data.Controls.Add(data_arc_in_days)
        orbital_data.Controls.Add(observations_used)
        orbital_data.Controls.Add(orbital_period)

        orbital_data.Controls.Add(morestats)

        orbital_data.Controls.Add(orbit_uncertainty)
        orbital_data.Controls.Add(minimum_orbit_intersection)
        orbital_data.Controls.Add(jupiter_tisserand_invariant)
        orbital_data.Controls.Add(epoch_osculation)
        orbital_data.Controls.Add(eccentricity)
        orbital_data.Controls.Add(semi_major_axis)
        orbital_data.Controls.Add(inclination)
        orbital_data.Controls.Add(ascending_node_longitude)
        orbital_data.Controls.Add(perihelion_distance)
        orbital_data.Controls.Add(perihelion_argument)
        orbital_data.Controls.Add(aphelion_distance)
        orbital_data.Controls.Add(perihelion_time)
        orbital_data.Controls.Add(mean_anomaly)
        orbital_data.Controls.Add(mean_motion)
        orbital_data.Controls.Add(equinox)
        orbital_data.Controls.Add(orbit_class_type)
        orbital_data.Controls.Add(orbit_class_description)
        orbital_data.Controls.Add(orbit_class_range)


        Me.Controls.Add(orbital_data)

        Dim ToolTip As New ToolTip
        ToolTip.ShowAlways = True
        ToolTip.SetToolTip(orbit_id, "The unique ID for the asteroid's orbit data.")
        ToolTip.SetToolTip(orbit_determination_date, "The date and time when the orbital data was determined.")
        ToolTip.SetToolTip(first_observation_date, "The date when the asteroid was first observed.")
        ToolTip.SetToolTip(last_observation_date, "The date when the most recent observation was made.")
        ToolTip.SetToolTip(data_arc_in_days, "The number of days over which observations have been recorded.")
        ToolTip.SetToolTip(observations_used, "The number of observations that were used to calculate the orbital data.")
        ToolTip.SetToolTip(orbital_period, "The time it takes for the asteroid to complete one orbit around the Sun.")
        ToolTip.SetToolTip(orbit_uncertainty, "The uncertainty level In the orbit determination, with 0 indicating no uncertainty.")
        ToolTip.SetToolTip(minimum_orbit_intersection, "The minimum distance between the asteroid's orbit and Earth's orbit.")
        ToolTip.SetToolTip(jupiter_tisserand_invariant, "A number used to classify the asteroid's orbit relative to Jupiter's orbit, which helps indicate whether the object is a comet or an asteroid.")
        ToolTip.SetToolTip(epoch_osculation, "The reference epoch for the orbital elements.")
        ToolTip.SetToolTip(eccentricity, "A measure of the orbit's deviation from circularity, where 0 is a perfect circle.")
        ToolTip.SetToolTip(semi_major_axis, "The average distance from the asteroid to the Sun, or the length of the long axis of the elliptical orbit.")
        ToolTip.SetToolTip(inclination, "The tilt of the asteroid's orbit relative to the plane of the solar system.")
        ToolTip.SetToolTip(ascending_node_longitude, "The longitude Of the ascending node, which Is where the asteroid crosses the plane of the solar system from South to North.")
        ToolTip.SetToolTip(perihelion_distance, "The closest point in the asteroid's orbit to the Sun.")
        ToolTip.SetToolTip(perihelion_argument, "The angle of the perihelion relative to the orbit.")
        ToolTip.SetToolTip(aphelion_distance, "The farthest point in the asteroid's orbit from the Sun.")
        ToolTip.SetToolTip(perihelion_time, "The Julian date when the asteroid last passed its perihelion.")
        ToolTip.SetToolTip(mean_anomaly, "The position of the asteroid in its orbit at a specific time.")
        ToolTip.SetToolTip(mean_motion, "The rate at which the asteroid moves along its orbit.")
        ToolTip.SetToolTip(equinox, "The equinox used to describe the asteroid's orbit.")
        ToolTip.SetToolTip(orbit_class_type, "The type of orbit.")
        ToolTip.SetToolTip(orbit_class_description, "A description of the orbit type.")
        ToolTip.SetToolTip(orbit_class_range, "The range of perihelion distances for this orbit class.")

    End Sub

    Private Sub initializeCloseApproachData()
        close_approach_data = New Panel
        With close_approach_data

            .Size = New Point(500, 1320)
            .Location = New Point(50, 707)
            .Font = New Font(pfc.Families(0), 16, FontStyle.Regular)
        End With

        Me.Controls.Add(close_approach_data)

        Dim close_approach_label = New Label
        With close_approach_label
            .Size = New Size(500, 40)
            .Font = New Font(pfc.Families(0), 16, FontStyle.Regular)
            .Text = "Close Approach Data"
            .TextAlign = ContentAlignment.MiddleCenter
        End With

        close_approach_data.Controls.Add(close_approach_label)

        Dim orbiting_body = New Label
        With orbiting_body
            .Location = New Point(0, 40)
            .Size = New Size(500, 40)
            .Font = New Font(pfc.Families(0), 16, FontStyle.Regular)
            .Text = "Orbiting Body: " & data.item("close_approach_data")(0).item("orbiting_body")
        End With

        close_approach_data.Controls.Add(orbiting_body)

        Dim times As New List(Of Integer)                       'create list and populate it with x axis data for graph
        For Each i In data.item("close_approach_data")
            Dim temp As String = i.item("close_approach_date")
            temp = temp.Substring(0, 4)
            times.Add(temp)
        Next

        Dim relativeVels As New List(Of Double)                 'create list and populate it with y axis data for graph
        For Each i In data.item("close_approach_data")
            relativeVels.Add(i.item("relative_velocity").item("kilometers_per_second"))
        Next

        Dim relativeVelUnits As New List(Of String)             'create and populates list with units for the listbox
        For Each l In data.Item("close_approach_data")(0).item("relative_velocity").keys
            Dim temp As String = ""
            For i = 0 To l.length - 1
                If l(i).Equals("_"c) Then
                    temp += " "
                Else
                    temp += l(i)
                End If
            Next
            relativeVelUnits.Add(temp)
        Next

        ' Create chart
        chart.Dock = DockStyle.None ' Make it fill the form
        chart.Location = New Point(0, 80)
        chart.Size = New Size(500, 500)
        close_approach_data.Controls.Add(chart)

        ' Create chart area
        Dim chartArea As New ChartArea("MainArea")
        chart.ChartAreas.Add(chartArea)

        ' Create series
        series = New Series("Relative Velocity")
        series.ChartType = SeriesChartType.Line
        series.BorderWidth = 2
        series.Color = Color.Green

        ' Add data points
        For i As Integer = 0 To relativeVels.Count - 1
            series.Points.AddXY(times(i), relativeVels(i))
        Next

        'fix the auto padding
        chart.ChartAreas("MainArea").AxisX.Minimum = times(0)
        chart.ChartAreas("MainArea").AxisX.Maximum = times(times.Count - 1)
        chart.ChartAreas("MainArea").AxisX.Interval = Math.Floor((chart.ChartAreas("MainArea").AxisX.Maximum - chart.ChartAreas("MainArea").AxisX.Minimum) / 10)


        chart.ChartAreas("MainArea").AxisY.Minimum = relativeVels.Min - 0.5
        chart.ChartAreas("MainArea").AxisY.Maximum = relativeVels.Max + 0.5

        ' Add series to chart
        chart.Series.Add(series)

        ' Add titles (optional)
        chart.Titles.Add("Relative Velocity Over Time")
        chart.ChartAreas("MainArea").AxisX.Title = "Time (year)"
        chart.ChartAreas("MainArea").AxisY.Title = "Relative Velocity (" & relativeVelUnits(0) & ")"

        ' Font formatting
        chart.ChartAreas("MainArea").AxisY.LabelStyle.Font = New Font("Arial", 7, FontStyle.Regular)

        selection1 = New TextBox
        With selection1
            .Size = New Size(310, 40)
            .Location = New Point(close_approach_data.Width / 2 - selection1.Width / 2 - 20, 580)
            .Font = New Font(pfc.Families(0), 15.5, FontStyle.Regular)
            .ReadOnly = True
            .Text = relativeVelUnits(0)
        End With

        Dim dropdownbutton1 As New Button
        With dropdownbutton1
            .Size = New Size(40, 40)
            .Location = New Point(close_approach_data.Width / 2 + selection1.Width / 2 - 20, 580)
            .Size = New Point(40, 40)
            .Font = New Font(pfc.Families(0), 14, FontStyle.Regular)
            .Image = Image.FromFile(IO.Path.Combine(Directory.GetParent(Application.StartupPath).Parent.Parent.Parent.FullName, "dropdownarrow.png"))
            .FlatStyle = FlatStyle.Flat
        End With

        units1 = New ListBox
        With units1
            .Size = New Size(350, 120)
            .Location = New Point(close_approach_data.Width / 2 - units1.Width / 2 - 1, 620)
            For Each l In relativeVelUnits
                .Items.Add(l)
            Next
            .Font = New Font(pfc.Families(0), 16, FontStyle.Regular)
            .Show()
            .Hide()
        End With
        close_approach_data.Controls.Add(selection1)
        close_approach_data.Controls.Add(dropdownbutton1)
        close_approach_data.Controls.Add(units1)

        AddHandler dropdownbutton1.Click, AddressOf dropdownbutton1_click
        AddHandler units1.SelectedIndexChanged, AddressOf units1_selectedIndexChanged



        Dim times1 As New List(Of Integer)
        For Each i In data.item("close_approach_data")
            Dim temp As String = i.item("close_approach_date")
            temp = temp.Substring(0, 4)
            times1.Add(temp)
        Next

        Dim missDistance As New List(Of Double)
        For Each i In data.item("close_approach_data")
            missDistance.Add(i.item("miss_distance").item("astronomical"))
        Next

        Dim missDistanceUnits As New List(Of String)
        For Each l In data.Item("close_approach_data")(0).item("miss_distance").keys
            Dim temp As String = ""
            For i = 0 To l.length - 1
                If l(i).Equals("_"c) Then
                    temp += " "
                Else
                    temp += l(i)
                End If
            Next
            missDistanceUnits.Add(temp)
        Next

        ' Create chart
        chart1.Dock = DockStyle.None ' Make it fill the form
        chart1.Location = New Point(0, 620)
        chart1.Size = New Size(500, 500)
        close_approach_data.Controls.Add(chart1)

        ' Create chart area
        Dim chartArea1 As New ChartArea("MainArea")
        chart1.ChartAreas.Add(chartArea1)

        ' Create series
        series1 = New Series("Miss Distance")
        series1.ChartType = SeriesChartType.Line
        series1.BorderWidth = 2
        series1.Color = Color.Green

        ' Add data points
        For i As Integer = 0 To missDistance.Count - 1
            series1.Points.AddXY(times(i), missDistance(i))
        Next

        'fix the auto padding
        chart1.ChartAreas("MainArea").AxisX.Minimum = times(0)
        chart1.ChartAreas("MainArea").AxisX.Maximum = times(times.Count - 1)
        chart1.ChartAreas("MainArea").AxisX.Interval = Math.Floor((chart1.ChartAreas("MainArea").AxisX.Maximum - chart1.ChartAreas("MainArea").AxisX.Minimum) / 10)


        chart1.ChartAreas("MainArea").AxisY.Minimum = missDistance.Min - 0.05
        chart1.ChartAreas("MainArea").AxisY.Maximum = missDistance.Max + 0.05

        ' Add series to chart
        chart1.Series.Add(series1)

        ' Add titles (optional)
        chart1.Titles.Add("Miss Distance Over Time")
        chart1.ChartAreas("MainArea").AxisX.Title = "Time (year)"
        chart1.ChartAreas("MainArea").AxisY.Title = "Miss Distance (" & missDistanceUnits(0) & ")"

        ' Font formatting
        chart1.ChartAreas("MainArea").AxisY.LabelStyle.Font = New Font("Arial", 7, FontStyle.Regular)

        selection2 = New TextBox
        With selection2
            .Size = New Size(310, 40)
            .Location = New Point(close_approach_data.Width / 2 - selection1.Width / 2 - 20, 1120)
            .Font = New Font(pfc.Families(0), 15.5, FontStyle.Regular)
            .ReadOnly = True
            .Text = missDistanceUnits(0)
        End With

        Dim dropdownbutton2 As New Button
        With dropdownbutton2
            .Size = New Size(40, 40)
            .Location = New Point(close_approach_data.Width / 2 + selection1.Width / 2 - 20, 1120)
            .Size = New Point(40, 40)
            .Font = New Font(pfc.Families(0), 14, FontStyle.Regular)
            .Image = Image.FromFile(IO.Path.Combine(Directory.GetParent(Application.StartupPath).Parent.Parent.Parent.FullName, "dropdownarrow.png"))
            .FlatStyle = FlatStyle.Flat
        End With

        units2 = New ListBox
        With units2
            .Size = New Size(350, 160)
            .Location = New Point(close_approach_data.Width / 2 - units1.Width / 2 - 1, 1160)
            For Each l In missDistanceUnits
                .Items.Add(l)
            Next
            .Font = New Font(pfc.Families(0), 16, FontStyle.Regular)
            .Show()
            .Hide()
        End With
        close_approach_data.Controls.Add(selection2)
        close_approach_data.Controls.Add(dropdownbutton2)
        close_approach_data.Controls.Add(units2)

        AddHandler dropdownbutton2.Click, AddressOf dropdownbutton2_click
        AddHandler units2.SelectedIndexChanged, AddressOf units2_selectedIndexChanged
    End Sub

    Private Sub initializeFooter()
        footer = New Panel
        With footer
            .Location = New Point(0, 707 + close_approach_data.Height)
            .Size = New Size(Me.Width - 40, 60)
        End With

        Me.Controls.Add(footer)
    End Sub

    Private Sub nasa_jpl_url_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs)  'visits website on link click
        Dim url As String = e.Link.LinkData.ToString()

        Process.Start(New ProcessStartInfo("cmd", $"/c start {url}") With {.CreateNoWindow = True})
    End Sub

    Private Sub dropdownbutton_click(sender As Object, e As EventArgs)              'changes button icon and shows listbox
        units.Location = New Point(selection.Location.X, selection.Location.Y + 40)
        If units.Visible Then
            units.Hide()
        Else
            units.Show()
        End If
    End Sub

    Private Sub units_selectedIndexChanged(sender As Object, e As EventArgs)     'updates the selection box and hides the listbox on new listbox selection
        selection.Text = units.Text
        units.Hide()
        estimated_diameter.Text = "Estimated Diameter: " &
            vbCrLf & "Est. Min: " & data.item("estimated_diameter").item(selection.Text.ToLower).item("estimated_diameter_min") &
            vbCrLf & "Est. Max: " & data.item("estimated_diameter").item(selection.Text.ToLower).item("estimated_diameter_max")
    End Sub

    Private Sub form1_scroll(sender As Object, e As EventArgs) Handles Me.Scroll        'updates listbox positions on scroll
        If newWindow Then
            units.Location = New Point(selection.Location.X, selection.Location.Y + 40)
            units1.Location = New Point(selection1.Location.X, selection1.Location.Y + 40)
            units2.Location = New Point(selection2.Location.X, selection2.Location.Y + 40)
        End If
    End Sub

    Private Sub form1_mousewheel(sender As Object, e As EventArgs) Handles Me.MouseWheel            'update listbox positions on mousewheel scroll
        If newWindow Then
            units.Location = New Point(selection.Location.X, selection.Location.Y + 40)
            units1.Location = New Point(selection1.Location.X, selection1.Location.Y + 40)
            units2.Location = New Point(selection2.Location.X, selection2.Location.Y + 40)
        End If
    End Sub

    Private Sub morestats_click(sender As Object, e As EventArgs)       'expands the orbital data panel and shows the other stats, or hides depending on if its shown already or not
        units.Location = New Point(selection.Location.X, selection.Location.Y + 40)
        units1.Location = New Point(selection1.Location.X, selection1.Location.Y + 40)
        units2.Location = New Point(selection2.Location.X, selection2.Location.Y + 40)

        If Not moreStatsExpand Then
            orbital_data.Size = New Point(500, 1720)
            morestats.Image = Image.FromFile(IO.Path.Combine(Directory.GetParent(Application.StartupPath).Parent.Parent.Parent.FullName, "dropdownarrowUP.png"))
            moreStatsExpand = True
        Else
            orbital_data.Size = New Point(500, 440)
            morestats.Image = Image.FromFile(IO.Path.Combine(Directory.GetParent(Application.StartupPath).Parent.Parent.Parent.FullName, "dropdownarrow.png"))
            moreStatsExpand = False
        End If

        If close_approach_data.Height > orbital_data.Height Then
            footer.Location = New Point(0, 707 + close_approach_data.Height)
        Else
            footer.Location = New Point(0, 1300)
        End If
    End Sub

    Private Sub dropdownbutton1_click(sender As Object, e As EventArgs)         'change button icon and show listbox
        units1.Location = New Point(selection1.Location.X, selection1.Location.Y + 40)
        If units1.Visible Then
            units1.Hide()
        Else
            units1.Show()
        End If
    End Sub

    Private Sub units1_selectedIndexChanged(sender As Object, e As EventArgs)       'updates the selection box and hides the listbox on new listbox selection
        Dim position = Me.AutoScrollPosition
        selection1.Text = units1.Text

        chart.Series("Relative Velocity").Points.Clear()

        Dim selection1Fixed As String = ""
        For i = 0 To selection1.Text.Length - 1
            If selection1.Text(i) = " "c Then
                selection1Fixed += "_"
            Else
                selection1Fixed += selection1.Text(i)
            End If
        Next

        Dim times As New List(Of Integer)
        For Each i In data.item("close_approach_data")
            Dim temp As String = i.item("close_approach_date")
            temp = temp.Substring(0, 4)
            times.Add(temp)
        Next

        Dim relativeVels As New List(Of Double)
        For Each i In data.item("close_approach_data")
            relativeVels.Add(i.item("relative_velocity").item(selection1Fixed))
        Next

        ' Add data points
        For i As Integer = 0 To relativeVels.Count - 1
            series.Points.AddXY(times(i), relativeVels(i))
        Next

        'fix the auto padding
        If selection1.Text = "kilometers per second" Then
            chart.ChartAreas("MainArea").AxisY.Minimum = relativeVels.Min - 0.5
            chart.ChartAreas("MainArea").AxisY.Maximum = relativeVels.Max + 0.5
        Else
            chart.ChartAreas("MainArea").AxisY.Minimum = relativeVels.Min - 500
            chart.ChartAreas("MainArea").AxisY.Maximum = relativeVels.Max + 500
        End If

        ' Add titles (optional)
        chart.ChartAreas("MainArea").AxisY.Title = "Relative Velocity (" & selection1.Text & ")"
        units1.Hide()
        Me.AutoScrollPosition = New Point(0, 700)
    End Sub

    Private Sub dropdownbutton2_click(sender As Object, e As EventArgs)         'change button icon and show listbox
        units2.Location = New Point(selection2.Location.X, selection2.Location.Y + 40)
        If units2.Visible Then
            units2.Hide()
        Else
            units2.Show()
        End If
    End Sub

    Private Sub units2_selectedIndexChanged(sender As Object, e As EventArgs)           'updates the selection box and hides the listbox on new listbox 
        Dim position = Me.AutoScrollPosition
        selection2.Text = units2.Text

        chart1.Series("Miss Distance").Points.Clear()

        Dim selection1Fixed As String = ""
        For i = 0 To selection2.Text.Length - 1
            If selection2.Text(i) = " "c Then
                selection1Fixed += "_"
            Else
                selection1Fixed += selection2.Text(i)
            End If
        Next

        Dim times As New List(Of Integer)
        For Each i In data.item("close_approach_data")
            Dim temp As String = i.item("close_approach_date")
            temp = temp.Substring(0, 4)
            times.Add(temp)
        Next

        Dim missDistance As New List(Of Double)
        For Each i In data.item("close_approach_data")
            missDistance.Add(i.item("miss_distance").item(selection1Fixed))
        Next

        ' Add data points
        For i As Integer = 0 To missDistance.Count - 1
            series1.Points.AddXY(times(i), missDistance(i))
        Next

        'fix the auto padding
        If selection2.Text = "astronomical" Then
            chart1.ChartAreas("MainArea").AxisY.Minimum = missDistance.Min - 0.05
            chart1.ChartAreas("MainArea").AxisY.Maximum = missDistance.Max + 0.05
        ElseIf selection2.Text = "lunar" Then
            chart1.ChartAreas("MainArea").AxisY.Minimum = missDistance.Min - 10
            chart1.ChartAreas("MainArea").AxisY.Maximum = missDistance.Max + 10
        Else
            chart1.ChartAreas("MainArea").AxisY.Minimum = missDistance.Min - 1000000
            chart1.ChartAreas("MainArea").AxisY.Maximum = missDistance.Max + 1000000
        End If

        ' Add titles (optional)
        chart1.ChartAreas("MainArea").AxisY.Title = "Miss Distance (" & selection2.Text & ")"
        units2.Hide()
        Me.AutoScrollPosition = New Point(0, 1200)
    End Sub
End Class
