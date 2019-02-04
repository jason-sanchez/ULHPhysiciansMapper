'ShelbyMapper.exe
'20091214 Start project
Imports System
Imports System.IO
Imports System.Collections
Imports System.data.sqlclient
Module Module1
    Public objIniFile As New INIFile("c:\newfeeds\HL7Mapper.ini")
    Dim strInputDirectory As String = ""
    Dim strOutputDirectory As String = ""
    Dim strOutputSubDirectory As String = ""
    Dim strMapperFile As String = ""
    Dim strLogFile As String = ""
    Dim myHT As New Hashtable
    Dim filecounter As Integer = 0
    Sub Main()
        Dim strLTWOutput As String = ""
        Dim boolUseIt As Boolean = False '20081120
        'declarations for split function

        Dim dictNVP As New Hashtable
        Dim s As String
        Dim s1 As String
        
        Dim dir As String
        strInputDirectory = objIniFile.GetString("physicians", "physiciansinputDir", "(none)") 'c:\feeds\HL7\physician\
        strOutputDirectory = objIniFile.GetString("physicians", "physiciansoutputdirectory", "(none)") 'c:\feeds\ltw\physicians\
        strMapperFile = objIniFile.GetString("physicians", "physiciansmapper", "(none)")
        strLogFile = objIniFile.GetString("Settings", "logs", "(none)") 'c:\feeds\logs\

        'Dim LogFile As StreamWriter = File.AppendText(strLogFile & "ShelbyMapperLog.txt")

        Dim delimStr As String = "|"
        Dim delimiter As Char() = delimStr.ToCharArray()
        Dim theFile As FileInfo
        Dim OBXCounter As Integer = 0
        Dim OBRCounter As Integer = 0
        Dim NTECounter As Integer = 0
        '20100719
        Dim RXCCounter As Integer = 0

        'declarations for stream reader
        Dim strLine As String = ""
        'setup directory
        Dim dirs As String() = Directory.GetFiles(strInputDirectory, "HL7.*")

        '20080714 create the reference hash table with mapper data
        CreateHashTable()

        For Each dir In dirs
            filecounter = filecounter + 1
            '20091117 change to 1001 from 201
            If filecounter >= 1001 Then Exit For
            OBXCounter = 0
            NTECounter = 0
            OBRCounter = 0
            RXCCounter = 0

            strLTWOutput = ""
            theFile = New FileInfo(dir)
            'LogFile.WriteLine(theFile.FullName)
            'If theFile.Extension <> ".$#$" Then

            '1.set up the streamreader to get a file
            'myfile = File.OpenText(dir)
            Dim myfile As StreamReader = New StreamReader(theFile.FullName)
            'LogFile.WriteLine(myfile)
            'and read the first line
            'strLine = myfile.ReadLine()

            Do
                Dim myArray As String() = Nothing
                Dim TestPos As Integer = 0
                strLine = myfile.ReadLine()
                Dim segId As String = ""
                Dim segIDFull As String = ""
                Dim counter As Integer = 0

                Dim segname As String = ""
                'get the segment Id which is the first three Characters of the string
                segId = Mid(strLine, 1, 3)

                If segId = "MSH" Then
                    counter = +1
                    If counter = 1 Then
                        counter = +1
                    End If
                End If

                If segId = "OBX" Then
                    OBXCounter = OBXCounter + 1
                    If OBXCounter = 1 Then
                        OBXCounter = +1
                    End If
                End If

                If segId = "NTE" Then
                    NTECounter = NTECounter + 1
                    If NTECounter = 1 Then
                        NTECounter = +1
                    End If
                End If

                If segId = "OBR" Then
                    OBRCounter = OBRCounter + 1
                    If OBRCounter = 1 Then
                        OBRCounter = +1
                    End If
                End If

                If segId = "RXC" Then
                    RXCCounter = RXCCounter + 1
                    If RXCCounter = 1 Then
                        RXCCounter = +1
                    End If
                End If
                
                'LogFile.WriteLine("---------------------------------------")
                If strLine <> "" Then
                    myArray = strLine.Split(delimiter)
                    'add array key and item to hashtable
                    For Each s In myArray
                        'counter += 1

                        'If s <> "" Then
                        Dim mySubArray As String() = Nothing
                        mySubArray = s.Split("^")
                        segIDFull = segId & "_" & counter

                        If myHT.Item(segIDFull) <> "" Then
                            segname = myHT.Item(segIDFull)
                            boolUseIt = True ''20081120
                        Else
                            segname = segIDFull
                            boolUseIt = False ''20081120
                        End If


                        If boolUseIt = True Then '20081120
                            '=================================================================================
                            If segId = "OBX" Then
                                'LogFile.Write(segname & "_" & OBXCounter & "=")
                                strLTWOutput = strLTWOutput & segname & BuildSegCounter(OBXCounter) & "="


                            ElseIf segId = "NTE" Then
                                'LogFile.Write(segname & "_" & OBXCounter & "=")
                                strLTWOutput = strLTWOutput & segname & BuildSegCounter(NTECounter) & "="


                            ElseIf segId = "OBR" Then
                                'LogFile.Write(segname & "_" & OBXCounter & "=")
                                strLTWOutput = strLTWOutput & segname & BuildSegCounter(OBRCounter) & "="

                            ElseIf segId = "RXC" Then
                                'LogFile.Write(segname & "_" & OBXCounter & "=")
                                strLTWOutput = strLTWOutput & segname & BuildSegCounter(RXCCounter) & "="

                            Else
                                'LogFile.Write(segname & "=")
                                strLTWOutput = strLTWOutput & segname & "="
                            End If
                            '=================================================================================
                        End If '20081120

                        '20080721===========================================================================
                        If segIDFull = "MSH_5" Then
                            'strOutputSubDirectory = ""
                            'CreateOutputSubDirectory(s)


                        End If
                        '20080721 - end====================================================================
                        If boolUseIt Then '20081120
                            strLTWOutput = strLTWOutput & s & vbCrLf
                        End If '20081120
                        TestPos = InStr(1, s, "^")
                        '20091201 - make testpos >= 0 instead of > 0
                        If TestPos >= 0 Then
                            Dim subCounter As Integer = 0
                            For Each s1 In mySubArray
                                subCounter += 1
                                segIDFull = segId & "_" & counter & "_" & subCounter
                                If myHT.Item(segIDFull) <> "" Then
                                    segname = myHT.Item(segIDFull)
                                    boolUseIt = True '20081120
                                Else
                                    segname = segIDFull
                                    boolUseIt = False '20081120
                                End If
                                If boolUseIt Then '20081120
                                    '=================================================================================
                                    If segId = "OBX" Then
                                        'LogFile.Write(segname & "_" & OBXCounter & "=")
                                        strLTWOutput = strLTWOutput & segname & BuildSegCounter(OBXCounter) & "="


                                    ElseIf segId = "NTE" Then
                                        'LogFile.Write(segname & "_" & OBXCounter & "=")
                                        strLTWOutput = strLTWOutput & segname & BuildSegCounter(NTECounter) & "="


                                    ElseIf segId = "OBR" Then
                                        'LogFile.Write(segname & "_" & OBXCounter & "=")
                                        strLTWOutput = strLTWOutput & segname & BuildSegCounter(OBRCounter) & "="

                                    ElseIf segId = "RXC" Then
                                        'LogFile.Write(segname & "_" & OBXCounter & "=")
                                        strLTWOutput = strLTWOutput & segname & BuildSegCounter(RXCCounter) & "="

                                    Else
                                        If boolUseIt Then '20081120
                                            strLTWOutput = strLTWOutput & segname & "="
                                        End If '20081120
                                    End If

                                    'LogFile.WriteLine(s1)
                                    strLTWOutput = strLTWOutput & s1 & vbCrLf
                                End If '20081120
                                '=================================================================================
                            Next
                        End If
                        'End If
                        counter += 1
                    Next

                End If
            Loop Until (strLine Is Nothing)

            myfile.Close()

            'code to write ltw file to disk ================================
            'LogFile.Write(strLTWOutput)
            CreateOutputFile(strLTWOutput)

            theFile.Delete()
            '===============================================================
            'End If
        Next

        'LogFile.Close()
    End Sub
    Public Sub CreateOutputFile(ByVal strLTWOutput As String)
        'Function to create an HL7 output file

        Dim line As String = ""
        Dim objTStreamCounter As Object
        Dim intCounter As Integer = 0

        Dim filename As String
        Dim objTStreamOutput As Object

        'If the file does not exist, create it.
        If Not File.Exists(strOutputDirectory & "counter.txt") Then
            objTStreamCounter = File.CreateText(strOutputDirectory & "counter.txt")
            objTStreamCounter.WriteLine("000")
            objTStreamCounter.Close()
        End If

        'read the present file number for counter.Txt. convert it to an integer and increment it.
        objTStreamCounter = New StreamReader(strOutputDirectory & "counter.txt")

        line = objTStreamCounter.readline
        intCounter = CInt(line)
        intCounter = intCounter + 1
        If intCounter >= 100000 Then intCounter = 0
        objTStreamCounter.Close()

        'write the LTW file to strOutputDirectory a new file is created
        'If strOutputSubDirectory <> "" Then
        'filename = strOutputDirectory & "\" & strOutputSubDirectory & "\LTW." & padleft(Str(intCounter), 3)
        'Else
        filename = strOutputDirectory & "\LTW." & padleft(Str(intCounter), 3)
        'End If

        objTStreamOutput = File.AppendText(filename)
        objTStreamOutput.Write(strLTWOutput)
        objTStreamOutput.Close()

        'update the counter file
        objTStreamCounter = New StreamWriter(strOutputDirectory & "counter.txt")
        objTStreamCounter.WriteLine(padleft(Str(intCounter), 3))
        objTStreamCounter.Close()
    End Sub


    Public Function padleft(ByRef inputStr As String, ByRef strLength As Short) As String
        'pad an input string with zeros based on desired strLength
        Dim varLength As Short
        Dim strOutput As String
        Dim i As Short


        strOutput = ""
        varLength = Len(Trim(inputStr))
        For i = 1 To ((strLength - varLength))
            strOutput = strOutput & "0"
        Next
        strOutput = strOutput & Trim(inputStr)
        padleft = strOutput


    End Function
    Public Sub CreateHashTable()
        'Dim s As String
        Dim segID As String = ""
        Dim segDescription As String = ""
        Dim delimStr As String = "="
        Dim strLine As String = ""
        Dim delimiter As Char() = delimStr.ToCharArray()
        Using sr As StreamReader = New StreamReader(strMapperFile)
            Dim line As String
            ' Read and display the lines from the file until the end 
            ' of the file is reached.
            Do

                line = sr.ReadLine()

                If line <> "" Then
                    Dim myArray As String() = Nothing
                    myArray = line.Split(delimiter)
                    segID = myArray(0)
                    segDescription = myArray(1)
                    If myHT.ContainsKey(segID) = False Then
                        myHT.Add(segID, segDescription)
                    End If
                End If
                'Next
            Loop Until line Is Nothing
            sr.Close()
        End Using
        'TextBox1.AppendText(myHT.Item("MSH_3"))

    End Sub
    Public Sub CreateOutputSubDirectory(ByVal s As String)
        If s <> "" Then
            Dim di As DirectoryInfo = New DirectoryInfo(strOutputDirectory & "\" & s)
            If di.Exists Then
                strOutputSubDirectory = s
            Else
                di.Create()
                strOutputSubDirectory = s
            End If
        End If
    End Sub
    Public Function BuildSegCounter(ByRef theCounter As Integer) As String
        If theCounter <= 1 Then
            BuildSegCounter = ""
        End If
        If theCounter > 1 And theCounter < 10 Then
            BuildSegCounter = "_000" & theCounter
        End If

        If theCounter >= 10 And theCounter < 100 Then
            BuildSegCounter = "_00" & theCounter
        End If

        If theCounter >= 100 And theCounter < 1000 Then
            BuildSegCounter = "_0" & theCounter
        End If

        If theCounter >= 1000 Then
            BuildSegCounter = "_" & theCounter
        End If
    End Function
End Module
