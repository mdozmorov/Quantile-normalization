Imports System.IO
Public Class Form1

    Private Sub btnQuantile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnQuantile.Click
        'Quantile normalization of a matrix with missing values.
        'Matrix should have row and column headers. Numbers should be >=0 
        'Designed for normalizing gene expression matrix, with rows=experiments, columns=genes
        Dim GEM(,) As Integer, GEMline() As Integer, Keys() As Integer, GEMaverage() As Integer, NumOfNonnull As Integer
        Dim Line1 As String, sp() As String, UB, MaxSamples, i, j, NullPlaceholder As Integer
        Dim PathName As String = "F:\"
        Dim FileNameIn As String = "test.txt"
        Dim FileNameOut As String = "testout2.txt"
        '===First, get matrix dimensions
        Using reader As StreamReader = New StreamReader(PathName & FileNameIn)
            Line1 = reader.ReadLine()                           'read header
            sp = Line1.Split(vbTab) : UB = sp.GetUpperBound(0)  'Get horizontal dimensions
            While Not reader.EndOfStream
                Line1 = reader.ReadLine() : MaxSamples += 1     'Get vertical dimensions
            End While
        End Using
        Debug.Print(TimeOfDay & " Matrix dimensions done")
        '===Second, read in full matrix
        NullPlaceholder = -1                                    'Dummy number to signal if a value is missing
        ReDim GEM(MaxSamples - 1, UB - 1)                       'Prepare array to store full matrix, less one due to row/column headers
        MaxSamples = 0                                          'Reuse row counter
        Using reader As StreamReader = New StreamReader(PathName & FileNameIn)
            Line1 = reader.ReadLine()                           'read header
            While Not reader.EndOfStream
                Line1 = reader.ReadLine() : MaxSamples += 1     'Go through each line
                sp = Line1.Split(vbTab)
                For i = 1 To UB                                 'within line, through each gene
                    If sp(i) <> vbNullString Then               'Check if missing
                        GEM(MaxSamples - 1, i - 1) = sp(i)      'Store value if present
                    Else
                        GEM(MaxSamples - 1, i - 1) = NullPlaceholder      'Dummy number if missing
                    End If
                Next i
            End While
        End Using
        Debug.Print(TimeOfDay & " Reading in full matrix done")
        '===Sort full matrix line by line
        ReDim GEMline(UB - 1)                                   'Array to store data for one string
        For j = 1 To MaxSamples                                 'Go through each sample
            For i = 1 To UB                                     'and through each gene within it
                GEMline(i - 1) = GEM(j - 1, i - 1)              'get one full line
            Next
            Array.Sort(GEMline)                                 'Sort it, dummy numbers will be first
            For i = 1 To UB                                     'Replace values in the full matrix, to make it sorted
                GEM(j - 1, i - 1) = GEMline(i - 1)
            Next
        Next
        Debug.Print(TimeOfDay & " Sorting full matrix done")
        '===Count average distribution and output it to a file
        ReDim GEMaverage(UB - 1)                                    'Array to store average distribution
        Using writer As StreamWriter = New StreamWriter(PathName & FileNameOut & "average distribution.txt")
            For i = 1 To UB                                         'Now go through each gene columnwise
                NumOfNonnull = 0                                    'Number of nonmissing values
                For j = 1 To MaxSamples                             'For a gene go through each experiment
                    'Only add up and count values that are non-missing
                    If GEM(j - 1, i - 1) <> NullPlaceholder Then GEMaverage(i - 1) += GEM(j - 1, i - 1) : NumOfNonnull += 1
                Next
                If NumOfNonnull > 0 Then                            'If at least some are nonmissing
                    GEMaverage(i - 1) = GEMaverage(i - 1) / NumOfNonnull    'Count average
                Else
                    GEMaverage(i - 1) = NullPlaceholder             'or put placeholder to mark them missing
                End If
            Next
            For i = 1 To UB : writer.WriteLine(GEMaverage(i - 1)) : Next    'Save average distribution
        End Using
        Debug.Print(TimeOfDay & " Average matrix done")
        '===Read in original matrix line by line, sort keeping original order, replace with average, sort back and store
        ReDim GEMline(UB - 1) : ReDim Keys(UB - 1)      'Array for storing current line values
        For i = 1 To UB : Keys(i - 1) = i - 1 : Next    'Keys hold 1,2,3 index
        MaxSamples = 0
        Using reader As StreamReader = New StreamReader(PathName & FileNameIn)
            Using writer As StreamWriter = New StreamWriter(PathName & FileNameOut)
                Line1 = reader.ReadLine()               'read header
                writer.WriteLine(Line1)                 'and write it to the output file
                While Not reader.EndOfStream
                    Line1 = reader.ReadLine() : MaxSamples += 1     'Get vertical dimensions
                    sp = Line1.Split(vbTab) : Debug.Print(TimeOfDay & " Final processing of sample " & MaxSamples)
                    For i = 1 To UB                     'For each line go through each gene
                        If sp(i) <> "" Then
                            GEMline(i - 1) = sp(i)      'Store value if present
                        Else
                            GEMline(i - 1) = NullPlaceholder    'Or dummy if not
                        End If
                    Next i
                    Array.Sort(GEMline, Keys)           'Sort the line, Keys keep order
                    For i = 1 To UB                     'Go through sorted line
                        If GEMline(i - 1) <> NullPlaceholder Then   'If the original value is not null
                            GEMline(i - 1) = GEMaverage(i - 1)      'Replace it with the average distribution
                        Else
                            GEMline(i - 1) = NullPlaceholder        'or with dummy
                        End If
                    Next
                    Array.Sort(Keys, GEMline)           'Sort back to the original order in Keys
                    writer.Write(sp(0) & vbTab)         'Write GEM header
                    For i = 1 To UB                     'Go through each gene
                        If GEMline(i - 1) <> NullPlaceholder Then   'If non-missing
                            writer.Write(GEMline(i - 1) & vbTab)    'output it
                        Else
                            writer.Write(vbNullString & vbTab)      'or output null string
                        End If
                    Next
                    writer.Write(vbCrLf)                'end of the line
                End While
            End Using
        End Using
        MsgBox("Quantile normalization done")
    End Sub
End Class
