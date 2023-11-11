# PyPoll
Sub Votes()

    Dim TotalNumberofVotesT7 As Long
    Dim ListCandidateRecieveVoteT9 As Long
    Dim PercentageofVotesEachWonT11 As Integer
    Dim TotalNumberofVotesWinT13 As Long
    Dim WinnerOfElectionT14 As String
    Dim i As Long
    Dim j As Integer
    Dim Candidate1 As String
    Dim Candidate2 As String
    Dim Candidate3 As String
    Dim Count1 As Long
    Dim Count2 As Long
    Dim Count3 As Long
    Dim R As Integer
    Dim Total_Count As Long
    Dim PercentageCandidate1 As Integer
    Dim PercentageCandidate2 As Integer
    Dim PercentageCandidate3 As Integer
    Dim Winner As String
    
   'Total Votes
   TotalNumberofVotesT7 = Application.Worksheet.Count(Range(Cells(2, 3), Cells(Cells(Rows.Count, 1).End(xlUp).Row, 3)))
    
    R = 0
    For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
             
         If Cells(i, 3) <> Cells(i + 1, 3) Then
         Candidate = Cells(i, 3)
         R = R + 1
         Cells(R, 5) = Candidate
         Else
    End If
    Next i
    
    'List of Candidate Recieved a Vote
    Candidate1 = Cells(1, 5)
    Candidate2 = Cells(2, 5)
    Candidate3 = Cells(3, 5)
          
   For j = 2 To Cells(Rows.Count, 1).End(xlUp).Row
            
         Count1 = 0
         Count2 = 0
         Count3 = 0
         
         If Cells(j, 3) = Candidate1 Then
            Count1 = Count1 + 1
         
         ElseIf Cells(j, 3) = Candidate2 Then
            Count2 = Count2 + 1
                     
         Else
            If Cells(j, 3) = Candidate3 Then
            Count3 = Count3 + 1
        End If
        Next j
    
    
'The percentage of votes each candidate won
    Total_Count = Count1 + Count2 + Count3
    PercentageCandidate1 = Count1 / Total_Count
    PercentageCandidate2 = Count2 / Total_Count
    PercentageCandidate3 = Count2 / Total_Count
    
'Winner Candidate

    WinnerCount = Application.WorksheetFunction.Max(Count1, Count2, Count3)
    
    If WinnerCount = Count1 Then
        Winner = Candidate1
        Else
        Winner = Candidate2
        Else
        Winner = Candidate3
    End If
    
' Report

    Range(T7) = TotalNumberofVotesT7
    Range(T9) = List(Candidate1, PercentageCandidate1, Count1)
    Range(T10) = List(Candidate2, PercentageCandidate2, Count2)
    Renage(T11) = List(Candidate3, PercentageCandidate2, Count3)
    Range(T16) = Winner

End Sub
