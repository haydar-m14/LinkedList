Structure NodeRec 
    Dim TeamName As String
    Dim Pointer As String 
End Structure 

    Dim TeamList(5) As NodeRec
    Dim Pointer, CurrentPointer, HeadPointer, FreePointer, PreviousPointer, NextFreeNodeAddress As Integer
    Sub Main()
        Console.Title = "Linked List"
        Call CreateLinkedList()
        Call Menu()
    End Sub
    'The CreateLinkedList operation links all nodes to form the free list and initialises the HeadPointer and FreePointer.
    Sub CreateLinkedList()
        HeadPointer = 0
        FreePointer = 1
        For Index = 1 To 4
            TeamList(Index).Pointer = Index + 1
        Next
        TeamList(5).Pointer = 0
    End Sub
    Sub Menu()
        Try
            Dim Time As Date
            Dim Name As String
            Time = Date.Now
            Dim Choice As Integer
            Console.ForegroundColor = ConsoleColor.Green
            Console.WriteLine("                                                  " & Time)
            Console.ForegroundColor = ConsoleColor.White
            Console.WriteLine("1. Add Team Name")
            Console.WriteLine("2. Search Team Name")
            Console.WriteLine("3. Delete Team Name")
            Console.WriteLine("4. Display Linked List")
            Console.WriteLine("5. End")
            Console.WriteLine("Select one of the options above: ")
            Choice = Console.ReadLine
            Select Case Choice
                Case 1
                    Console.Clear()
                    Console.WriteLine("Enter the new team name: ")
                    Name = Console.ReadLine
                    Call AddName(Name)
                Case 2
                    Console.Clear()
                    Console.WriteLine("Enter the team name for searching: ")
                    Name = Console.ReadLine
                    Call SearchName(Name)
                Case 3
                    Console.Clear()
                    Console.WriteLine("Enter the team name to be deleted: ")
                    Name = Console.ReadLine
                    Call RemoveName(Name)
                Case 4
                    Call Display()
                Case 5
                    End
                Case Else
                    Console.ForegroundColor = ConsoleColor.Red
                    Console.WriteLine("Select an option between 1 and 5 inclusive!")
                    Console.ReadKey()
                    Console.Clear()
                    Call Menu()
            End Select
        Catch ex As Exception
            Call TryCatchError(ex)
        End Try
    End Sub
    Sub AddName(ByVal Name As String)
        Try
            Dim Choice As Integer
            Dim Name_2 As String
            'To place item at first free node
            If FreePointer <> 0 Then
                TeamList(FreePointer).TeamName = Name
                NextFreeNodeAddress = TeamList(FreePointer).Pointer 'In order to keep track of the next free node
                CurrentPointer = HeadPointer 'Searching for the position of where the string is to be inserted
                While TeamList(CurrentPointer).TeamName < Name And CurrentPointer <> 0
                    PreviousPointer = CurrentPointer
                    CurrentPointer = TeamList(CurrentPointer).Pointer
                End While
                If CurrentPointer = HeadPointer Then 'If string to be inserted at start of linked list with or without nodes
                    TeamList(FreePointer).Pointer = HeadPointer
                    HeadPointer = FreePointer
                Else
                    'If the node is to be inserted between exisiting or after all nodes
                    TeamList(FreePointer).Pointer = CurrentPointer
                    TeamList(PreviousPointer).Pointer = FreePointer
                End If
                ' Set freepointer to the new free node
                FreePointer = NextFreeNodeAddress
            Else
                Console.WriteLine("There is no free space!")
            End If
            While Choice <> 1 And Choice <> 2 'In order to add more data into the list
                Console.WriteLine("To add more data, press 1 and enter. To return to Menu, press 2 and enter.")
                Choice = Console.ReadLine
                If Choice = 1 Then
                    Console.Clear()
                    Console.ForegroundColor = ConsoleColor.White
                    Console.WriteLine("Enter the new team name: ")
                    Name_2 = Console.ReadLine
                    Console.Beep()
                    Call AddName(Name_2)
                ElseIf Choice = 2 Then
                    Console.Clear()
                    Call Menu()
                Else
                    Console.ForegroundColor = ConsoleColor.Red
                End If
            End While
        Catch ex As Exception
            Call TryCatchError(ex)
        End Try
    End Sub
    Sub SearchName(ByVal Name As String)
        Try
            Dim IsFound As Boolean = False
            CurrentPointer = HeadPointer 'In order to go through each node
            While CurrentPointer <> 0 And IsFound = False 'In order to go through the whole list
                If TeamList(CurrentPointer).TeamName = Name Then
                    Console.WriteLine()
                    IsFound = True
                End If
                CurrentPointer = TeamList(CurrentPointer).Pointer
            End While
            If IsFound = True Then
                Console.WriteLine(Name & " is found at address " & CurrentPointer & " and it points to the node at address " & TeamList(CurrentPointer).Pointer & ".")
            ElseIf IsFound = False Then
                Console.WriteLine("Item is not in linked list!")
            End If
            Console.ReadKey()
            Console.Clear()
            Call Menu()
        Catch ex As Exception
            Call TryCatchError(ex)
        End Try
    End Sub
    Sub RemoveName(ByVal Name As String)
        Try
            Dim IsFound As Boolean = False
            Dim Choice As Integer
            Dim Name_2 As String
            CurrentPointer = HeadPointer
            If TeamList(CurrentPointer).TeamName = Name Then
                HeadPointer = TeamList(CurrentPointer).Pointer
                TeamList(CurrentPointer).Pointer = FreePointer
                FreePointer = CurrentPointer
                IsFound = True
            Else
                While CurrentPointer <> 0 And IsFound = False
                    PreviousPointer = CurrentPointer
                    CurrentPointer = TeamList(CurrentPointer).Pointer
                    If TeamList(CurrentPointer).TeamName = Name Then
                        TeamList(PreviousPointer).Pointer = TeamList(CurrentPointer).Pointer
                        TeamList(CurrentPointer).Pointer = FreePointer
                        FreePointer = CurrentPointer
                        IsFound = True
                        Console.WriteLine("Team name " & Name & " was found and deleted!")
                    End If
                End While
            End If
            If IsFound = False Then
                Console.WriteLine("Team name to be deleted was not found in linked list!")
                Console.ReadKey()
            End If
            While Choice <> 1 And Choice <> 2 'Option to delete more data
                Console.WriteLine("To delete more data, press 1 and enter. To return to Menu, press 2 and enter.")
                Choice = Console.ReadLine
                If Choice = 1 Then
                    Console.Clear()
                    Console.ForegroundColor = ConsoleColor.White
                    Console.WriteLine("Enter the  team name to be deleted: ")
                    Name_2 = Console.ReadLine
                    Call RemoveName(Name_2)
                ElseIf Choice = 2 Then
                    Console.Clear()
                    Call Menu()
                Else
                    Console.ForegroundColor = ConsoleColor.Red
                End If
            End While
        Catch ex As Exception
            Call TryCatchError(ex)
        End Try
    End Sub
    Sub TryCatchError(ByVal ex)
        Console.ForegroundColor = ConsoleColor.Red
        Console.WriteLine(ex.Message)
        Console.ReadKey()
        Console.Clear()
        Call Menu()
    End Sub
    Sub Display()
        Try
            Console.Clear()
            CurrentPointer = HeadPointer
            Console.ForegroundColor = ConsoleColor.Yellow
            While CurrentPointer <> 0
                Console.WriteLine("Current Pointer: " & CurrentPointer & " | Team name: " & TeamList(CurrentPointer).TeamName & "| Points to: " & TeamList(CurrentPointer).Pointer)
                CurrentPointer = TeamList(CurrentPointer).Pointer
            End While
            Console.ReadKey()
            Console.Clear()
            Call Menu()
        Catch ex As Exception
            Call TryCatchError(ex)
        End Try

    End Sub
