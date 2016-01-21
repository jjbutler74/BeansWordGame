'
' ControlArrayUtils.VB - Provide VB6 like support of control arrays
' 
' Author -      P Krawec, Inovision Software Solutions, Inc.
' Date   -      September 1, 2005
'
' This code is free software; you can redistribute it and/or
' modify it any way you want.
' This library is distributed in the hope that it will be useful,
' but without any waranty. 
'
' Any additions or modifications please email to: 
'   kepler77@gmail.com
'
'
' How it works:
' VB.NET does not support control arrays as VB6 did with the index property.
' There is a way to get around this with minimal coding. 
' It requires a standard naming convention of the controls.
' Then the call to one function to search the screen based on that naming 
' convention to fill an array.
' 
' Example:
' To have a control array of 5 TextBoxes, you would have to name the textboxes 
' with the following convention:
'
' myTextBox0, myTextBox1, myTextBox2,myTextBox3, myTextBox4
' 
'Then in code, you would call: 
'dim myTextBoxArray as TextBox() = _
'    ControlArrayUtils.getControlArray(FORM, "myTextBox")
'
'The function will then search through all controls directly on the form that 
'match the name and have an integer value following it.
'Returning the new control array.
'
'ToDo: 
'     Add exception handling
'     Support recursive searching for controls within other controls.
'
Public Class ControlArrayUtils

    'Converts same type of controls on a form to a control array by 
    'using the notation ControlName_1, ControlName_2, where the _ 
    'can be replaced by any separator string
    Public Shared Function getControlArray( _
        ByVal frm As Windows.Forms.Form, _
        ByVal controlName As String, _
        Optional ByVal separator As String = "") _
        As System.Array
        Dim i As Short
        Dim startOfIndex As Short
        Dim alist As New ArrayList
        Dim controlType As System.Type

        Dim ctl As System.Windows.Forms.Control
        Dim ctrls() As System.Windows.Forms.Control
        Dim strSuffix As String
        Dim maxIndex As Short = -1 'Default

        'Loop through all controls, looking for controls with the 
        'matching name pattern
        'Find the highest indexed control
        For Each ctl In frm.Controls
            startOfIndex = ctl.Name.ToLower.IndexOf(controlName.ToLower & separator)
            If startOfIndex = 0 Then
                strSuffix = ctl.Name.Substring(controlName.Length)
                'Check that the suffix is an integer (index of the array)
                If IsInteger(strSuffix) Then
                    If Val(strSuffix) > maxIndex Then maxIndex = _
                        Val(strSuffix) 'Find the highest indexed Element
                End If
            End If
        Next ctl

        'Add to the list of controls in correct order
        If maxIndex > -1 Then
            For i = 0 To maxIndex
                Dim aControl As Control = _
                    getControlFromName(frm, controlName, i, separator)
                If Not (aControl Is Nothing) Then
                    'Save the object Type (uses the last control found as the Type)
                    controlType = aControl.GetType
                End If
                alist.Add(aControl)
            Next
        End If

        Return alist.ToArray(controlType)
    End Function

    'Converts any type of like named controls on a form to a control array 
    'by using the notation ControlName_1, ControlName_2, where the _ can be 
    'replaced by any separator string
    Public Shared Function getMixedControlArray( _
        ByVal frm As Windows.Forms.Form, _
        ByVal controlName As String, _
        Optional ByVal separator As String = "") _
        As Control()
        Dim i As Short
        Dim startOfIndex As Short
        Dim alist As New ArrayList
        Dim controlType As System.Type

        Dim ctl As System.Windows.Forms.Control
        Dim ctrls() As System.Windows.Forms.Control
        Dim strSuffix As String
        Dim maxIndex As Short = -1 'Default

        'Loop through all controls, looking for controls with _
        'the matching name pattern
        'Find the highest indexed control
        For Each ctl In frm.Controls
            startOfIndex = _
                ctl.Name.ToLower.IndexOf(controlName.ToLower & separator)
            If startOfIndex = 0 Then
                strSuffix = ctl.Name.Substring(controlName.Length)
                'Check that the suffix is an integer (index of the array)
                If IsInteger(strSuffix) Then
                    'Find the highest indexed Element
                    If Val(strSuffix) > maxIndex Then maxIndex = Val(strSuffix)
                End If
            End If
        Next ctl

        'Add to the list of controls in correct order
        If maxIndex > -1 Then
            For i = 0 To maxIndex
                Dim aControl As Control = _
                    getControlFromName(frm, controlName, i, separator)
                alist.Add(aControl)
            Next
        End If

        Return alist.ToArray(GetType(Control))
    End Function

    Private Shared Function getControlFromName( _
        ByRef frm As Windows.Forms.Form, _
        ByVal controlName As String, _
        ByVal index As Short, ByVal separator As String) _
        As System.Windows.Forms.Control

        controlName = controlName & separator & index
        For Each ctl As Control In frm.Controls
            If String.Compare(ctl.Name, controlName, True) = 0 Then
                Return ctl
            End If
        Next ctl

        Return Nothing  'Could not find this control by name
    End Function

    Private Shared Function IsInteger(ByVal Value As String) As Boolean

        If Value = "" Then Return False

        For Each chr As Char In Value
            If Not Char.IsDigit(chr) Then
                Return False
            End If
        Next
        Return True
    End Function

End Class
