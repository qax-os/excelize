'/*
'Contains information about run-time errors.
'
'
'*/
Public Class Err()

'/*
'Returns or sets a string expression containing a 
'descriptive string associated with an object. Read/write.
'
'*/
Public Property Description As String

'/*
'Returns or sets a numeric value specifying an error. 
'Number is the Err object's default property. Read/write. 
'
'*/
Public Property Number As Integer

    
'/*
'Returns or sets a string expression specifying the name of 
'the object or application that originally generated the error. Read/write. 
'
'*/
Public Property Source As String

End Class