Attribute VB_Name = "workspace"
Option Explicit

Sub dev()

Console.log "Hello!"

Call unittest

End Sub

Private Sub unittest()
'''' *************************************************
Console.info "-------------------- start !!"
Package.Include
''

''
Package.Terminate
Console.info "-------------------- end ...."
''
End Sub

