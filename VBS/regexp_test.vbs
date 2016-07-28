'Testing regular Expressions

sTestString = "S:\CAX\Computer Resources\IT Utilities\this, is ,a , test,folder"
wscript.echo fReEx(sTestString)

'Subroutine to check for commas (,) in a string.
Public Function fReEx(Path)
  fReEx = Replace(Path, ",", " ")
End Function