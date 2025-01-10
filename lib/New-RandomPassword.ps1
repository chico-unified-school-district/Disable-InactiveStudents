function New-RandomPassword {
 function Get-RandomCharacters($length, $characters) {
  $random = 1..$length | ForEach-Object { Get-Random -Maximum $characters.length }
  $private:ofs = ''
  return [String]$characters[$random]
 }
 $chars = 'ABCDEFGHKLMNOPRSTUVWXYZabcdefghiklmnoprstuvwxyz1234567890!$#%&*@'
 do { $pw = (Get-RandomCharacters -length 30 -characters $chars) }
 until ($pw -match '[A-Za-z\d!$#%&*@]') # Make sure minimum criteria met using regex p@ttern.
 $pw # Output random password
}