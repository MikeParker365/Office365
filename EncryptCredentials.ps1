# Path to the script to be created:
$path = 'c:\Mike\Credentials.txt'

# Create empty template script:
New-Item -ItemType File $path -Force -ErrorAction SilentlyContinue

$pwd = Read-Host 'Enter Password' -AsSecureString
$user = Read-Host 'Enter Username'
$key = 1..32 | ForEach-Object { Get-Random -Maximum 256 }
$pwdencrypted = $pwd | ConvertFrom-SecureString -Key $key

$private:ofs = ' '
('$password = "{0}"' -f $pwdencrypted) | Out-File $path
('$key = "{0}"' -f "$key") | Out-File $path -Append

'$passwordSecure = ConvertTo-SecureString -String $password -Key ([Byte[]]$key.Split(" "))' | 
	Out-File $path -Append
('$cred = New-Object system.Management.Automation.PSCredential("{0}", $passwordSecure)' -f $user) |
	Out-File $path -Append
'$cred' | Out-File $path -Append

ise $path
