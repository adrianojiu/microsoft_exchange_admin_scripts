# Convert return from legacydn cache error to X500 address.

$IMCEAEX= Read-Host "Please enter the IMCEAEX that needs to be converted"
$Clean = $IMCEAEX.replace("+20", " ").replace("+28", "(").Replace("+29", ")").replace("IMCEAEX-", "X500:").replace("_", "/").replace("+2E", ".").replace("+2C", ",").split("@")[0]

Write-Host "The converted X500 is: $Clean" -ForegroundColor Yellow
