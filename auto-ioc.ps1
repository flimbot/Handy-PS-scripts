# AUTO-IOC

$textblock = """
ADD TEXT BLOCK FROM EMAIL OR DOCUMENT HERE
"""

#---------------------------------------------------------------------------------------------------------

# Note: The first argument in a replacement is a regular expression
$defangreplaces = @(
    @('hxxp','http'),
    @('fxp','ftp'),
    @('\[\:\]',':'),
    @('\[\.\]','.'),
    @('\[@\]','@')
)

$ioctypes = @{
    'domainsfiles' = @{
        'title'    = 'Domains and files'
        'identify' = "[\d\w\-]+\[?\.[\[\]\d\w\.\-]+"
    };
    'http' = @{
        'title'    = 'HTTP, HTTPS, FTP, and FTPS'
        'identify' = "(f[tx]ps?|h(tt|xx)ps?)\[?:\]?//[\d\w\-]+\[?\.[\[\]\d\w\.\-]+"
    };
    'hashsha1' = @{
        'title'    = 'Hash - SHA1'
        'identify' = "([a-fA-F0-9]{40})"
    };
    'hashmd5' = @{
        'title'    = 'Hash - MD5'
        'identify' = "([a-fA-F0-9]{32})"
    };
    'email' = @{
        'title'    = 'Email'
        'identify' = "[\w\d\-]+\[?@\]?[\d\w\-]+\[?\.[\[\]\d\w\.\-]+"
    };
}

Write-Host "Identifying potential IOCs in text..."
$timestamp = (Get-Date).ToFileTime()

$ioctypes.Keys | Sort-Object | %{
    $type = $ioctypes[$_].title
    Write-Host $type

    $iocsRaw = $textblock | Select-String -Pattern $ioctypes[$_].identify -AllMatches | %{$_.matches.value} | Select -Unique
    $iocsDefanged = $iocsRaw | %{
        $ioc = $_
        $defangreplaces | %{
            $ioc = $ioc -replace $_[0],$_[1]
        }
        $ioc
    }
    
    $iocsDefanged | %{ 
        Write-Host "`t$_" 
        $_ | Select @{Name='Type';Expression={$type}},@{Name='Value';Expression={$_}}
    }
} | Export-csv "IOC-$timestamp.csv"