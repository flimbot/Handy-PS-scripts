$mcSession = $null
$mcDomain = "https://online.mortgagechoice.com.au"

function Login-MortgageChoice {
    Param(
        [Parameter(Mandatory=$true)][ValidatePattern("^\d{8}$")][String]$username,
        [Parameter(Mandatory=$true)][String]$password
    )

    $headers = @{
        #'referrer'= "$mcDomain/sepas/serve?TAM_OP=login&USERNAME=unauthenticated&ERROR_CODE=0x00000000&URL=%2Fmortgagechoice-online&HOSTNAME=online.mortgagechoice.com.au&PROTOCOL=https";
        'content-type'='application/x-www-form-urlencoded';
        'DNT'='1';
    } 
        $result = Invoke-WebRequest -Method Post -Uri "$mcDomain/pkmslogin.form" -Body @{'username'=$username; 'password'=$password; 'login-form-type'='pwd'} -SessionVariable mcSession -Headers $headers -UseBasicParsing -MaximumRedirection 0 -ErrorAction Ignore

    if($result.StatusCode -eq 302) {
        $p3p = $result.headers['P3P']
        $redirParams = @{}
        $result.headers['location'] | ?{$_ -match '\?(.+)$'} | %{$matches[1] -split '&'} | %{
            $a = @($_ -split '=')
            $redirParams[$a[0]] = $a[1];
        }

        if($redirParams['ERROR_CODE'] -ne '0x00000000') {
            $result.Content -match '\<h3\>(.+)\<\/h3\>' | %{$matches}
            throw $redirParams
        }
        else {
            Invoke-WebRequest -Method Post -Uri "$mcDomain/pkmslogin.form" -Body @{'username'=$username; 'password'=$password; 'login-form-type'='pwd'} -WebSession $mcSession -Headers $headers -UseBasicParsing -MaximumRedirection 0 -ErrorAction Ignore
        }
    }
    
    if( $result.Content -match '\<h[1-3]\>(.+)\<\/h[1-3]\>' ) {
        throw $matches[1]
    }
    else {
        $result
    }

    #$result = Invoke-RestMethod -Method Post -Uri "$mcDomain/pkmslogin.form" -Body @{'username'=$username; 'password'=$password; 'login-form-type'='pwd'} -SessionVariable mcSession -Headers $headers
    #Invoke-WebRequest -Method Post -Uri "$mcDomain/pkmslogin.form" -Body @{'username'=$username; 'password'=$password; 'login-form-type'='pwd'} -SessionVariable mcSession -UseBasicParsing
}
#$test = Login-MortgageChoice -username '89295383' -password '12345678'



function Logout-MortgageChoice {
    Invoke-RestMethod -Method Get -WebSession $mcSession -Uri "$mcDomain/pkmslogout"
}

function Get-MortgageChoice-Accounts {
    $embedSwitches = @('rates','product','position','interest','repayment','borrowers','securities','permissions','related_accounts','related_parties','nominated_accounts','account_name','card_details')
    $productTypeSwitches = @('cash','credit_card','savings','transaction','mortgage','super','term_deposit','investment','super','pension');
    $limit = 100
    $includeBalances = $true
    $apiKey = ''
    (Invoke-RestMethod -Method Get -WebSession $mcSession -Uri "$mcDomain/white/api/channel/account/v3s/accounts-facilities?limit=$limit&includeBalances=$includeBalances&$('embed='+($embedSwitches -join '&embed='))&$('product-type='+($productTypeSwitches -join '&product-type='))&api_key=$apiKey").data.accounts
}

function Get-MortgageChoice-Announcements {
    $params = @{
        'brand'='mortgagechoice'
    }
    (Invoke-RestMethod -Method Get -WebSession $mcSession -Uri "$mcDomain/white/api/channel/home/v1/announcements" -Body $params).data
}

function Get-MortgageChoice-Permissions {
    (Invoke-RestMethod -Method Get -WebSession $mcSession -Uri "$mcDomain/white/api/channel/pay/v4/permissions").data
}

function Get-MortgageChoice-StatementSummary([string]$accountId){
    $params = @{
        'account-id'=$accountId
    }

    #optional:            &billing-cycle=01
    (Invoke-RestMethod -Method Get -WebSession $mcSession -Uri "$mcDomain/white/api/channel/statement/v3s/statement-summary" -Body $params).data
}

function Get-MortgageChoice-TransactionSync([string]$accountId){
    $params = @{
        'account-id'=$accountId
    }
    (Invoke-RestMethod -Method Get -WebSession $mcSession -Uri "$mcDomain/white/api/channel/transaction/v3s/transactions/sync" -Body $params).data
}

function Get-MortgageChoice-Transactions([string]$accountId){
    throw 'not yet implemented'
    #$params = @{
    #    'account-id'=$accountId
    #}
    #(Invoke-RestMethod -Method Get -WebSession $mcSession -Uri "$mcDomain/white/api/channel/transaction/v3s/transactions/sync" -Body $params).data
}

function Get-MortgageChoice-Payees {
    Param(
        [Parameter(Mandatory=$false)]
            [string]$keyword,
        [Parameter(Mandatory=$false)]
            [ValidateSet('frequency','somethignelse')]
            [string]$sort,
        [Switch]$desc
            
    )
    $params = @{}
    if($keyword) {
        $params['q'] = $keyword
    }
    if($sort) {   
        if($desc.IsPresent) {
            $params['sort'] = ('-' + $sort);
        }
        else {
            $params['sort'] = $sort;
        }
    }

    (Invoke-RestMethod -Method Get -WebSession $mcSession -Uri "$mcDomain/white/api/channel/payee/v3s/payees" -Body $params).data
}

# ----------------------------

try {
    Get-Credential -Message 'Please enter online banking login' | %{
        $loginresult = Login-MortgageChoice -username $_.UserName -password ($_.GetNetworkCredential().password)
        $accounts = Get-MortgageChoice-Accounts
        $accounts
    }
}
catch [Exception] {
    Write-Host 'Error:' $_.Exception.Message
    return
}
finally {
    Logout-MortgageChoice
}

