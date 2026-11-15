<#
.SYNOPSIS
    Fin du SMS et de l'appel vocal en MFA (Microsoft Entra ID) : inventaire des méthodes enregistrées et plan de bascule.

.DESCRIPTION
    Microsoft retire l'acheminement SMS et voix qu'il fournit pour la MFA Entra ID :
    01/02/2027 pour tous les utilisateurs, 01/07/2027 pour les administrateurs généraux actifs et les utilisateurs externes.
    Ce script :
    - lit les méthodes MFA enregistrées par chaque utilisateur (rapport userRegistrationDetails) ;
    - compte les utilisateurs par méthode enregistrée et par méthode par défaut ;
    - classe chaque utilisateur : prêt (passkey, FIDO2, Windows Hello…), Authenticator/OTP, SMS/voix par défaut,
      SMS/voix uniquement (critique), aucune MFA ;
    - indique l'échéance applicable et l'état des méthodes SMS et Voix dans la stratégie ;
    - en option (-SignInDays), compte les méthodes réellement utilisées en connexion ;
    - produit un CSV et un rapport HTML autonome (couleurs PulsivIT, logo UnderTenant).

    Prérequis : module Microsoft.Graph.Authentication ; scopes AuditLog.Read.All, Directory.Read.All, Policy.Read.All ;
    licence Microsoft Entra ID P1 ou P2 ; rôle Lecteur de rapports, Lecteur sécurité ou Lecteur général au minimum.
    Les comptes désactivés n'apparaissent pas dans le rapport Microsoft. Données mises à jour sous 36 h environ.

.PARAMETER OutputFolder
    Dossier de sortie du CSV et du rapport HTML (par défaut : dossier courant).

.PARAMETER SignInDays
    Nombre de jours de journaux de connexion à analyser (1 à 30). 0 (par défaut) = pas d'analyse des connexions.

.PARAMETER NoOpen
    N'ouvre pas le rapport HTML à la fin.

.EXAMPLE
    .\Get-MfaSmsVoiceReadiness.ps1 -OutputFolder C:\Temp

.EXAMPLE
    .\Get-MfaSmsVoiceReadiness.ps1 -OutputFolder C:\Temp -SignInDays 7

.NOTES
    Auteur  : Alexis Charrière (PulsivIT)
    Blog    : https://undertenant.fr
    Version : 1.0.0 (15/11/2026)
#>
#Requires -Modules Microsoft.Graph.Authentication
[CmdletBinding()]
param(
    [string]$OutputFolder = (Get-Location).Path,
    [ValidateRange(0, 30)][int]$SignInDays = 0,
    [switch]$NoOpen
)

$ErrorActionPreference = 'Stop'
if (-not (Test-Path $OutputFolder)) { New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null }

Connect-MgGraph -Scopes "AuditLog.Read.All","Directory.Read.All","Policy.Read.All" -NoWelcome

function Get-GraphAll([string]$Uri) {
    $items = New-Object System.Collections.Generic.List[object]
    while ($Uri) {
        $page = Invoke-MgGraphRequest -Method GET -Uri $Uri -OutputType PSObject
        foreach ($v in @($page.value)) { $items.Add($v) }
        $Uri = $page.'@odata.nextLink'
    }
    $items
}

# 1. Méthodes enregistrées par utilisateur (rapport Authentication Methods Activity, Entra ID P1/P2)
Write-Host "Lecture des méthodes MFA enregistrées..."
try {
    $details = @(Get-GraphAll "https://graph.microsoft.com/v1.0/reports/authenticationMethods/userRegistrationDetails")
} catch {
    throw "Lecture impossible du rapport des méthodes d'authentification (licence Entra ID P1/P2 et rôle Lecteur de rapports, Lecteur général ou Lecteur sécurité requis) : $($_.Exception.Message)"
}
Write-Host "$($details.Count) utilisateur(s) analysé(s)."

# 2. Administrateurs généraux actifs (échéance au 01/07/2027), y compris via groupe
$gaIds = @{}
try {
    $membres = Get-GraphAll "https://graph.microsoft.com/v1.0/directoryRoles(roleTemplateId='62e90394-69f5-4237-9190-012177145e10')/members?`$select=id"
    foreach ($m in $membres) {
        if ($m.'@odata.type' -eq '#microsoft.graph.group') {
            foreach ($x in (Get-GraphAll "https://graph.microsoft.com/v1.0/groups/$($m.id)/transitiveMembers/microsoft.graph.user?`$select=id")) { $gaIds[$x.id] = $true }
        } else { $gaIds[$m.id] = $true }
    }
} catch { Write-Warning "Administrateurs généraux non lus : $($_.Exception.Message)" }

# 3. État des méthodes SMS et Voix dans la stratégie des méthodes d'authentification
$etatPolitique = [ordered]@{}
foreach ($m in 'sms','voice') {
    try {
        $etatPolitique[$m] = (Invoke-MgGraphRequest -Method GET -OutputType PSObject `
            -Uri "https://graph.microsoft.com/v1.0/policies/authenticationMethodsPolicy/authenticationMethodConfigurations/$m").state
    } catch { $etatPolitique[$m] = 'inconnu' }
}

# 4. (Optionnel) Méthodes réellement utilisées en connexion sur N jours
$regexTel     = '(?i)sms|text message|phone call|voice'
$usageParUser = @{}
$usageGlobal  = @{}
if ($SignInDays -gt 0) {
    Write-Host "Analyse des connexions sur $SignInDays jour(s) (peut être long sur un gros tenant)..."
    $depuis = (Get-Date).ToUniversalTime().AddDays(-$SignInDays).ToString('yyyy-MM-ddTHH:mm:ssZ')
    $uri = "https://graph.microsoft.com/beta/auditLogs/signIns?`$filter=createdDateTime ge $depuis and status/errorCode eq 0&`$top=999"
    foreach ($s in (Get-GraphAll $uri)) {
        foreach ($d in @($s.authenticationDetails)) {
            $meth = [string]$d.authenticationMethod
            if (-not $d.succeeded -or -not $meth -or $meth -in @('Password', 'Previously satisfied')) { continue }
            $usageGlobal[$meth] = 1 + [int]$usageGlobal[$meth]
            if (-not $usageParUser.ContainsKey($s.userId)) { $usageParUser[$s.userId] = @{} }
            $usageParUser[$s.userId][$meth] = 1 + [int]$usageParUser[$s.userId][$meth]
        }
    }
}

# 5. Classement de chaque utilisateur
$libelles = @{
    mobilePhone = 'Téléphone mobile (SMS / appel)'; alternateMobilePhone = 'Mobile secondaire (appel)'; officePhone = 'Téléphone de bureau (appel)'
    microsoftAuthenticatorPush = 'Authenticator (notification)'; microsoftAuthenticatorPasswordless = 'Authenticator (sans mot de passe)'
    softwareOneTimePasscode = 'Code OTP logiciel'; hardwareOneTimePasscode = 'Jeton OTP matériel'; fido2 = 'Clé de sécurité FIDO2'
    passKeyDeviceBound = 'Passkey (liée à un appareil)'; passKeyDeviceBoundAuthenticator = 'Passkey dans Authenticator'
    passKeyDeviceBoundWindowsHello = 'Passkey Windows'; windowsHelloForBusiness = 'Windows Hello Entreprise'
    macOsSecureEnclaveKey = 'Platform SSO macOS'; email = 'E-mail (SSPR)'; temporaryAccessPass = 'Accès temporaire (TAP)'
    securityQuestion = 'Questions de sécurité'
}
$libDefaut = @{ push = 'Authenticator (notification)'; oath = 'Code OTP'; sms = 'SMS'; voiceMobile = 'Appel mobile'
                voiceAlternateMobile = 'Appel mobile secondaire'; voiceOffice = 'Appel bureau'; none = 'Aucune' }
$telephone = @('mobilePhone', 'alternateMobilePhone', 'officePhone')
$defautTel = @('sms', 'voiceMobile', 'voiceAlternateMobile', 'voiceOffice')
function Test-Resistant([string]$m) { $m -like 'passKey*' -or $m -in @('fido2', 'windowsHelloForBusiness', 'macOsSecureEnclaveKey') }
function Test-AutreFort([string]$m) { $m -like 'microsoftAuthenticator*' -or $m -in @('softwareOneTimePasscode', 'hardwareOneTimePasscode') }

$finFevrier = Get-Date '2027-02-01'
$finJuillet = Get-Date '2027-07-01'
$aujourdhui = Get-Date

$rapport = @(foreach ($d in $details) {
    $m    = @($d.methodsRegistered)
    $aTel = @($m | Where-Object { $_ -in $telephone }).Count -gt 0
    $aRes = @($m | Where-Object { Test-Resistant $_ }).Count -gt 0
    $aFort = @($m | Where-Object { Test-AutreFort $_ }).Count -gt 0
    $defaut = if ($d.isSystemPreferredAuthenticationMethodEnabled -and @($d.systemPreferredAuthenticationMethods).Count) {
                  [string]@($d.systemPreferredAuthenticationMethods)[0] } else { [string]$d.userPreferredMethodForSecondaryAuthentication }
    $ga      = [bool]$gaIds[$d.id]
    $externe = $d.userType -eq 'guest'
    $statut = if ($aRes)                              { 'Prêt (résistant à l''hameçonnage)' }
              elseif ($aTel -and -not $aFort)         { 'Critique : SMS/voix uniquement' }
              elseif ($aTel -and $defaut -in $defautTel) { 'À migrer : SMS/voix par défaut' }
              elseif ($aFort)                         { 'À migrer : Authenticator / OTP' }
              else                                    { 'Aucune méthode MFA' }
    $u = $usageParUser[$d.id]
    [pscustomobject]@{
        Nom              = $d.userDisplayName
        UPN              = $d.userPrincipalName
        Type             = if ($externe) { 'Invité' } else { 'Membre' }
        AdminGeneral     = $ga
        Statut           = $statut
        Echeance         = if ($ga -or $externe) { $finJuillet.ToString('dd/MM/yyyy') } else { $finFevrier.ToString('dd/MM/yyyy') }
        MethodeParDefaut = if ($libDefaut[$defaut]) { $libDefaut[$defaut] } elseif ($defaut) { $defaut } else { 'Non définie' }
        Methodes         = ($m | ForEach-Object { if ($libelles[$_]) { $libelles[$_] } else { $_ } }) -join ', '
        Connexions       = if ($u) { ($u.GetEnumerator() | Sort-Object Value -Descending | ForEach-Object { "$($_.Key) ($($_.Value))" }) -join ', ' } else { '' }
        SmsVoixUtilise   = if ($u) { [bool]@($u.Keys | Where-Object { $_ -match $regexTel }).Count } else { '' }
    }
})

# 6. Compteurs
$compteurMethodes = @{}
foreach ($d in $details) { foreach ($x in @($d.methodsRegistered)) { $compteurMethodes[$x] = 1 + [int]$compteurMethodes[$x] } }
$compteurDefaut = @{}
foreach ($r in $rapport) { $compteurDefaut[$r.MethodeParDefaut] = 1 + [int]$compteurDefaut[$r.MethodeParDefaut] }

Write-Host ""
Write-Host "SMS : $($etatPolitique['sms']) | Voix : $($etatPolitique['voice']) (stratégie des méthodes d'authentification)"
$rapport | Group-Object Statut | Sort-Object Count -Descending |
    Select-Object @{n='Statut';e={$_.Name}}, Count | Format-Table -AutoSize
$compteurMethodes.GetEnumerator() | Sort-Object Value -Descending |
    Select-Object @{n='Méthode enregistrée';e={ if ($libelles[$_.Key]) { $libelles[$_.Key] } else { $_.Key } }}, @{n='Utilisateurs';e={$_.Value}} |
    Format-Table -AutoSize

$base = Join-Path $OutputFolder ("MFA_SMS_Voix_{0:yyyyMMdd}" -f $aujourdhui)
$rapport | Sort-Object Statut, Nom | Export-Csv "$base.csv" -NoTypeInformation -Encoding utf8 -Delimiter ';'

# 7. Rapport HTML (couleurs PulsivIT, logo UnderTenant)
$inv = [Globalization.CultureInfo]::InvariantCulture
function Enc($s) { [System.Net.WebUtility]::HtmlEncode([string]$s) }
function Pct($n, $sur) { if ($sur) { ([math]::Round(100 * $n / $sur, 1)).ToString($inv) } else { '0' } }
function Barres($table, $total, $libs, [string]$regexAlerte) {
    $html = ($table.GetEnumerator() | Sort-Object Value -Descending | ForEach-Object {
        $nom = if ($libs -and $libs[$_.Key]) { $libs[$_.Key] } else { $_.Key }
        $cls = if ($regexAlerte -and ($_.Key -match $regexAlerte -or $nom -match $regexAlerte)) { ' alerte' } else { '' }
        "<div class=""bar-row$cls""><span>$(Enc $nom)</span><div class=""bar""><i style=""width:$(Pct $_.Value $total)%""></i></div><b>$($_.Value)</b></div>"
    }) -join "`n"
    if (-not $html) { $html = '<p class="muted">Aucune donnée.</p>' }
    $html
}

$statuts = [ordered]@{
    'Prêt (résistant à l''hameçonnage)' = 'ok'; 'À migrer : Authenticator / OTP' = 'mid'; 'À migrer : SMS/voix par défaut' = 'warn'
    'Critique : SMS/voix uniquement' = 'crit'; 'Aucune méthode MFA' = 'none'
}
$nb = @{}; foreach ($s in $statuts.Keys) { $nb[$s] = @($rapport | Where-Object Statut -eq $s).Count }
$total   = $rapport.Count
$jours   = [math]::Max(0, [int][math]::Ceiling(($finFevrier - $aujourdhui).TotalDays))
$exposes = $nb['Critique : SMS/voix uniquement'] + $nb['À migrer : SMS/voix par défaut']

try {
    $png  = (Invoke-WebRequest "https://undertenant.fr/images/undertenant-logo.png" -UseBasicParsing -TimeoutSec 10).RawContentStream.ToArray()
    $logo = '<img src="data:image/png;base64,' + [Convert]::ToBase64String($png) + '" alt="UnderTenant">'
} catch { $logo = '<span class="logo-txt">UnderTenant</span>' }

$kpis = @(
    "<div class=""kpi accent""><b>J-$jours</b><span>avant le 01/02/2027</span></div>"
    "<div class=""kpi""><b>$total</b><span>Utilisateurs analysés</span></div>"
    "<div class=""kpi ok""><b>$(Pct $nb['Prêt (résistant à l''hameçonnage)'] $total) %</b><span>Prêts (résistants à l'hameçonnage)</span></div>"
    "<div class=""kpi crit""><b>$($nb['Critique : SMS/voix uniquement'])</b><span>SMS/voix uniquement</span></div>"
    "<div class=""kpi warn""><b>$exposes</b><span>Dépendants du SMS/voix</span></div>"
) -join "`n"
$politique = ($etatPolitique.GetEnumerator() | ForEach-Object {
    $cls = if ($_.Value -eq 'enabled') { 'warn' } elseif ($_.Value -eq 'disabled') { 'ok' } else { 'none' }
    $nom = if ($_.Key -eq 'sms') { 'SMS' } else { 'Appel vocal' }
    "<span class=""badge $cls"">$nom : $(Enc $_.Value)</span>" }) -join ' '
$barre   = ($statuts.Keys | Where-Object { $nb[$_] } | ForEach-Object {
    "<span class=""seg $($statuts[$_])"" style=""width:$(Pct $nb[$_] $total)%"" title=""$(Enc $_) : $($nb[$_])""></span>" }) -join ''
$legende = ($statuts.Keys | ForEach-Object { "<li><i class=""dot $($statuts[$_])""></i>$(Enc $_) <b>$($nb[$_])</b></li>" }) -join "`n"
$blocMethodes = Barres $compteurMethodes $total $libelles '(?i)phone|t[ée]l[ée]phone'
$blocDefaut   = Barres $compteurDefaut $total $null '(?i)^(sms|appel)'
$blocUsage    = if ($SignInDays -gt 0) {
    "<section class=""full""><h2>Méthodes utilisées en connexion ($SignInDays derniers jours)</h2>`n" +
    (Barres $usageGlobal (($usageGlobal.Values | Measure-Object -Sum).Sum) $null $regexTel) + "`n</section>" } else { '' }
$lignes = ($rapport | Sort-Object Statut, Nom | ForEach-Object {
    $ga = if ($_.AdminGeneral) { ' <span class="tag">Admin général</span>' } else { '' }
    "<tr data-s=""$($statuts[$_.Statut])""><td>$(Enc $_.Nom)$ga</td><td>$(Enc $_.UPN)</td><td>$(Enc $_.Type)</td><td><span class=""badge $($statuts[$_.Statut])"">$(Enc $_.Statut)</span></td><td>$(Enc $_.MethodeParDefaut)</td><td>$(Enc $_.Methodes)</td><td>$(Enc $_.Echeance)</td></tr>" }) -join "`n"

$modele = @'
<!DOCTYPE html>
<html lang="fr" dir="ltr">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="color-scheme" content="light dark">
<title>Fin du SMS et de la voix en MFA · %%TENANT%%</title>
<style>
:root{--navy:#253471;--ciel:#009DDA;--gris:#678597;--fond:#f4f6fa;--blanc:#fff;--texte:#1c2440;--ligne:#dde3ee;
--ok:#1f9d55;--mid:#009DDA;--warn:#e0a100;--crit:#d64545;--none:#9aa5b1;--titre:#253471;--zebra:#f8f9fc}
@media (prefers-color-scheme:dark){:root{--fond:#0f1729;--blanc:#16213b;--texte:#e7ecf5;--ligne:#2a3a5c;--titre:#7cc8f0;--zebra:#1a2645}}
*{box-sizing:border-box}
body{margin:0;background:var(--fond);color:var(--texte);font:15px/1.55 Montserrat,"Segoe UI",system-ui,sans-serif}
header{background:var(--navy);color:#fff;padding:22px 32px;display:flex;align-items:center;gap:24px;flex-wrap:wrap;border-block-end:4px solid var(--ciel)}
header img{height:52px;width:auto}
.logo-txt{font-weight:800;font-size:1.5rem}
header h1{margin:0;font-size:1.35rem;font-weight:800}
header p{margin:2px 0 0;opacity:.85;font-size:.9rem}
main{max-width:1180px;margin:0 auto;padding:26px 24px 10px}
.policy{margin:0 0 18px;font-size:.9rem;color:var(--gris)}
.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:14px;margin-block-end:22px}
.kpi{background:var(--blanc);border:1px solid var(--ligne);border-radius:12px;padding:16px 18px;border-block-start:4px solid var(--navy)}
.kpi b{display:block;font-size:1.9rem;font-weight:800;color:var(--titre);line-height:1.2}
.kpi span{color:var(--gris);font-size:.85rem}
.kpi.ok{border-block-start-color:var(--ok)}.kpi.ok b{color:var(--ok)}
.kpi.crit{border-block-start-color:var(--crit)}.kpi.crit b{color:var(--crit)}
.kpi.warn{border-block-start-color:var(--warn)}.kpi.warn b{color:var(--warn)}
.kpi.accent{border-block-start-color:var(--ciel)}.kpi.accent b{color:var(--ciel)}
.grid2{display:grid;grid-template-columns:1fr 1fr;gap:18px;margin-block-end:22px}
section{background:var(--blanc);border:1px solid var(--ligne);border-radius:12px;padding:18px 20px}
h2{margin:0 0 14px;font-size:1.05rem;color:var(--titre);font-weight:800}
.stack{display:flex;height:22px;border-radius:11px;overflow:hidden;background:var(--ligne)}
.ok.seg,.ok.dot,.ok.badge{background:var(--ok)}.mid.seg,.mid.dot,.mid.badge{background:var(--mid)}
.warn.seg,.warn.dot,.warn.badge{background:var(--warn)}.crit.seg,.crit.dot,.crit.badge{background:var(--crit)}
.none.seg,.none.dot,.none.badge{background:var(--none)}
.legend{list-style:none;padding:0;margin:14px 0 0;display:grid;grid-template-columns:1fr;gap:6px;font-size:.9rem}
.dot{display:inline-block;width:11px;height:11px;border-radius:50%;margin-inline-end:8px}
.legend b{float:inline-end}
.bar-row{display:grid;grid-template-columns:210px 1fr 44px;align-items:center;gap:10px;margin:7px 0;font-size:.9rem}
.bar{height:12px;background:var(--fond);border-radius:6px;overflow:hidden}
.bar i{display:block;height:100%;background:linear-gradient(90deg,var(--navy),var(--ciel))}
.bar-row.alerte i{background:linear-gradient(90deg,var(--warn),var(--crit))}
.bar-row b{text-align:end;color:var(--titre)}
table{width:100%;border-collapse:collapse;font-size:.86rem}
th{background:var(--navy);color:#fff;text-align:start;padding:9px 10px;font-weight:600}
td{padding:8px 10px;border-block-end:1px solid var(--ligne);vertical-align:top}
tbody tr:nth-child(even){background:var(--zebra)}
.badge{color:#fff;border-radius:20px;padding:2px 10px;font-size:.78rem;white-space:nowrap;display:inline-block}
.tag{border:1px solid var(--crit);color:var(--crit);border-radius:6px;padding:0 6px;font-size:.72rem;white-space:nowrap}
.tools{display:flex;gap:8px;flex-wrap:wrap;align-items:center;margin-block-end:12px}
.tools input{flex:1;min-width:220px;padding:8px 12px;border:1px solid var(--ligne);border-radius:8px;font:inherit;background:var(--blanc);color:var(--texte)}
.tools button{border:1px solid var(--titre);background:var(--blanc);color:var(--titre);border-radius:8px;padding:7px 12px;font:inherit;font-size:.85rem;cursor:pointer}
.tools button.on{background:var(--navy);color:#fff;border-color:var(--navy)}
.muted{color:var(--gris)}
.full{margin-block-end:22px}
footer{text-align:center;color:var(--gris);font-size:.8rem;padding:10px 20px 26px}
footer a{color:var(--ciel)}
@media (max-width:820px){.grid2{grid-template-columns:1fr}.bar-row{grid-template-columns:140px 1fr 40px}}
@media print{:root{--fond:#fff;--blanc:#fff;--texte:#1c2440;--ligne:#dde3ee;--titre:#253471;--zebra:#f8f9fc}.tools{display:none}section,.kpi{break-inside:avoid}}
</style>
</head>
<body>
<header>
  %%LOGO%%
  <div>
    <h1>Fin du SMS et de l'appel vocal en MFA</h1>
    <p>%%TENANT%% · généré le %%DATE%% · retrait Microsoft : 01/02/2027 (01/07/2027 pour les administrateurs généraux et les externes)</p>
  </div>
</header>
<main>
  <p class="policy">Stratégie des méthodes d'authentification : %%POLITIQUE%%</p>
  <div class="kpis">
%%KPIS%%
  </div>
  <div class="grid2">
    <section>
      <h2>Préparation des utilisateurs</h2>
      <div class="stack">%%BARRE%%</div>
      <ul class="legend">
%%LEGENDE%%
      </ul>
    </section>
    <section>
      <h2>Méthode MFA par défaut</h2>
%%DEFAUT%%
    </section>
  </div>
  <section class="full">
    <h2>Méthodes enregistrées (nombre d'utilisateurs)</h2>
%%METHODES%%
  </section>
%%USAGE%%
  <section class="full">
    <h2>Utilisateurs (%%NB%%)</h2>
    <div class="tools">
      <input id="q" type="search" placeholder="Rechercher un nom, une méthode…" aria-label="Rechercher">
      <button type="button" class="on" data-f="">Tous</button>
      <button type="button" data-f="crit">SMS/voix uniquement</button>
      <button type="button" data-f="warn">SMS/voix par défaut</button>
      <button type="button" data-f="mid">Authenticator / OTP</button>
      <button type="button" data-f="none">Sans MFA</button>
      <button type="button" data-f="ok">Prêts</button>
    </div>
    <table id="users"><thead><tr><th>Nom</th><th>UPN</th><th>Type</th><th>Statut</th><th>Méthode par défaut</th><th>Méthodes enregistrées</th><th>Échéance</th></tr></thead>
    <tbody>
%%USERS%%
    </tbody></table>
  </section>
</main>
<footer>Rapport généré avec le script de l'article « Fin du SMS et de l'appel vocal en MFA » · <a href="https://undertenant.fr">undertenant.fr</a></footer>
<script>
(function(){
  var q=document.getElementById('q'),f='',rows=document.querySelectorAll('#users tbody tr'),btns=document.querySelectorAll('.tools button');
  function apply(){var t=q.value.toLowerCase();rows.forEach(function(r){r.hidden=!((!f||r.dataset.s===f)&&r.textContent.toLowerCase().indexOf(t)>-1);});}
  q.addEventListener('input',apply);
  btns.forEach(function(b){b.addEventListener('click',function(){btns.forEach(function(x){x.classList.remove('on');});b.classList.add('on');f=b.dataset.f;apply();});});
})();
</script>
</body>
</html>
'@

$valeurs = [ordered]@{
    '%%LOGO%%'      = $logo
    '%%TENANT%%'    = Enc (Get-MgContext).TenantId
    '%%DATE%%'      = $aujourdhui.ToString('dd/MM/yyyy HH:mm')
    '%%POLITIQUE%%' = $politique
    '%%KPIS%%'      = $kpis;  '%%BARRE%%' = $barre;  '%%LEGENDE%%' = $legende
    '%%DEFAUT%%'    = $blocDefaut;  '%%METHODES%%' = $blocMethodes;  '%%USAGE%%' = $blocUsage
    '%%USERS%%'     = $lignes;  '%%NB%%' = $total
}
try { $valeurs['%%TENANT%%'] = Enc (Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/organization?`$select=displayName" -OutputType PSObject).value[0].displayName } catch { }
$html = $modele
foreach ($k in $valeurs.Keys) { $html = $html.Replace($k, [string]$valeurs[$k]) }
$html | Set-Content "$base.html" -Encoding utf8
Write-Host "Rapports : $base.csv et $base.html"
if (-not $NoOpen) { Invoke-Item "$base.html" }
