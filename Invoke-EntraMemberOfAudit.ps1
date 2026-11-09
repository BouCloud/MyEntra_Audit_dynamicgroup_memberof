#Requires -Version 5.1
<#
.SYNOPSIS
    Recense les règles dynamiques Entra ID qui utilisent memberOf et tout ce qui en dépend dans le tenant.

.DESCRIPTION
    Microsoft arrête le traitement de l'opérateur memberOf le 3 novembre 2026 (MC1448379).
    Les groupes dynamiques, AU dynamiques et stratégies d'affectation automatique EM qui
    l'utilisent ne seront pas supprimés : leurs membres resteront figés. Tout ce qui s'appuie
    dessus (licences, accès conditionnel, Intune...) ne suivra plus les mouvements de personnel.

    Le script fait l'inventaire, cherche les dépendances et sort un classeur Excel ainsi
    qu'une page HTML : une fiche par groupe (où il est utilisé) et une carte des dépendances.
    Il ne fait que de la lecture (requêtes GET) : aucune modification du tenant.

    Au lancement, il affiche ce qu'il va faire, contrôle les prérequis (version PowerShell,
    modules, rôle attendu) et demande confirmation. Utiliser -Unattended pour une tâche planifiée.

.PARAMETER TenantId
    Tenant cible (GUID ou domaine .onmicrosoft.com). Obligatoire en mode application.

.PARAMETER ClientId
    App ID de l'inscription d'application (mode app-only).

.PARAMETER CertificateThumbprint
    Empreinte du certificat de l'application (magasin CurrentUser\My ou LocalMachine\My).

.PARAMETER UseDeviceCode
    Connexion déléguée par code d'appareil, pour un serveur sans navigateur.

.PARAMETER OutputPath
    Fichier Excel produit. Par défaut : .\Audit-MemberOf-AAAAMMJJ-HHmm.xlsx

.PARAMETER LogPath
    Journal d'exécution (transcript). Par défaut : même nom que le fichier Excel, en .log

.PARAMETER MapPath
    Page HTML (fiches par objet + carte des dépendances). Par défaut : même nom que le fichier Excel, en .html

.PARAMETER NoMap
    Ne génère pas la page HTML.

.PARAMETER IncludeSourceGroupDependencies
    Cherche aussi les dépendances des groupes sources cités dans les règles.

.PARAMETER SkipIntune
    Ne collecte pas les affectations Intune (plus rapide, ou tenant sans Intune).

.PARAMETER InstallMissingModules
    Installe les modules manquants sans poser la question (portée CurrentUser).

.PARAMETER Unattended
    Pas d'écran de confirmation. À utiliser en tâche planifiée.

.EXAMPLE
    .\Invoke-EntraMemberOfAudit.ps1

    Connexion interactive avec un compte Lecteur général.

.EXAMPLE
    .\Invoke-EntraMemberOfAudit.ps1 -TenantId contoso.onmicrosoft.com -UseDeviceCode -SkipIntune

.EXAMPLE
    .\Invoke-EntraMemberOfAudit.ps1 -TenantId contoso.onmicrosoft.com -ClientId <AppId> `
        -CertificateThumbprint <Empreinte> -OutputPath D:\Audits\memberof.xlsx -Unattended

    Exécution planifiée en mode application.

.NOTES
    Auteur      : Chris Bousquet - https://github.com/BouCloud
    Compatible  : Windows PowerShell 5.1 et PowerShell 7.x
    Référence   : Microsoft 365 Message Center MC1448379

    Historique
    1.0  12/10/2026  Première version : inventaire, dépendances, export Excel.
    1.1  13/10/2026  Écran de présentation, contrôle des prérequis, confirmation,
                     journal d'exécution, vérification du rôle du compte connecté.
    1.2  13/10/2026  Teams, SharePoint et boîte du groupe détaillés (avec URL), colonne Famille,
                     onglet « Vue par objet », page HTML avec fiches et carte des dépendances.
    1.2.1 04/11/2026 Ajout de DeviceManagementScripts.Read.All : sans elle, les scripts Intune
                     et les remédiations renvoyaient Forbidden.
#>
[CmdletBinding(DefaultParameterSetName = 'Interactive')]
param(
    [Parameter(ParameterSetName = 'Interactive')]
    [Parameter(ParameterSetName = 'AppCert', Mandatory)]
    [string]$TenantId,

    [Parameter(ParameterSetName = 'AppCert', Mandatory)]
    [string]$ClientId,

    [Parameter(ParameterSetName = 'AppCert', Mandatory)]
    [string]$CertificateThumbprint,

    [Parameter(ParameterSetName = 'Interactive')]
    [switch]$UseDeviceCode,

    [string]$OutputPath = (Join-Path (Get-Location) ('Audit-MemberOf-{0:yyyyMMdd-HHmm}.xlsx' -f (Get-Date))),

    [string]$LogPath,

    [string]$MapPath,

    [switch]$NoMap,
    [switch]$IncludeSourceGroupDependencies,
    [switch]$SkipIntune,
    [switch]$InstallMissingModules,
    [switch]$Unattended
)

$ErrorActionPreference = 'Stop'

$ScriptVersion  = '1.2.1'
$RetirementDate = [datetime]'2026-11-03'
$GraphV1        = 'https://graph.microsoft.com/v1.0'
$GraphBeta      = 'https://graph.microsoft.com/beta'
$MemberOfRegex  = '(?i)\bmemberof\b'
$GuidRegex      = '[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}'

if (-not $LogPath) { $LogPath = [System.IO.Path]::ChangeExtension($OutputPath, '.log') }
if (-not $MapPath) { $MapPath = [System.IO.Path]::ChangeExtension($OutputPath, '.html') }

# Regroupement des catégories en familles lisibles (colonnes de la vue par objet, couleurs de la carte)
$FamilyOrder = @(
    'Licences', 'Accès conditionnel', 'Rôles', 'MFA / authentification', 'Intune', 'Applications',
    'Teams', 'SharePoint', 'Exchange', 'Règles chaînées', 'Entitlement Management', 'B2B',
    'Imbrication', 'Revues d''accès', 'Autre'
)

# Modules nécessaires. On n'utilise volontairement que Microsoft.Graph.Authentication
# (Invoke-MgGraphRequest) plutôt que le SDK complet : une seule dépendance, pas de conflit
# de versions entre sous-modules Microsoft.Graph.*.
$RequiredModules = @(
    [pscustomobject]@{ Name = 'Microsoft.Graph.Authentication'; MinVersion = '2.0.0'; Usage = 'Connexion et requêtes Microsoft Graph' }
    [pscustomobject]@{ Name = 'ImportExcel';                    MinVersion = '7.8.0'; Usage = 'Écriture du fichier Excel (Excel non requis sur le poste)' }
)

# Permissions Graph, toutes en lecture. Utilisées en délégué (scopes demandés à la connexion)
# et listées à l'écran pour le mode application, où elles se configurent sur l'inscription d'app.
$ReadScopes = [ordered]@{
    'Directory.Read.All'                     = 'Objets d''annuaire'
    'Group.Read.All'                         = 'Groupes, règles, affectations d''applications'
    'AdministrativeUnit.Read.All'            = 'Unités administratives'
    'Policy.Read.All'                        = 'Accès conditionnel, méthodes d''authentification, B2B'
    'Application.Read.All'                   = 'Applications d''entreprise'
    'RoleManagement.Read.Directory'          = 'Rôles Entra actifs et éligibles (PIM)'
    'EntitlementManagement.Read.All'         = 'Packages d''accès'
    'AccessReview.Read.All'                  = 'Revues d''accès'
    'Organization.Read.All'                  = 'Nom du tenant, références de licences'
    'Team.ReadBasic.All'                     = 'Équipes Teams (nom, lien)'
    'Sites.Read.All'                         = 'Site SharePoint des groupes M365 (lien)'
    'DeviceManagementConfiguration.Read.All' = 'Intune : profils, conformité, sécurité'
    'DeviceManagementApps.Read.All'          = 'Intune : applications, MAM'
    'DeviceManagementServiceConfig.Read.All' = 'Intune : Autopilot, inscription, mises à jour'
    'DeviceManagementRBAC.Read.All'          = 'Intune : rôles'
    'DeviceManagementScripts.Read.All'       = 'Intune : scripts PowerShell, shell macOS, remédiations'
}

# Rôles Entra jugés suffisants pour un audit en délégué (identifiants de modèle, fixes pour tous les tenants)
$SufficientRoles = @{
    'f2ef992c-3afb-46b9-b7cf-a126ee74c451' = 'Lecteur général'
    '62e90394-69f5-4237-9190-012177145e10' = 'Administrateur général'
}

$script:Errors       = [System.Collections.Generic.List[object]]::new()
$script:Dependencies = [System.Collections.Generic.List[object]]::new()
$script:TargetGroups = @{}   # id en minuscules -> nom affiché
$script:TargetAUs    = @{}

#region Affichage et interaction

function Write-Step {
    param([string]$Message)
    Write-Host ('[{0:HH:mm:ss}] {1}' -f (Get-Date), $Message) -ForegroundColor Cyan
}

function Write-Section {
    param([string]$Title)
    Write-Host ''
    Write-Host "  $Title" -ForegroundColor White
    Write-Host ('  ' + ('-' * $Title.Length)) -ForegroundColor DarkGray
}

function Write-Check {
    # Ligne de contrôle alignée : libellé ......... [ OK ] détail
    param([string]$Label, [ValidateSet('OK', 'KO', 'INFO', 'ATTENTION')][string]$State, [string]$Detail = '')
    $color = @{ OK = 'Green'; KO = 'Red'; INFO = 'Gray'; ATTENTION = 'Yellow' }[$State]
    Write-Host ('  {0,-42}' -f $Label) -NoNewline
    Write-Host ('[{0}]' -f $State) -ForegroundColor $color -NoNewline
    Write-Host ((' ' * (11 - $State.Length)) + $Detail)
}

function Show-Banner {
    $rule = '=' * 78
    $daysLeft = [math]::Ceiling(($RetirementDate - (Get-Date)).TotalDays)
    Write-Host ''
    Write-Host $rule -ForegroundColor DarkCyan
    Write-Host '  AUDIT DES RÈGLES DYNAMIQUES memberOf - MICROSOFT ENTRA ID' -ForegroundColor White
    Write-Host ("  Version {0}   |   Réf. MC1448379   |   Échéance 03/11/2026 (J-{1})" -f $ScriptVersion, [math]::Max(0, $daysLeft)) -ForegroundColor Gray
    Write-Host $rule -ForegroundColor DarkCyan

    Write-Section 'Pourquoi cet audit'
    Write-Host '  À partir du 3 novembre 2026, Entra ID ne recalcule plus les règles dynamiques'
    Write-Host '  qui utilisent memberOf. Les groupes et AU concernés restent en place mais leurs'
    Write-Host '  membres sont figés : licences, accès conditionnel ou affectations Intune qui'
    Write-Host '  s''appuient dessus ne suivent plus les arrivées, départs et mutations.'

    Write-Section 'Ce que fait le script'
    Write-Host '  1. Liste les groupes dynamiques, AU dynamiques et stratégies d''affectation'
    Write-Host '     automatique (Entitlement Management) dont la règle contient memberOf.'
    Write-Host '  2. Identifie les groupes sources cités dans ces règles.'
    Write-Host '  3. Cherche où ces objets sont utilisés : licences, accès conditionnel,'
    Write-Host '     applications, rôles, Intune, packages d''accès, revues, MFA, B2B...'
    Write-Host '  4. Produit un classeur Excel (piste de remplacement par objet) et une page'
    Write-Host '     HTML : une fiche par groupe et une carte des dépendances.'

    Write-Section 'Ce qu''il ne fait pas'
    Write-Host '  Aucune modification du tenant : uniquement des requêtes GET (lecture).'
    Write-Host '  Ne couvre pas SharePoint (droits directs), Teams (stratégies), Purview,'
    Write-Host '  Exchange ni Azure RBAC, hors de portée de Microsoft Graph pour ce besoin.'
}

function Get-ModuleStatus {
    foreach ($m in $RequiredModules) {
        $installed = Get-Module -ListAvailable -Name $m.Name | Sort-Object Version -Descending | Select-Object -First 1
        $state = 'OK'
        if (-not $installed) { $state = 'Absent' }
        elseif ($installed.Version -lt [version]$m.MinVersion) { $state = 'Trop ancien' }
        [pscustomobject]@{
            Name       = $m.Name
            MinVersion = $m.MinVersion
            Installed  = $(if ($installed) { $installed.Version.ToString() } else { '' })
            State      = $state
            Usage      = $m.Usage
        }
    }
}

function Show-Prerequisites {
    param([object[]]$ModuleStatus)

    Write-Section 'Environnement'
    $psOk = $PSVersionTable.PSVersion -ge [version]'5.1'
    Write-Check -Label 'PowerShell' -State $(if ($psOk) { 'OK' } else { 'KO' }) `
        -Detail ("{0} ({1})" -f $PSVersionTable.PSVersion, $PSVersionTable.PSEdition)
    Write-Check -Label 'Fichier Excel produit' -State INFO -Detail $OutputPath
    Write-Check -Label 'Carte HTML' -State INFO -Detail $(if ($NoMap) { 'non générée (-NoMap)' } else { $MapPath })
    Write-Check -Label 'Journal d''exécution' -State INFO -Detail $LogPath

    Write-Section 'Modules PowerShell'
    foreach ($m in $ModuleStatus) {
        switch ($m.State) {
            'OK'          { Write-Check -Label $m.Name -State OK -Detail ("v{0} - {1}" -f $m.Installed, $m.Usage) }
            'Absent'      { Write-Check -Label $m.Name -State KO -Detail ("non installé (v{0} minimum requise)" -f $m.MinVersion) }
            'Trop ancien' { Write-Check -Label $m.Name -State KO -Detail ("v{0} installée, v{1} minimum requise" -f $m.Installed, $m.MinVersion) }
        }
    }

    Write-Section 'Droits nécessaires (lecture seule)'
    if ($PSCmdlet.ParameterSetName -eq 'AppCert') {
        Write-Check -Label 'Mode de connexion' -State INFO -Detail 'Application (certificat)'
        Write-Check -Label 'Inscription d''application' -State INFO -Detail $ClientId
        Write-Host '  Permissions Graph de type Application, avec consentement administrateur :' -ForegroundColor Gray
    }
    else {
        Write-Check -Label 'Mode de connexion' -State INFO -Detail $(if ($UseDeviceCode) { 'Délégué (code d''appareil)' } else { 'Délégué (interactif)' })
        Write-Check -Label 'Rôle Entra conseillé' -State INFO -Detail 'Lecteur général (Global Reader)'
        Write-Host '  Permissions déléguées demandées (consentement admin au premier lancement) :' -ForegroundColor Gray
    }
    foreach ($s in $ReadScopes.Keys) {
        Write-Host ('    {0,-40} {1}' -f $s, $ReadScopes[$s]) -ForegroundColor DarkGray
    }

    Write-Section 'Options retenues'
    Write-Check -Label 'Collecte Intune' -State INFO -Detail $(if ($SkipIntune) { 'non (-SkipIntune)' } else { 'oui' })
    Write-Check -Label 'Dépendances des groupes sources' -State INFO -Detail $(if ($IncludeSourceGroupDependencies) { 'oui' } else { 'non' })
    Write-Host ''
}

function Test-Interactive {
    # Pas de question possible en tâche planifiée, en -NonInteractive ou si l'appelant l'a demandé
    if ($Unattended) { return $false }
    if (-not [Environment]::UserInteractive) { return $false }
    if ([Environment]::GetCommandLineArgs() -match '^-noni') { return $false }
    return $true
}

function Read-YesNo {
    param([string]$Question, [bool]$Default = $true)
    $hint = if ($Default) { '[O/n]' } else { '[o/N]' }
    while ($true) {
        $answer = "$(Read-Host "  $Question $hint")".Trim().ToLower()
        if ($answer -eq '') { return $Default }
        if ($answer -in 'o', 'oui', 'y', 'yes') { return $true }
        if ($answer -in 'n', 'non', 'no')      { return $false }
    }
}

function Install-RequiredModules {
    param([object[]]$Missing)

    # Windows PowerShell 5.1 négocie encore parfois TLS 1.0 par défaut : la PowerShell Gallery le refuse.
    if ($PSVersionTable.PSEdition -eq 'Desktop') {
        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
        if (-not (Get-PackageProvider -ListAvailable -Name NuGet -ErrorAction SilentlyContinue)) {
            Write-Step 'Installation du fournisseur NuGet (portée utilisateur)'
            Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Scope CurrentUser -Force | Out-Null
        }
    }
    foreach ($m in $Missing) {
        Write-Step "Installation de $($m.Name) (v$($m.MinVersion) minimum, portée utilisateur)"
        Install-Module -Name $m.Name -MinimumVersion $m.MinVersion -Scope CurrentUser -Repository PSGallery -Force -AllowClobber
    }
}

function Export-Sheet {
    param([object[]]$Data, [string]$Name, [object[]]$ConditionalText)
    $rows = @($Data | Where-Object { $null -ne $_ })
    if ($rows.Count -eq 0) { $rows = @([pscustomobject]@{ Information = 'Aucun élément trouvé' }) }
    $p = @{
        Path            = $OutputPath
        WorksheetName   = $Name
        AutoSize        = $true
        MaxAutoSizeRows = 500
        FreezeTopRow    = $true
        BoldTopRow      = $true
        TableName       = 'T_' + ($Name -replace '[^A-Za-z0-9]', '_')
        TableStyle      = 'Medium2'
    }
    if ($ConditionalText) { $p.ConditionalText = $ConditionalText }
    $rows | Export-Excel @p
}

#endregion

#region Accès Microsoft Graph (GET uniquement)

function Add-AuditError {
    param([string]$Source, [string]$Message)
    $script:Errors.Add([pscustomobject]@{
        Horodatage = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        Source     = $Source
        Erreur     = $Message
    })
    Write-Warning "$Source : $Message"
}

function Invoke-GraphGetAll {
    # Suit la pagination (@odata.nextLink) et renvoie toujours un tableau.
    # Par défaut une erreur est consignée et la collecte continue : un droit manquant
    # sur Intune ne doit pas empêcher d'avoir le reste de l'audit.
    param(
        [Parameter(Mandatory)][string]$Uri,
        [hashtable]$Headers = @{},
        [string]$Source = $Uri,
        [switch]$Throw
    )
    $results = [System.Collections.Generic.List[object]]::new()
    $next = $Uri
    try {
        while ($next) {
            $resp = Invoke-MgGraphRequest -Method GET -Uri $next -Headers $Headers -OutputType PSObject
            # On teste la présence de la propriété 'value' et non son contenu :
            # une collection vide renvoie value = @(), qui serait évalué à faux.
            if ($resp.PSObject.Properties.Name -contains 'value') {
                foreach ($v in @($resp.value)) { if ($null -ne $v) { $results.Add($v) } }
                $next = $resp.'@odata.nextLink'
            }
            else {
                $results.Add($resp)
                $next = $null
            }
        }
    }
    catch {
        if ($Throw) { throw }
        Add-AuditError -Source $Source -Message $_.Exception.Message
    }
    return , $results.ToArray()
}

function Invoke-GraphGetOne {
    param([Parameter(Mandatory)][string]$Uri, [string]$Source = $Uri, [switch]$Quiet)
    try {
        return Invoke-MgGraphRequest -Method GET -Uri $Uri -OutputType PSObject
    }
    catch {
        if (-not $Quiet) { Add-AuditError -Source $Source -Message $_.Exception.Message }
        return $null
    }
}

function Get-MemberCount {
    # $count exige l'en-tête ConsistencyLevel: eventual
    param([Parameter(Mandatory)][string]$RelativePath)
    try {
        $r = Invoke-MgGraphRequest -Method GET -Uri "$GraphV1/$RelativePath/members/`$count" `
                -Headers @{ ConsistencyLevel = 'eventual' } -OutputType Json
        return [int]("$r".Trim())
    }
    catch { return $null }
}

#endregion

#region Analyse

function Get-RuleGuids {
    # memberOf référence les groupes par objectId : on récupère tous les GUID de la règle
    param([string]$Rule)
    if ([string]::IsNullOrWhiteSpace($Rule)) { return @() }
    return @([regex]::Matches($Rule, $GuidRegex) | ForEach-Object { $_.Value.ToLower() } | Select-Object -Unique)
}

function Get-ObjName {
    param($Obj)
    foreach ($p in 'displayName', 'name', 'id') {
        if ($Obj.PSObject.Properties.Name -contains $p -and $Obj.$p) { return [string]$Obj.$p }
    }
    return ''
}

function Get-TargetName {
    param([string]$Id)
    if ($script:TargetGroups.ContainsKey($Id)) { return $script:TargetGroups[$Id] }
    if ($script:TargetAUs.ContainsKey($Id))    { return $script:TargetAUs[$Id] }
    return $Id
}

function Get-TargetType {
    param([string]$Id)
    if ($script:TargetAUs.ContainsKey($Id)) { return 'Unité administrative' }
    return 'Groupe'
}

function Get-DepFamily {
    param([string]$Categorie)
    switch -Wildcard ($Categorie) {
        'Licences*'           { 'Licences'; break }
        'Accès conditionnel'  { 'Accès conditionnel'; break }
        'Rôle*'               { 'Rôles'; break }
        'Méthodes*'           { 'MFA / authentification'; break }
        'Intune*'             { 'Intune'; break }
        'Application*'        { 'Applications'; break }
        'Microsoft Teams'     { 'Teams'; break }
        'SharePoint'          { 'SharePoint'; break }
        'Exchange*'           { 'Exchange'; break }
        'Règle dynamique*'    { 'Règles chaînées'; break }
        'Entitlement*'        { 'Entitlement Management'; break }
        'Accès accordé*'      { 'Entitlement Management'; break }
        'Accès inter*'        { 'B2B'; break }
        'Membre*'             { 'Imbrication'; break }
        'Revue*'              { 'Revues d''accès'; break }
        default               { 'Autre' }
    }
}

function Add-Dependency {
    param(
        [string]$ObjetId, [string]$ObjetNom, [string]$ObjetType,
        [string]$Categorie, [string]$Element, [string]$ElementId,
        [string]$Mode, [string]$Detail,
        [ValidateSet('Critique', 'Élevée', 'Moyenne')][string]$Criticite
    )
    $script:Dependencies.Add([pscustomobject]@{
        'Objet impacté' = $ObjetNom
        'Objet Id'      = $ObjetId
        'Type objet'    = $ObjetType
        'Famille'       = (Get-DepFamily $Categorie)
        'Catégorie'     = $Categorie
        'Élément'       = $Element
        'Élément Id'    = $ElementId
        'Mode'          = $Mode
        'Détail'        = $Detail
        'Criticité'     = $Criticite
    })
}

function Find-TargetsInJson {
    # Pour les configurations dont le schéma varie beaucoup (méthodes d'authentification,
    # revues d'accès, B2B, RBAC Intune...), on cherche simplement l'Id du groupe dans le JSON.
    # Moins précis qu'une lecture champ par champ, mais rien ne passe à travers.
    param(
        [object[]]$Objects, [string]$Categorie, [string]$Criticite,
        [string]$Prefix = '', [switch]$IncludeAUs
    )
    $ids = @($script:TargetGroups.Keys)
    if ($IncludeAUs) { $ids += @($script:TargetAUs.Keys) }
    if ($ids.Count -eq 0) { return }
    foreach ($o in @($Objects)) {
        if ($null -eq $o) { continue }
        $json = $o | ConvertTo-Json -Depth 30 -Compress
        foreach ($id in $ids) {
            if ($json -match [regex]::Escape($id)) {
                Add-Dependency -ObjetId $id -ObjetNom (Get-TargetName $id) -ObjetType (Get-TargetType $id) `
                    -Categorie $Categorie -Element ($Prefix + (Get-ObjName $o)) -ElementId $o.id `
                    -Mode 'Référencé' -Detail 'Id trouvé dans la configuration' -Criticite $Criticite
            }
        }
    }
}

function Get-DepStats {
    param([string]$Id)
    $d = @($script:DepsById[$Id])
    if ($d.Count -eq 0 -or $null -eq $d[0]) { return @{ Count = 0; Max = '-'; Cats = '' } }
    $max = ($d | Sort-Object { $script:Rank[$_.'Criticité'] } -Descending | Select-Object -First 1).'Criticité'
    return @{ Count = $d.Count; Max = $max; Cats = (($d.'Catégorie' | Select-Object -Unique) -join ' | ') }
}

function Get-Recommendation {
    # Proposition de remplacement selon la nature des groupes sources.
    # C'est une piste de travail, pas une décision : à valider objet par objet.
    param([string]$Rule)
    $src = @(Get-RuleGuids $Rule | Where-Object { $script:SourceInfo.ContainsKey($_) } | ForEach-Object { $script:SourceInfo[$_] })
    if ($src.Count -eq 0) {
        return 'Aucune source résolue : vérifier la règle (groupe source supprimé ?) puis convertir en Assigné.'
    }
    if (@($src | Where-Object { $_.Type -ne 'group' -or $_.Name -eq '' }).Count -gt 0) {
        return 'Source introuvable ou qui n''est pas un groupe : nettoyer la règle avant toute conversion.'
    }
    if (@($src | Where-Object { $_.Sync -ne $true }).Count -eq 0) {
        return 'Sources synchronisées AD : poser un attribut côté AD (extensionAttribute1-15, department...) répliqué par Entra Connect, puis règle attributaire (ex. user.extensionAttribute5 -eq "XXX").'
    }
    if (@($src | Where-Object { $_.Dynamic -ne $true }).Count -eq 0) {
        return 'Sources dynamiques : reprendre leurs critères dans la règle du groupe cible (fusion par -or).'
    }
    return 'Sources assignées cloud : convertir en Assigné (la conversion vide les membres : les exporter avant) ou piloter par package d''accès / script.'
}

#endregion

#region Page HTML : fiches par objet et carte des dépendances

function New-DependencyMap {
    # Page autonome : les données sont embarquées, aucun serveur nécessaire.
    # La carte utilise vis-network (CDN). Hors ligne, les fiches restent utilisables.
    param([Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)][object]$Data)

    # -InputObject et non le pipeline : on garde les tableaux à un élément tels quels.
    # '</' est neutralisé pour qu'une valeur ne puisse pas fermer la balise <script>.
    $json = (ConvertTo-Json -InputObject $Data -Depth 10 -Compress) -replace '</', '<\/'

    $template = @'
<!doctype html>
<html lang="fr">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Audit memberOf</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Atkinson+Hyperlegible:wght@400;700&family=Barlow+Condensed:wght@500;600&family=JetBrains+Mono&display=swap" rel="stylesheet">
<style>
:root{
  --ink:#1f2933;--ink-soft:#52606d;--paper:#e9edf0;--sheet:#f8f9fa;--rule:#cbd2d9;
  --tape:#f5c518;--tape-ink:#1f2933;
  --crit:#c8322b;--high:#b8730a;--mid:#5b6773;
  --code:#1f2933;--code-ink:#e4e7eb;
}
@media (prefers-color-scheme:dark){:root{
  --ink:#e4e7eb;--ink-soft:#9aa5b1;--paper:#1b2229;--sheet:#232c35;--rule:#3a4652;
  --crit:#f0645c;--high:#e5a33a;--mid:#9aa5b1;--code:#151b21;
}}
*{box-sizing:border-box}
body{margin:0;background:var(--paper);color:var(--ink);font:16px/1.55 "Atkinson Hyperlegible",system-ui,sans-serif}
.wrap{max-width:1180px;margin:0 auto;padding:0 20px}
.tape{display:inline-block;background:var(--tape);color:var(--tape-ink);font-family:"Barlow Condensed",system-ui,sans-serif;font-weight:600;line-height:1.15;padding:3px 11px 4px;border-radius:2px;box-shadow:0 1px 0 rgba(0,0,0,.2)}
header.top{padding:40px 0 8px}
header.top h1{margin:0;font-size:clamp(30px,5vw,46px);transform:rotate(-1deg);transform-origin:left center}
.summary{max-width:66ch;margin:22px 0 8px;font-size:19px;line-height:1.5}
.summary .c{color:var(--crit);font-weight:700}
.summary .warn{display:block;margin-top:8px;font-size:16px;color:var(--high)}
.meta{color:var(--ink-soft);font-size:14px;margin:0}
nav.tabs{display:flex;gap:4px;border-bottom:2px solid var(--ink);margin-top:26px}
nav.tabs button{font:600 19px/1 "Barlow Condensed",system-ui,sans-serif;background:none;border:2px solid transparent;border-bottom:0;color:var(--ink-soft);padding:10px 16px 8px;margin-bottom:-2px;border-radius:4px 4px 0 0;cursor:pointer}
nav.tabs button.on{color:var(--ink);background:var(--sheet);border-color:var(--ink);border-bottom:2px solid var(--sheet)}
button:focus-visible,input:focus-visible,select:focus-visible,a:focus-visible{outline:3px solid var(--tape);outline-offset:2px}
main{padding:24px 0 36px}
.tools{display:flex;gap:10px;flex-wrap:wrap;margin-bottom:34px}
.tools input,.tools select{font:inherit;font-size:15px;color:var(--ink);background:var(--sheet);border:1px solid var(--rule);border-radius:4px;padding:9px 12px;flex:1 1 240px}
.grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(330px,1fr));gap:38px 22px;align-items:start}
.fiche{position:relative;background:var(--sheet);border:1px solid var(--rule);border-top:4px solid var(--mid);border-radius:4px;padding:30px 18px 16px}
.fiche.crit{border-top-color:var(--crit)}
.fiche.high{border-top-color:var(--high)}
.fiche>.tape{position:absolute;top:-17px;left:14px;font-size:20px;max-width:calc(100% - 28px);overflow-wrap:anywhere}
.fiche .line{display:flex;justify-content:space-between;gap:12px;font-size:14px;color:var(--ink-soft)}
.lvl{font-weight:700;white-space:nowrap}
.lvl.crit{color:var(--crit)}
.lvl.high{color:var(--high)}
pre.rule{margin:12px 0 8px;background:var(--code);color:var(--code-ink);font:12.5px/1.5 "JetBrains Mono",ui-monospace,monospace;padding:10px 12px;border-radius:3px;white-space:pre-wrap;word-break:break-all}
.src{font-size:14px;color:var(--ink-soft);margin:0 0 12px}
.usage{border-left:5px solid var(--c,#8795a1);padding:3px 0 3px 12px;margin:12px 0}
.usage h4{margin:0;font:600 17px/1.2 "Barlow Condensed",system-ui,sans-serif}
.usage h4 small{font:400 13px "Atkinson Hyperlegible",system-ui,sans-serif;color:var(--ink-soft);margin-left:6px}
.usage ul{list-style:none;margin:3px 0 0;padding:0}
.usage li{font-size:15px;padding:2px 0;overflow-wrap:anywhere}
.usage li span{color:var(--ink-soft);font-size:13px}
.usage a{color:inherit;font-size:13px;text-decoration-color:var(--tape);text-decoration-thickness:2px;text-underline-offset:3px;margin-left:6px}
.todo{margin:16px 0 0;padding-top:12px;border-top:1px dashed var(--rule);font-size:14px}
.todo b{font:600 16px "Barlow Condensed",system-ui,sans-serif;margin-right:4px}
.none{color:var(--ink-soft);font-size:15px}
#net{height:72vh;min-height:440px;background:var(--sheet);border:1px solid var(--rule);border-radius:4px}
.legend{display:flex;flex-wrap:wrap;gap:6px 18px;margin-top:12px;font-size:14px;color:var(--ink-soft)}
.legend i{display:inline-block;width:20px;height:5px;border-radius:3px;vertical-align:middle;margin-right:6px}
.empty{color:var(--ink-soft);padding:32px;text-align:center}
footer{color:var(--ink-soft);font-size:13px;padding:0 0 32px}
@media (prefers-reduced-motion:no-preference){
  header.top h1{animation:stick .45s cubic-bezier(.2,.8,.3,1.2) both}
  @keyframes stick{from{opacity:0;transform:rotate(-5deg) translateY(-8px)}to{opacity:1;transform:rotate(-1deg)}}
}
@media (max-width:520px){.grid{grid-template-columns:1fr}.summary{font-size:17px}}
</style>
</head>
<body>
<div class="wrap">
<header class="top">
  <h1 class="tape">Audit memberOf</h1>
  <p class="summary" id="summary"></p>
  <p class="meta" id="meta"></p>
</header>
<nav class="tabs" aria-label="Vues">
  <button data-tab="cards" class="on">Fiches par objet</button>
  <button data-tab="map">Carte des dépendances</button>
</nav>
<main>
  <section id="cards">
    <div class="tools">
      <input id="q" type="search" placeholder="Chercher un groupe, une stratégie, une application, un site" aria-label="Chercher">
      <select id="fam" aria-label="Type d'usage"><option value="">Tous les usages</option></select>
    </div>
    <div class="grid" id="grid"></div>
  </section>
  <section id="map" hidden>
    <div id="net"></div>
    <div class="legend" id="legend"></div>
  </section>
</main>
<footer>Rapport généré en lecture seule par Invoke-EntraMemberOfAudit.ps1, référence Microsoft MC1448379.</footer>
</div>
<script>const DATA = __DATA__;</script>
<script>
// Couleurs des cordons de brassage : une couleur par type d'usage, identique dans les fiches et la carte
const FAM = {"Licences":"#8e44ad","Accès conditionnel":"#c0392b","Rôles":"#1f78c1","MFA / authentification":"#d6589f","Intune":"#2e8b57","Applications":"#1aa3b8","Teams":"#5b5fc7","SharePoint":"#138a8a","Exchange":"#8b5a2b","Règles chaînées":"#e07b1a","Entitlement Management":"#c99a00","B2B":"#a0526a","Imbrication":"#8795a1","Revues d'accès":"#6b8e23","Autre":"#8795a1"};
const LINK = {"Teams":"ouvrir dans Teams","SharePoint":"ouvrir le site"};
const arr = x => Array.isArray(x) ? x : (x == null ? [] : [x]);
const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const cls = c => c === 'Critique' ? 'crit' : c === 'Élevée' ? 'high' : '';
const plural = (n, one, many) => `${n} ${n > 1 ? many : one}`;
const order = arr(DATA.meta.families);
const objects = arr(DATA.objects);
const m = DATA.meta;

document.getElementById('summary').innerHTML =
  `${plural(m.groups, 'groupe', 'groupes')}, ${plural(m.aus, 'unité administrative', 'unités administratives')} et ${plural(m.em, 'stratégie EM', 'stratégies EM')} utilisent memberOf. ` +
  `L'audit a trouvé ${plural(m.deps, 'usage', 'usages')}, dont <span class="c">${plural(m.critical, 'critique', 'critiques')}</span>. ` +
  `Ces objets ne seront plus recalculés à partir du 3 novembre 2026, ${m.daysLeft > 0 ? 'dans ' + plural(m.daysLeft, 'jour', 'jours') : 'échéance passée'}.` +
  (m.errors > 0 ? `<span class="warn">${plural(m.errors, 'collecte a échoué', 'collectes ont échoué')} : l'audit est incomplet, voir l'onglet Erreurs collecte du classeur Excel.</span>` : '');
document.getElementById('meta').textContent = `Tenant ${m.tenant}, audit du ${m.date}, script v${m.version}`;

const famSel = document.getElementById('fam');
order.filter(f => objects.some(o => arr(o.deps).some(d => d.family === f)))
     .forEach(f => famSel.insertAdjacentHTML('beforeend', `<option>${esc(f)}</option>`));

function fiche(o) {
  const by = {};
  arr(o.deps).forEach(d => { (by[d.family] = by[d.family] || []).push(d); });
  const fams = Object.keys(by).sort((a, b) => order.indexOf(a) - order.indexOf(b));
  const usages = fams.map(f => `<section class="usage" style="--c:${FAM[f] || '#8795a1'}">
      <h4>${esc(f)}<small>${by[f].length}</small></h4>
      <ul>${by[f].map(d => `<li>${esc(d.element || d.category)}${d.mode ? ` <span>${esc(d.mode)}</span>` : ''}${
        /^https?:\/\//.test(d.detail || '') ? `<a href="${esc(d.detail)}" target="_blank" rel="noopener">${LINK[f] || 'ouvrir'}</a>` : ''}</li>`).join('')}</ul>
    </section>`).join('');
  const src = arr(o.sources).map(s => esc(s.name)).join(', ');
  const level = o.crit === '-' ? 'Aucun usage trouvé' : o.crit;
  const members = o.members != null ? `, ${plural(o.members, 'membre', 'membres')}` : '';
  return `<article class="fiche ${cls(o.crit)}">
    <span class="tape">${esc(o.name)}</span>
    <div class="line"><span>${esc(o.type)}${members}</span><span class="lvl ${cls(o.crit)}">${esc(level)}</span></div>
    <pre class="rule">${esc(o.rule)}</pre>
    <p class="src">Groupes sources : ${src || 'aucun résolu'}</p>
    ${usages || '<p class="none">Aucun usage trouvé dans le tenant.</p>'}
    <p class="todo"><b>À faire</b>${esc(o.reco)}</p>
  </article>`;
}

function renderCards() {
  const q = document.getElementById('q').value.trim().toLowerCase();
  const f = famSel.value;
  const rank = {'Critique':3, 'Élevée':2, 'Moyenne':1};
  const list = objects
    .filter(o => (!q || JSON.stringify(o).toLowerCase().includes(q)) && (!f || arr(o.deps).some(d => d.family === f)))
    .sort((a, b) => (rank[b.crit] || 0) - (rank[a.crit] || 0) || a.name.localeCompare(b.name));
  document.getElementById('grid').innerHTML = list.map(fiche).join('') ||
    (objects.length ? '<p class="empty">Aucun objet ne correspond à cette recherche.</p>' : '<p class="empty">Aucun objet n\'utilise memberOf dans ce tenant.</p>');
}
document.getElementById('q').addEventListener('input', renderCards);
famSel.addEventListener('change', renderCards);
renderCards();

let network = null;
function renderMap() {
  if (network) return;
  const el = document.getElementById('net');
  if (!window.vis) { el.innerHTML = '<p class="empty">La carte a besoin de la bibliothèque vis-network, qui n\'a pas pu être chargée (poste hors ligne ?). Les fiches restent utilisables.</p>'; return; }
  if (!objects.length) { el.innerHTML = '<p class="empty">Aucun objet n\'utilise memberOf : rien à cartographier.</p>'; return; }
  const css = getComputedStyle(document.documentElement);
  const v = n => css.getPropertyValue(n).trim();
  const ink = v('--ink'), soft = v('--ink-soft');
  const border = {'Critique':v('--crit'), 'Élevée':v('--high')};
  const nodes = new Map(), edges = [];
  const objIds = new Set(objects.map(o => o.id));
  // Chaque groupe concerné est posé sur une grille, du plus utilisé au moins utilisé :
  // seuls ses usages s'organisent autour de lui, les grappes ne se chevauchent pas.
  const cols = Math.ceil(Math.sqrt(objects.length));
  [...objects].sort((a, b) => arr(b.deps).length - arr(a.deps).length).forEach((o, i) => nodes.set('o:' + o.id, {
    x:(i % cols) * 900, y:Math.floor(i / cols) * 750, fixed:true,
    id:'o:' + o.id, label:o.name, shape:'box', margin:{top:8,bottom:8,left:12,right:12}, borderWidth:3,
    widthConstraint:{maximum:200}, shapeProperties:{borderRadius:2},
    color:{background:v('--tape'), border:border[o.crit] || v('--mid'), highlight:{background:v('--tape'), border:ink}},
    font:{face:'Barlow Condensed', size:19, color:'#1f2933'}, title:`${o.type}, criticité ${o.crit}`
  }));
  objects.forEach(o => {
    arr(o.sources).forEach(s => {
      const sid = objIds.has(s.id) ? 'o:' + s.id : 's:' + s.id;
      if (!nodes.has(sid)) nodes.set(sid, {id:sid, label:s.name, shape:'box', borderWidth:1, shapeProperties:{borderDashes:[4,3], borderRadius:2},
        color:{background:v('--sheet'), border:soft}, font:{face:'Atkinson Hyperlegible', size:13, color:ink}, title:'Groupe source cité dans la règle'});
      edges.push({from:sid, to:'o:' + o.id, arrows:{to:{scaleFactor:.6}}, dashes:[5,4], color:{color:soft}, width:1.5, title:'memberOf'});
    });
    arr(o.deps).forEach(d => {
      if (d.family === 'Règles chaînées' && objIds.has(String(d.elementId || '').toLowerCase())) return;
      const key = 'e:' + d.family + '|' + (d.elementId || d.element);
      const c = FAM[d.family] || '#8795a1';
      if (!nodes.has(key)) nodes.set(key, {id:key, label:d.element || d.category, shape:'dot', size:9, color:{background:c, border:c},
        font:{face:'Atkinson Hyperlegible', size:13, color:ink}, widthConstraint:{maximum:170}, title:`${d.family} : ${d.category}`});
      edges.push({from:'o:' + o.id, to:key, color:{color:c}, width:3, dashes:d.mode === 'Exclusion' ? [6,4] : false, title:d.mode || ''});
    });
  });
  network = new vis.Network(el, {nodes:new vis.DataSet([...nodes.values()]), edges:new vis.DataSet(edges)}, {
    physics:{solver:'forceAtlas2Based', forceAtlas2Based:{gravitationalConstant:-110, centralGravity:0.002, springLength:150, avoidOverlap:0.8}, stabilization:{iterations:400}},
    interaction:{hover:true, tooltipDelay:120}, edges:{smooth:{type:'continuous'}}
  });
  // Une fois le calcul terminé : on fige la carte et on cadre tout dans la zone visible
  network.once('stabilizationIterationsDone', () => { network.setOptions({physics:false}); network.fit({animation:false}); });
  const used = order.filter(f => objects.some(o => arr(o.deps).some(d => d.family === f)));
  document.getElementById('legend').innerHTML = used.map(f => `<span><i style="background:${FAM[f]}"></i>${esc(f)}</span>`).join('') +
    '<span>Trait pointillé : lien memberOf ou exclusion</span>';
}

document.querySelectorAll('nav.tabs button').forEach(b => b.addEventListener('click', () => {
  document.querySelectorAll('nav.tabs button').forEach(x => x.classList.toggle('on', x === b));
  document.getElementById('cards').hidden = b.dataset.tab !== 'cards';
  document.getElementById('map').hidden = b.dataset.tab !== 'map';
  if (b.dataset.tab === 'map') renderMap();
}));
</script>
<script src="https://unpkg.com/vis-network@9.1.9/standalone/umd/vis-network.min.js"></script>
</body>
</html>
'@
    $html = $template.Replace('__DATA__', $json)
    # Écriture UTF-8 sans BOM, identique sous 5.1 et 7
    [System.IO.File]::WriteAllText($Path, $html, (New-Object System.Text.UTF8Encoding($false)))
}

#endregion

# =============================================================================================
#  Présentation, prérequis, confirmation
# =============================================================================================

$interactive = Test-Interactive
$started     = Get-Date

if ($interactive) { Show-Banner }

$moduleStatus = @(Get-ModuleStatus)
$missing      = @($moduleStatus | Where-Object { $_.State -ne 'OK' })

if ($interactive) { Show-Prerequisites -ModuleStatus $moduleStatus }

if ($missing.Count -gt 0) {
    $doInstall = $InstallMissingModules
    if (-not $doInstall -and $interactive) {
        $doInstall = Read-YesNo -Question ("{0} module(s) à installer ou mettre à jour. Le faire maintenant (portée utilisateur, sans droits admin) ?" -f $missing.Count)
    }
    if (-not $doInstall) {
        Write-Host ''
        Write-Host '  Installation manuelle :' -ForegroundColor Yellow
        foreach ($m in $missing) {
            Write-Host ("    Install-Module {0} -MinimumVersion {1} -Scope CurrentUser" -f $m.Name, $m.MinVersion) -ForegroundColor Yellow
        }
        Write-Host ''
        Write-Host '  Prérequis non satisfaits, audit non lancé.' -ForegroundColor Red
        exit 1
    }
    Install-RequiredModules -Missing $missing
}

if ($interactive) {
    if (-not (Read-YesNo -Question 'Lancer l''audit ?')) {
        Write-Host '  Audit annulé, rien n''a été fait.' -ForegroundColor Yellow
        return
    }
    Write-Host ''
}

foreach ($m in $RequiredModules) {
    Import-Module -Name $m.Name -MinimumVersion $m.MinVersion -ErrorAction Stop
}

try { Start-Transcript -Path $LogPath -Append | Out-Null; $transcript = $true }
catch { $transcript = $false; Write-Warning "Journal non créé ($LogPath) : $($_.Exception.Message)" }

# =============================================================================================
#  Connexion
# =============================================================================================

Write-Step 'Connexion à Microsoft Graph'
if ($PSCmdlet.ParameterSetName -eq 'AppCert') {
    Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -NoWelcome
}
else {
    $cp = @{ Scopes = @($ReadScopes.Keys); NoWelcome = $true }
    if ($TenantId)      { $cp.TenantId = $TenantId }
    if ($UseDeviceCode) { $cp.UseDeviceCode = $true }
    Connect-MgGraph @cp
}

try {
    $ctx = Get-MgContext

    # Le script ne fait que des GET, mais si le jeton porte des droits d'écriture on le signale :
    # en audit, mieux vaut un jeton qui ne peut physiquement rien modifier.
    $writeScopes = @($ctx.Scopes | Where-Object { $_ -match 'ReadWrite|\.Write' })
    if ($writeScopes.Count -gt 0) {
        Write-Warning ("Le jeton contient des droits d'écriture ({0}). Le script reste en lecture seule, mais un jeton limité à Read est préférable." -f ($writeScopes -join ', '))
    }

    $org = Invoke-GraphGetOne -Uri "$GraphV1/organization?`$select=id,displayName" -Source 'Organisation'
    $tenantName = if ($org -and $org.value) { $org.value[0].displayName } else { $ctx.TenantId }
    Write-Step "Tenant : $tenantName ($($ctx.TenantId)) - connexion $($ctx.AuthType) - $(if ($ctx.Account) { $ctx.Account } else { $ctx.ClientId })"

    # En délégué, on vérifie que le compte porte un rôle suffisant. Un rôle PIM éligible
    # mais non activé n'apparaît pas ici : l'activer avant de lancer l'audit.
    if ($ctx.AuthType -eq 'Delegated') {
        $myRoles = Invoke-GraphGetAll -Source 'Rôles du compte connecté' `
            -Uri "$GraphV1/me/transitiveMemberOf/microsoft.graph.directoryRole?`$select=displayName,roleTemplateId"
        $okRole = @($myRoles | Where-Object { $SufficientRoles.ContainsKey([string]$_.roleTemplateId) })
        if ($okRole.Count -gt 0) {
            Write-Step ("Rôle du compte : {0}" -f (($okRole | ForEach-Object { $_.displayName }) -join ', '))
        }
        else {
            Write-Warning 'Le compte n''a ni Lecteur général ni Administrateur général actif. Certaines sections risquent d''être incomplètes (voir l''onglet Erreurs collecte).'
        }
    }

    # =========================================================================================
    #  1. Inventaire des objets qui utilisent memberOf
    # =========================================================================================

    Write-Step '1/4 - Groupes dynamiques'
    $grpSelect = 'id,displayName,description,membershipRule,membershipRuleProcessingState,groupTypes,securityEnabled,mailEnabled,mail,onPremisesSyncEnabled,assignedLicenses,resourceProvisioningOptions,createdDateTime'
    $dynGroups = Invoke-GraphGetAll -Source 'Groupes dynamiques' `
        -Uri "$GraphV1/groups?`$filter=groupTypes/any(c:c eq 'DynamicMembership')&`$select=$grpSelect&`$top=999"

    # Filtrage côté client : un filtre Graph sur startsWith(membershipRule) raterait les règles
    # où memberOf n'est pas en tête (ex. "(user.department -eq 'X') -or (user.memberof ...)").
    $impactedGroups = @($dynGroups | Where-Object { $_.membershipRule -match $MemberOfRegex })
    foreach ($g in $impactedGroups) { $script:TargetGroups[$g.id.ToLower()] = $g.displayName }
    Write-Step ("      {0} groupe(s) dynamique(s), dont {1} avec memberOf" -f $dynGroups.Count, $impactedGroups.Count)

    Write-Step '      Unités administratives dynamiques'
    $allAUs = Invoke-GraphGetAll -Source 'Unités administratives' `
        -Uri "$GraphBeta/administrativeUnits?`$select=id,displayName,description,membershipType,membershipRule,membershipRuleProcessingState,visibility&`$top=999"
    $dynAUs      = @($allAUs | Where-Object { $_.membershipType -eq 'Dynamic' })
    $impactedAUs = @($dynAUs | Where-Object { $_.membershipRule -match $MemberOfRegex })
    foreach ($a in $impactedAUs) { $script:TargetAUs[$a.id.ToLower()] = $a.displayName }
    Write-Step ("      {0} AU dynamique(s), dont {1} avec memberOf" -f $dynAUs.Count, $impactedAUs.Count)

    Write-Step '      Stratégies Entitlement Management'
    $emPolicies = Invoke-GraphGetAll -Source 'EM - stratégies' `
        -Uri "$GraphV1/identityGovernance/entitlementManagement/assignmentPolicies?`$expand=accessPackage"
    # La règle d'affectation automatique est enfouie dans specificAllowedTargets : plutôt que de
    # dépendre du schéma exact, on cherche memberOf dans le JSON et on extrait membershipRule.
    $impactedEm = @(foreach ($p in $emPolicies) {
        $json = $p | ConvertTo-Json -Depth 30 -Compress
        if ($json -match $MemberOfRegex) {
            $rule = ''
            $m = [regex]::Match($json, '"membershipRule"\s*:\s*"((?:[^"\\]|\\.)*)"')
            if ($m.Success) { $rule = [regex]::Unescape($m.Groups[1].Value) }
            [pscustomobject]@{ Policy = $p; Rule = $rule }
        }
    })
    Write-Step ("      {0} stratégie(s) EM, dont {1} avec memberOf" -f $emPolicies.Count, $impactedEm.Count)

    # =========================================================================================
    #  2. Groupes sources cités dans les règles
    # =========================================================================================

    Write-Step '2/4 - Groupes sources'
    $sourceMap = @{}   # id source -> objets qui le citent
    $addSource = {
        param($rule, $refName)
        foreach ($sid in (Get-RuleGuids $rule)) {
            if (-not $sourceMap.ContainsKey($sid)) { $sourceMap[$sid] = [System.Collections.Generic.List[string]]::new() }
            $sourceMap[$sid].Add($refName)
        }
    }
    foreach ($g in $impactedGroups) { & $addSource $g.membershipRule "Groupe : $($g.displayName)" }
    foreach ($a in $impactedAUs)    { & $addSource $a.membershipRule "AU : $($a.displayName)" }
    foreach ($e in $impactedEm)     { & $addSource $e.Rule "EM : $($e.Policy.accessPackage.displayName) / $($e.Policy.displayName)" }

    $script:SourceInfo = @{}
    $sourceRows = foreach ($sid in $sourceMap.Keys) {
        # directoryObjects plutôt que groups : une règle peut citer un Id qui n'est plus un groupe
        $o = Invoke-GraphGetOne -Uri "$GraphV1/directoryObjects/$sid" -Quiet
        $type = 'Introuvable (supprimé ?)'; $name = ''; $sync = ''; $isDyn = ''; $count = $null; $rule = ''
        if ($o) {
            $name = Get-ObjName $o
            $type = ($o.'@odata.type' -replace '#microsoft.graph.', '')
            if ($type -eq 'group') {
                $sync  = [bool]$o.onPremisesSyncEnabled
                $isDyn = [bool](@($o.groupTypes) -contains 'DynamicMembership')
                $rule  = $o.membershipRule
                $count = Get-MemberCount "groups/$sid"
            }
        }
        $script:SourceInfo[$sid] = [pscustomobject]@{ Name = $name; Type = $type; Sync = $sync; Dynamic = $isDyn }
        [pscustomobject]@{
            'Groupe source'        = $name
            'Id'                   = $sid
            'Type objet'           = $type
            'Synchronisé AD'       = $sync
            'Dynamique'            = $isDyn
            'Règle (si dynamique)' = $rule
            'Nb membres directs'   = $count
            'Référencé par'        = ($sourceMap[$sid] | Select-Object -Unique) -join ' | '
        }
        if ($IncludeSourceGroupDependencies -and $type -eq 'group' -and -not $script:TargetGroups.ContainsKey($sid)) {
            $script:TargetGroups[$sid] = "$name (source)"
        }
    }
    Write-Step ("      {0} objet(s) source(s) résolu(s)" -f @($sourceRows).Count)

    # =========================================================================================
    #  3. Dépendances
    # =========================================================================================

    Write-Step '3/4 - Recherche des dépendances'
    $targetGroupIds = @($script:TargetGroups.Keys)

    # Licences attribuées par groupe : c'est la conséquence la plus visible d'un gel
    # (nouveaux arrivants sans licence, départs qui continuent de consommer).
    Write-Step '      Licences'
    $skuMap = @{}
    foreach ($s in (Invoke-GraphGetAll -Uri "$GraphV1/subscribedSkus" -Source 'Références de licences')) {
        $skuMap[[string]$s.skuId] = $s.skuPartNumber
    }
    $groupsWithLic = @($impactedGroups)
    if ($IncludeSourceGroupDependencies) {
        foreach ($sid in $targetGroupIds | Where-Object { $script:SourceInfo.ContainsKey($_) }) {
            $sg = Invoke-GraphGetOne -Uri "$GraphV1/groups/$sid`?`$select=id,displayName,assignedLicenses" -Quiet
            if ($sg) { $groupsWithLic += $sg }
        }
    }
    foreach ($g in $groupsWithLic) {
        foreach ($l in @($g.assignedLicenses)) {
            if (-not $l) { continue }
            $sku = if ($skuMap.ContainsKey([string]$l.skuId)) { $skuMap[[string]$l.skuId] } else { [string]$l.skuId }
            Add-Dependency -ObjetId $g.id.ToLower() -ObjetNom (Get-TargetName $g.id.ToLower()) -ObjetType 'Groupe' `
                -Categorie 'Licences (attribution par groupe)' -Element $sku -ElementId $l.skuId -Mode 'Attribution' `
                -Detail ("{0} plan(s) de service désactivé(s)" -f @($l.disabledPlans).Count) -Criticite 'Critique'
        }
    }

    # Accès conditionnel : une exclusion figée est aussi dangereuse qu'une inclusion figée
    # (ex. groupe d'exclusion MFA qui garde d'anciens membres).
    Write-Step '      Accès conditionnel'
    $caPolicies = Invoke-GraphGetAll -Uri "$GraphV1/identity/conditionalAccess/policies" -Source 'Accès conditionnel'
    foreach ($p in $caPolicies) {
        $u = $p.conditions.users
        foreach ($pair in @(@('Inclusion', @($u.includeGroups)), @('Exclusion', @($u.excludeGroups)))) {
            foreach ($gid in $pair[1]) {
                if (-not $gid) { continue }
                $gid = ([string]$gid).ToLower()
                if ($script:TargetGroups.ContainsKey($gid)) {
                    Add-Dependency -ObjetId $gid -ObjetNom (Get-TargetName $gid) -ObjetType 'Groupe' `
                        -Categorie 'Accès conditionnel' -Element $p.displayName -ElementId $p.id -Mode $pair[0] `
                        -Detail "État : $($p.state)" -Criticite 'Critique'
                }
            }
        }
    }

    # Rôles Entra : groupe porteur d'un rôle, ou rôle délégué sur une AU impactée
    Write-Step '      Rôles Entra (actifs et éligibles)'
    $roleDefs = @{}
    foreach ($rd in (Invoke-GraphGetAll -Uri "$GraphV1/roleManagement/directory/roleDefinitions?`$select=id,displayName" -Source 'Définitions de rôles')) {
        $roleDefs[$rd.id] = $rd.displayName
    }
    $roleSets = @(
        @('Actif',          (Invoke-GraphGetAll -Uri "$GraphV1/roleManagement/directory/roleAssignments" -Source 'Rôles actifs')),
        @('Éligible (PIM)', (Invoke-GraphGetAll -Uri "$GraphV1/roleManagement/directory/roleEligibilitySchedules" -Source 'Rôles éligibles'))
    )
    foreach ($set in $roleSets) {
        foreach ($ra in $set[1]) {
            $roleName = if ($roleDefs.ContainsKey($ra.roleDefinitionId)) { $roleDefs[$ra.roleDefinitionId] } else { $ra.roleDefinitionId }
            $principal = ([string]$ra.principalId).ToLower()
            if ($script:TargetGroups.ContainsKey($principal)) {
                Add-Dependency -ObjetId $principal -ObjetNom (Get-TargetName $principal) -ObjetType 'Groupe' -Categorie 'Rôle Entra' `
                    -Element $roleName -ElementId $ra.id -Mode $set[0] -Detail "Portée : $($ra.directoryScopeId)" -Criticite 'Critique'
            }
            if ($ra.directoryScopeId -match '^/administrativeUnits/(.+)$') {
                $auId = $Matches[1].ToLower()
                if ($script:TargetAUs.ContainsKey($auId)) {
                    $pr = Invoke-GraphGetOne -Uri "$GraphV1/directoryObjects/$($ra.principalId)" -Quiet
                    Add-Dependency -ObjetId $auId -ObjetNom (Get-TargetName $auId) -ObjetType 'Unité administrative' `
                        -Categorie 'Rôle délégué sur AU' -Element $roleName -ElementId $ra.id -Mode $set[0] `
                        -Detail ("Administrateur : {0}" -f $(if ($pr) { Get-ObjName $pr } else { $ra.principalId })) -Criticite 'Critique'
                }
            }
        }
    }

    Write-Step '      Applications d''entreprise, imbrications, Teams / M365'
    foreach ($gid in $targetGroupIds) {
        $gname = Get-TargetName $gid
        foreach ($ar in (Invoke-GraphGetAll -Uri "$GraphV1/groups/$gid/appRoleAssignments" -Source "Applications - $gname")) {
            Add-Dependency -ObjetId $gid -ObjetNom $gname -ObjetType 'Groupe' -Categorie 'Application d''entreprise' `
                -Element $ar.resourceDisplayName -ElementId $ar.resourceId -Mode 'Affectation' `
                -Detail ("AppRoleId : {0}" -f $ar.appRoleId) -Criticite 'Élevée'
        }
        # memberOf du groupe lui-même : groupes parents, AU, rôles d'annuaire
        foreach ($mo in (Invoke-GraphGetAll -Uri "$GraphV1/groups/$gid/memberOf?`$select=id,displayName" -Source "Appartenances - $gname")) {
            $t = $mo.'@odata.type' -replace '#microsoft.graph.', ''
            $cat = switch ($t) {
                'group'              { 'Membre d''un autre groupe' }
                'administrativeUnit' { 'Membre d''une AU' }
                'directoryRole'      { 'Rôle Entra' }
                default              { "Membre de ($t)" }
            }
            Add-Dependency -ObjetId $gid -ObjetNom $gname -ObjetType 'Groupe' -Categorie $cat `
                -Element (Get-ObjName $mo) -ElementId $mo.id -Mode 'Imbrication' -Detail '' -Criticite 'Moyenne'
        }
    }
    # Groupe Microsoft 365 dynamique : ses membres SONT les membres de l'équipe, du site et de la boîte.
    # Si la règle se fige, les arrivants n'ont pas accès et les partants gardent l'accès.
    # NB : les droits SharePoint donnés directement à un groupe de sécurité ne sont pas visibles via Graph.
    Write-Step '      Teams, SharePoint, boîtes de groupe'
    foreach ($g in $impactedGroups) {
        if (@($g.groupTypes) -notcontains 'Unified') { continue }
        $id = $g.id.ToLower()

        if (@($g.resourceProvisioningOptions) -contains 'Team') {
            $team = Invoke-GraphGetOne -Uri "$GraphV1/teams/$id`?`$select=displayName,webUrl,visibility" -Quiet
            Add-Dependency -ObjetId $id -ObjetNom $g.displayName -ObjetType 'Groupe' -Categorie 'Microsoft Teams' `
                -Element $(if ($team) { $team.displayName } else { $g.displayName }) -ElementId $id `
                -Mode 'Membres = accès à l''équipe' `
                -Detail $(if ($team -and $team.webUrl) { $team.webUrl } else { 'Lien non lu (Team.ReadBasic.All)' }) -Criticite 'Élevée'
        }

        $site = Invoke-GraphGetOne -Uri "$GraphV1/groups/$id/sites/root?`$select=id,displayName,webUrl" -Quiet
        Add-Dependency -ObjetId $id -ObjetNom $g.displayName -ObjetType 'Groupe' -Categorie 'SharePoint' `
            -Element $(if ($site) { $site.displayName } else { 'Site du groupe' }) -ElementId $(if ($site) { $site.id } else { "site-$id" }) `
            -Mode 'Membres = accès au site' `
            -Detail $(if ($site -and $site.webUrl) { $site.webUrl } else { 'Lien non lu (Sites.Read.All ou accès au site)' }) -Criticite 'Élevée'

        if ($g.mail) {
            Add-Dependency -ObjetId $id -ObjetNom $g.displayName -ObjetType 'Groupe' -Categorie 'Exchange (boîte du groupe)' `
                -Element $g.mail -ElementId $g.mail -Mode 'Membres = accès à la boîte' -Detail 'Boîte aux lettres et calendrier du groupe' -Criticite 'Moyenne'
        }
    }

    # Effet domino : un autre groupe ou une AU dynamique qui prend un groupe impacté comme source
    Write-Step '      Règles dynamiques chaînées'
    foreach ($dg in $dynGroups) {
        foreach ($ref in (Get-RuleGuids $dg.membershipRule)) {
            if ($script:TargetGroups.ContainsKey($ref) -and $ref -ne $dg.id.ToLower()) {
                Add-Dependency -ObjetId $ref -ObjetNom (Get-TargetName $ref) -ObjetType 'Groupe' -Categorie 'Règle dynamique chaînée' `
                    -Element $dg.displayName -ElementId $dg.id -Mode 'Source d''une règle' -Detail $dg.membershipRule -Criticite 'Élevée'
            }
        }
    }
    foreach ($au in $dynAUs) {
        foreach ($ref in (Get-RuleGuids $au.membershipRule)) {
            if ($script:TargetGroups.ContainsKey($ref)) {
                Add-Dependency -ObjetId $ref -ObjetNom (Get-TargetName $ref) -ObjetType 'Groupe' -Categorie 'Règle dynamique chaînée' `
                    -Element "AU : $($au.displayName)" -ElementId $au.id -Mode 'Source d''une règle' -Detail $au.membershipRule -Criticite 'Élevée'
            }
        }
    }

    Write-Step '      Entitlement Management'
    $accessPackages = Invoke-GraphGetAll -Uri "$GraphV1/identityGovernance/entitlementManagement/accessPackages?`$select=id,displayName" -Source 'EM - packages'
    $apResources = @{}
    foreach ($ap in $accessPackages) {
        # L'expand des ressources n'est pas disponible sur la liste, d'où un appel par package
        $full = Invoke-GraphGetOne -Source "EM - ressources de $($ap.displayName)" `
            -Uri "$GraphV1/identityGovernance/entitlementManagement/accessPackages/$($ap.id)?`$expand=resourceRoleScopes(`$expand=role,scope)"
        $apResources[$ap.id] = @(if ($full) { $full.resourceRoleScopes })
        foreach ($rrs in $apResources[$ap.id]) {
            $origin = ([string]$rrs.scope.originId).ToLower()
            if ($script:TargetGroups.ContainsKey($origin)) {
                Add-Dependency -ObjetId $origin -ObjetNom (Get-TargetName $origin) -ObjetType 'Groupe' -Categorie 'Entitlement Management' `
                    -Element $ap.displayName -ElementId $ap.id -Mode 'Ressource du package' -Detail "Rôle : $($rrs.role.displayName)" -Criticite 'Élevée'
            }
        }
    }
    # Pour une stratégie EM impactée, la « dépendance » est ce que le package distribue
    foreach ($e in $impactedEm) {
        foreach ($rrs in @($apResources[$e.Policy.accessPackage.id])) {
            if (-not $rrs) { continue }
            Add-Dependency -ObjetId $e.Policy.id -ObjetNom "$($e.Policy.accessPackage.displayName) / $($e.Policy.displayName)" `
                -ObjetType 'Stratégie EM (auto)' -Categorie 'Accès accordé par le package' -Element $rrs.scope.displayName `
                -ElementId $rrs.scope.originId -Mode $rrs.scope.originSystem -Detail "Rôle : $($rrs.role.displayName)" -Criticite 'Élevée'
        }
    }
    Find-TargetsInJson -Objects $emPolicies -Categorie 'Entitlement Management (demandeurs / approbateurs)' -Criticite 'Élevée' -Prefix 'Stratégie : '

    Write-Step '      Revues d''accès, méthodes d''authentification, B2B'
    $reviews = Invoke-GraphGetAll -Uri "$GraphV1/identityGovernance/accessReviews/definitions" -Source 'Revues d''accès'
    Find-TargetsInJson -Objects $reviews -Categorie 'Revue d''accès' -Criticite 'Moyenne' -IncludeAUs

    $amp = Invoke-GraphGetOne -Uri "$GraphV1/policies/authenticationMethodsPolicy" -Source 'Méthodes d''authentification'
    if ($amp) {
        Find-TargetsInJson -Objects @($amp.authenticationMethodConfigurations) -Categorie 'Méthodes d''authentification' -Criticite 'Critique' -Prefix 'Méthode : '
        if ($amp.registrationEnforcement) {
            $campaign = [pscustomobject]@{ id = 'registrationCampaign'; displayName = 'Campagne d''inscription MFA'; data = $amp.registrationEnforcement }
            Find-TargetsInJson -Objects @($campaign) -Categorie 'Méthodes d''authentification' -Criticite 'Critique'
        }
    }

    $ctaDefault  = Invoke-GraphGetOne -Uri "$GraphV1/policies/crossTenantAccessPolicy/default" -Source 'B2B - paramètres par défaut'
    $ctaPartners = Invoke-GraphGetAll -Uri "$GraphV1/policies/crossTenantAccessPolicy/partners" -Source 'B2B - partenaires'
    if ($ctaDefault) { $ctaDefault | Add-Member -NotePropertyName displayName -NotePropertyValue 'Paramètres par défaut' -Force }
    foreach ($cp in $ctaPartners) { $cp | Add-Member -NotePropertyName displayName -NotePropertyValue "Partenaire $($cp.tenantId)" -Force }
    Find-TargetsInJson -Objects (@($ctaDefault) + @($ctaPartners)) -Categorie 'Accès inter-locataires (B2B)' -Criticite 'Élevée'

    if (-not $SkipIntune) {
        Write-Step '      Intune'
        # Endpoints beta : c'est là que les affectations sont exposées de façon homogène.
        $intuneEndpoints = [ordered]@{
            'Profils de configuration'               = 'deviceManagement/deviceConfigurations'
            'Settings Catalog'                       = 'deviceManagement/configurationPolicies'
            'Modèles d''administration (ADMX)'       = 'deviceManagement/groupPolicyConfigurations'
            'Stratégies de conformité'               = 'deviceManagement/deviceCompliancePolicies'
            'Sécurité du point de terminaison'       = 'deviceManagement/intents'
            'Scripts PowerShell'                     = 'deviceManagement/deviceManagementScripts'
            'Scripts shell macOS'                    = 'deviceManagement/deviceShellScripts'
            'Remédiations'                           = 'deviceManagement/deviceHealthScripts'
            'Profils Autopilot'                      = 'deviceManagement/windowsAutopilotDeploymentProfiles'
            'Inscription (ESP, restrictions, WHfB)'  = 'deviceManagement/deviceEnrollmentConfigurations'
            'Mises à jour de fonctionnalités'        = 'deviceManagement/windowsFeatureUpdateProfiles'
            'Mises à jour qualité'                   = 'deviceManagement/windowsQualityUpdateProfiles'
            'Mises à jour de pilotes'                = 'deviceManagement/windowsDriverUpdateProfiles'
            'Applications'                           = 'deviceAppManagement/mobileApps'
            'Configuration d''applications (gérées)' = 'deviceAppManagement/mobileAppConfigurations'
            'Configuration d''applications (MAM)'    = 'deviceAppManagement/targetedManagedAppConfigurations'
            'Protection d''applications iOS'         = 'deviceAppManagement/iosManagedAppProtections'
            'Protection d''applications Android'     = 'deviceAppManagement/androidManagedAppProtections'
            'Protection d''applications Windows'     = 'deviceAppManagement/windowsManagedAppProtections'
        }
        foreach ($label in $intuneEndpoints.Keys) {
            $base = "$GraphBeta/$($intuneEndpoints[$label])"
            # Pour les applications on se limite à celles qui ont au moins une affectation (catalogue souvent volumineux)
            $query = if ($label -eq 'Applications') { "?`$filter=isAssigned eq true&`$expand=assignments" } else { "?`$expand=assignments" }
            try {
                $items = Invoke-GraphGetAll -Uri ($base + $query) -Throw
            }
            catch {
                # Certaines collections (intents notamment) refusent $expand : on lit les affectations une par une
                $items = Invoke-GraphGetAll -Uri $base -Source "Intune - $label"
                foreach ($it in $items) {
                    $as = Invoke-GraphGetAll -Uri "$base/$($it.id)/assignments" -Source "Intune - $label - affectations"
                    $it | Add-Member -NotePropertyName assignments -NotePropertyValue $as -Force
                }
            }
            foreach ($it in $items) {
                foreach ($a in @($it.assignments)) {
                    # Les cibles « Tous les utilisateurs / Tous les appareils » n'ont pas de groupId
                    if (-not $a -or -not $a.target -or -not $a.target.groupId) { continue }
                    $gid = ([string]$a.target.groupId).ToLower()
                    if (-not $script:TargetGroups.ContainsKey($gid)) { continue }
                    $mode = if ($a.target.'@odata.type' -match 'exclusion') { 'Exclusion' } else { 'Inclusion' }
                    $detail = @()
                    if ($a.PSObject.Properties.Name -contains 'intent' -and $a.intent) { $detail += "Intention : $($a.intent)" }
                    $filterId = $a.target.deviceAndAppManagementAssignmentFilterId
                    if ($filterId -and $filterId -ne '00000000-0000-0000-0000-000000000000') {
                        $detail += "Filtre : $($a.target.deviceAndAppManagementAssignmentFilterType) $filterId"
                    }
                    Add-Dependency -ObjetId $gid -ObjetNom (Get-TargetName $gid) -ObjetType 'Groupe' -Categorie "Intune - $label" `
                        -Element (Get-ObjName $it) -ElementId $it.id -Mode $mode -Detail ($detail -join ' ; ') -Criticite 'Élevée'
                }
            }
        }
        $intuneRbac = Invoke-GraphGetAll -Uri "$GraphBeta/deviceManagement/roleAssignments" -Source 'Intune - RBAC'
        Find-TargetsInJson -Objects $intuneRbac -Categorie 'Intune - RBAC (membres / étendue)' -Criticite 'Élevée'
    }

    # =========================================================================================
    #  4. Consolidation et export Excel
    # =========================================================================================

    Write-Step '4/4 - Consolidation et export Excel'
    $script:Rank     = @{ 'Critique' = 3; 'Élevée' = 2; 'Moyenne' = 1 }
    $script:DepsById = $script:Dependencies | Group-Object -Property 'Objet Id' -AsHashTable -AsString
    if (-not $script:DepsById) { $script:DepsById = @{} }
    $memberCounts = @{}

    $groupRows = foreach ($g in $impactedGroups) {
        $id = $g.id.ToLower(); $st = Get-DepStats $id
        [pscustomobject]@{
            'Groupe'                 = $g.displayName
            'Id'                     = $g.id
            'Type'                   = $(if (@($g.groupTypes) -contains 'Unified') { 'Microsoft 365' } elseif ($g.securityEnabled) { 'Sécurité' } else { 'Autre' })
            'Cible'                  = $(if ($g.membershipRule -match '(?i)^\s*\(?\s*device\.') { 'Appareils' } else { 'Utilisateurs' })
            'Règle'                  = $g.membershipRule
            'État traitement'        = $g.membershipRuleProcessingState
            'Nb membres'             = ($memberCounts[$id] = Get-MemberCount "groups/$id")
            'Groupes sources'        = ((Get-RuleGuids $g.membershipRule | ForEach-Object { if ($script:SourceInfo.ContainsKey($_) -and $script:SourceInfo[$_].Name) { $script:SourceInfo[$_].Name } else { $_ } }) -join ' | ')
            'Licences'               = ((@($g.assignedLicenses) | Where-Object { $_ } | ForEach-Object { if ($skuMap.ContainsKey([string]$_.skuId)) { $skuMap[[string]$_.skuId] } else { $_.skuId } }) -join ', ')
            'Nb dépendances'         = $st.Count
            'Criticité max'          = $st.Max
            'Catégories dépendances' = $st.Cats
            'Piste de remplacement'  = (Get-Recommendation $g.membershipRule)
            'Créé le'                = $g.createdDateTime
        }
    }

    $auRows = foreach ($a in $impactedAUs) {
        $id = $a.id.ToLower(); $st = Get-DepStats $id
        [pscustomobject]@{
            'Unité administrative'   = $a.displayName
            'Id'                     = $a.id
            'Règle'                  = $a.membershipRule
            'État traitement'        = $a.membershipRuleProcessingState
            'Nb membres'             = ($memberCounts[$id] = Get-MemberCount "directory/administrativeUnits/$id")
            'Visibilité'             = $a.visibility
            'Nb dépendances'         = $st.Count
            'Criticité max'          = $st.Max
            'Catégories dépendances' = $st.Cats
            'Piste de remplacement'  = (Get-Recommendation $a.membershipRule)
        }
    }

    $emRows = foreach ($e in $impactedEm) {
        $st = Get-DepStats ([string]$e.Policy.id)
        [pscustomobject]@{
            'Package d''accès'       = $e.Policy.accessPackage.displayName
            'Stratégie'              = $e.Policy.displayName
            'Id stratégie'           = $e.Policy.id
            'Règle auto-affectation' = $e.Rule
            'Nb ressources'          = $st.Count
            'Ressources accordées'   = ((@($script:DepsById[[string]$e.Policy.id]) | Where-Object { $_ } | ForEach-Object { $_.'Élément' }) -join ' | ')
            'Piste de remplacement'  = (Get-Recommendation $e.Rule)
        }
    }

    $depRows = @($script:Dependencies | Sort-Object @{ Expression = { $script:Rank[$_.'Criticité'] }; Descending = $true }, 'Objet impacté', 'Catégorie')
    $criticalCount = @($script:Dependencies | Where-Object { $_.'Criticité' -eq 'Critique' }).Count

    # Une entrée par objet impacté, avec toutes ses dépendances : sert à la fois
    # à l'onglet « Vue par objet » et à la page HTML (fiches + carte).
    $mapObjects = [System.Collections.Generic.List[object]]::new()
    $addMapObject = {
        param([string]$Id, [string]$Name, [string]$Type, [string]$Rule)
        $deps = @($script:DepsById[$Id] | Where-Object { $_ })
        $st = Get-DepStats $Id
        $mapObjects.Add([pscustomobject]@{
            id      = $Id
            name    = $Name
            type    = $Type
            rule    = $Rule
            crit    = $st.Max
            members = $memberCounts[$Id]
            reco    = (Get-Recommendation $Rule)
            sources = @(Get-RuleGuids $Rule | ForEach-Object {
                [pscustomobject]@{
                    id   = $_
                    name = $(if ($script:SourceInfo.ContainsKey($_) -and $script:SourceInfo[$_].Name) { $script:SourceInfo[$_].Name } else { $_ })
                }
            })
            deps    = @($deps | ForEach-Object {
                [pscustomobject]@{
                    family    = $_.'Famille'
                    category  = $_.'Catégorie'
                    element   = $_.'Élément'
                    elementId = $_.'Élément Id'
                    mode      = $_.Mode
                    detail    = $_.'Détail'
                    crit      = $_.'Criticité'
                }
            })
        })
    }
    foreach ($g in $impactedGroups) { & $addMapObject $g.id.ToLower() $g.displayName 'Groupe' $g.membershipRule }
    foreach ($a in $impactedAUs)    { & $addMapObject $a.id.ToLower() $a.displayName 'Unité administrative' $a.membershipRule }
    foreach ($e in $impactedEm)     { & $addMapObject ([string]$e.Policy.id) "$($e.Policy.accessPackage.displayName) / $($e.Policy.displayName)" 'Stratégie EM' $e.Rule }

    # Vue par objet : une ligne par groupe / AU, une colonne par famille de dépendance utilisée
    $usedFamilies = @($FamilyOrder | Where-Object { $f = $_; @($script:Dependencies | Where-Object { $_.'Famille' -eq $f }).Count -gt 0 })
    $viewRows = foreach ($o in $mapObjects) {
        $row = [ordered]@{
            'Objet'          = $o.name
            'Type'           = $o.type
            'Criticité max'  = $o.crit
            'Nb dépendances' = @($o.deps).Count
        }
        foreach ($f in $usedFamilies) {
            $row[$f] = (@($o.deps | Where-Object { $_.family -eq $f } | ForEach-Object {
                if ($_.mode -and $_.mode -notmatch '^(Référencé|Attribution|Affectation)$') { "$($_.element) [$($_.mode)]" } else { $_.element }
            }) -join ' | ')
        }
        [pscustomobject]$row
    }

    $catSummary = $script:Dependencies | Group-Object 'Catégorie' | Sort-Object Count -Descending |
        ForEach-Object { [pscustomobject]@{ Indicateur = "Dépendances - $($_.Name)"; Valeur = $_.Count } }

    $summary = @(
        [pscustomobject]@{ Indicateur = 'Tenant';                                   Valeur = "$tenantName ($($ctx.TenantId))" }
        [pscustomobject]@{ Indicateur = 'Date de l''audit';                         Valeur = $started.ToString('yyyy-MM-dd HH:mm') }
        [pscustomobject]@{ Indicateur = 'Compte / application';                     Valeur = $(if ($ctx.Account) { $ctx.Account } else { $ctx.ClientId }) }
        [pscustomobject]@{ Indicateur = 'Mode de connexion';                        Valeur = $ctx.AuthType }
        [pscustomobject]@{ Indicateur = 'Version du script';                        Valeur = $ScriptVersion }
        [pscustomobject]@{ Indicateur = 'Référence Microsoft';                      Valeur = 'MC1448379 - fin du traitement memberOf le 03/11/2026' }
        [pscustomobject]@{ Indicateur = 'Jours restants avant échéance';            Valeur = [math]::Max(0, [math]::Ceiling(($RetirementDate - $started).TotalDays)) }
        [pscustomobject]@{ Indicateur = 'Groupes dynamiques (total)';               Valeur = $dynGroups.Count }
        [pscustomobject]@{ Indicateur = 'Groupes dynamiques avec memberOf';         Valeur = $impactedGroups.Count }
        [pscustomobject]@{ Indicateur = 'AU dynamiques (total)';                    Valeur = $dynAUs.Count }
        [pscustomobject]@{ Indicateur = 'AU dynamiques avec memberOf';              Valeur = $impactedAUs.Count }
        [pscustomobject]@{ Indicateur = 'Stratégies EM avec memberOf';              Valeur = $impactedEm.Count }
        [pscustomobject]@{ Indicateur = 'Groupes sources référencés';               Valeur = @($sourceRows).Count }
        [pscustomobject]@{ Indicateur = 'Dépendances totales';                      Valeur = $script:Dependencies.Count }
        [pscustomobject]@{ Indicateur = 'Dépendances critiques';                    Valeur = $criticalCount }
        [pscustomobject]@{ Indicateur = 'Intune analysé';                           Valeur = $(if ($SkipIntune) { 'Non' } else { 'Oui' }) }
        [pscustomobject]@{ Indicateur = 'Dépendances des groupes sources incluses'; Valeur = $(if ($IncludeSourceGroupDependencies) { 'Oui' } else { 'Non' }) }
        [pscustomobject]@{ Indicateur = 'Erreurs de collecte';                      Valeur = $script:Errors.Count }
        [pscustomobject]@{ Indicateur = 'Carte HTML';                               Valeur = $(if ($NoMap) { 'Non générée' } else { $MapPath }) }
    ) + @($catSummary)

    $ct = @(
        New-ConditionalText -Text 'Critique' -BackgroundColor LightPink   -ConditionalTextColor DarkRed
        New-ConditionalText -Text 'Élevée'   -BackgroundColor LightYellow -ConditionalTextColor DarkGoldenrod
    )
    if (Test-Path $OutputPath) { Remove-Item $OutputPath -Force }
    Export-Sheet -Data $summary       -Name 'Synthèse'
    Export-Sheet -Data $viewRows      -Name 'Vue par objet'    -ConditionalText $ct
    Export-Sheet -Data $groupRows     -Name 'Groupes memberOf' -ConditionalText $ct
    Export-Sheet -Data $auRows        -Name 'AU memberOf'      -ConditionalText $ct
    Export-Sheet -Data $emRows        -Name 'EM memberOf'
    Export-Sheet -Data $sourceRows    -Name 'Groupes sources'
    Export-Sheet -Data $depRows       -Name 'Dépendances'      -ConditionalText $ct
    Export-Sheet -Data $script:Errors -Name 'Erreurs collecte'

    if (-not $NoMap) {
        $mapData = [pscustomobject]@{
            meta = [pscustomobject]@{
                tenant   = $tenantName
                date     = $started.ToString('dd/MM/yyyy HH:mm')
                daysLeft = [math]::Max(0, [math]::Ceiling(($RetirementDate - $started).TotalDays))
                version  = $ScriptVersion
                groups   = $impactedGroups.Count
                aus      = $impactedAUs.Count
                em       = $impactedEm.Count
                deps     = $script:Dependencies.Count
                critical = $criticalCount
                errors   = $script:Errors.Count
                families = $FamilyOrder
            }
            objects = @($mapObjects)
        }
        New-DependencyMap -Path $MapPath -Data $mapData
    }

    # Bilan console
    $elapsed = (Get-Date) - $started
    Write-Host ''
    Write-Host ('=' * 78) -ForegroundColor DarkCyan
    Write-Host '  AUDIT TERMINÉ' -ForegroundColor White
    Write-Host ('=' * 78) -ForegroundColor DarkCyan
    Write-Check -Label 'Groupes dynamiques avec memberOf' -State $(if ($impactedGroups.Count) { 'ATTENTION' } else { 'OK' }) -Detail $impactedGroups.Count
    Write-Check -Label 'AU dynamiques avec memberOf'      -State $(if ($impactedAUs.Count)    { 'ATTENTION' } else { 'OK' }) -Detail $impactedAUs.Count
    Write-Check -Label 'Stratégies EM avec memberOf'      -State $(if ($impactedEm.Count)     { 'ATTENTION' } else { 'OK' }) -Detail $impactedEm.Count
    Write-Check -Label 'Dépendances (dont critiques)'     -State INFO -Detail ("{0} ({1})" -f $script:Dependencies.Count, $criticalCount)
    Write-Check -Label 'Erreurs de collecte'              -State $(if ($script:Errors.Count) { 'ATTENTION' } else { 'OK' }) -Detail $script:Errors.Count
    Write-Check -Label 'Durée'                            -State INFO -Detail ('{0:mm\:ss}' -f $elapsed)
    Write-Check -Label 'Rapport Excel'                    -State INFO -Detail $OutputPath
    if (-not $NoMap) { Write-Check -Label 'Carte HTML'    -State INFO -Detail $MapPath }
    Write-Host ''
    if ($script:Errors.Count -gt 0) {
        Write-Warning 'Des collectes ont échoué (souvent un droit ou une licence manquante) : voir l''onglet Erreurs collecte.'
    }
}
finally {
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    if ($transcript) { Stop-Transcript | Out-Null }
}
