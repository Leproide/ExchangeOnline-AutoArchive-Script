#Requires -Version 5.1
<#
.GitHub
https://github.com/Leproide/ExchangeOnline-AutoArchive-Script

.LICENSE
GPL v3 https://www.gnu.org/licenses/gpl-3.0.html

.SYNOPSIS
    Abilita l'archivio online Exchange, imposta la retention policy e avvia l'archiviazione per un utente.

.DESCRIPTION
    Script interattivo che:
    - Verifica la presenza del modulo ExchangeOnlineManagement
    - Esegue il login con le credenziali postmaster
    - Chiede l'utente target
    - Chiede la soglia di archiviazione (1-6 anni o personalizzato)
    - Mostra un riepilogo e chiede conferma prima di procedere

.NOTES
    Richiede: ExchangeOnlineManagement PowerShell module
    Ruoli minimi richiesti: Mail Recipients, Retention Management
    Documentazione: https://learn.microsoft.com/en-us/purview/enable-archive-mailboxes
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ---------------------------------------------------------------------------
# FUNZIONI DI UTILITA'
# ---------------------------------------------------------------------------

function Write-Header {
    param([string]$Text)
    $line = "=" * 60
    Write-Host ""
    Write-Host $line -ForegroundColor Cyan
    Write-Host "  $Text" -ForegroundColor Cyan
    Write-Host $line -ForegroundColor Cyan
    Write-Host ""
}

function Write-Step {
    param([string]$Text)
    Write-Host "[*] $Text" -ForegroundColor Yellow
}

function Write-OK {
    param([string]$Text)
    Write-Host "[OK] $Text" -ForegroundColor Green
}

function Write-Fail {
    param([string]$Text)
    Write-Host "[ERRORE] $Text" -ForegroundColor Red
}

function Write-Info {
    param([string]$Text)
    Write-Host "     $Text" -ForegroundColor Gray
}

# ---------------------------------------------------------------------------
# STEP 1 — Verifica modulo ExchangeOnlineManagement
# ---------------------------------------------------------------------------

Write-Header "Exchange Online - Attivazione Archivio Utente"

Write-Step "Verifica modulo ExchangeOnlineManagement..."

$module = Get-Module -ListAvailable -Name ExchangeOnlineManagement | Sort-Object Version -Descending | Select-Object -First 1

if (-not $module) {
    Write-Fail "Il modulo ExchangeOnlineManagement NON e' installato."
    Write-Host ""
    Write-Host "  Per installarlo, apri PowerShell come Amministratore ed esegui:" -ForegroundColor White
    Write-Host "  Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser" -ForegroundColor Cyan
    Write-Host ""
    exit 1
}

Write-OK "Modulo trovato: ExchangeOnlineManagement v$($module.Version)"

# ---------------------------------------------------------------------------
# STEP 2 — Login con credenziali postmaster
# ---------------------------------------------------------------------------

Write-Header "Connessione a Exchange Online"

Write-Step "Inserisci le credenziali dell'account postmaster / amministratore Exchange."
Write-Info "Verra' aperta la finestra di autenticazione Microsoft."
Write-Host ""

try {
    # Connect-ExchangeOnline apre il browser/popup OAuth2 se non passi -Credential
    Connect-ExchangeOnline -ShowBanner:$false
    Write-OK "Connessione a Exchange Online stabilita."
} catch {
    Write-Fail "Connessione fallita: $_"
    exit 1
}

# ---------------------------------------------------------------------------
# Abilitazione Auto-Expanding Archive a livello tenant
# ---------------------------------------------------------------------------

Write-Step "Abilitazione Auto-Expanding Archive a livello organizzazione..."
try {
    Set-OrganizationConfig -AutoExpandingArchive -ErrorAction Stop
    Write-OK "Auto-Expanding Archive abilitato a livello organizzazione."
} catch {
    Write-Fail "Impossibile abilitare Auto-Expanding Archive a livello organizzazione: $_"
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
    exit 1
}

Write-Step "Verifica Auto-Expanding Archive a livello organizzazione..."
$orgCfg = Get-OrganizationConfig
Write-OK "AutoExpandingArchiveEnabled = $($orgCfg.AutoExpandingArchiveEnabled)"

# ---------------------------------------------------------------------------
# STEP 3 — Richiesta utente target
# ---------------------------------------------------------------------------

Write-Header "Selezione Utente"

do {
    $userInput = Read-Host "Inserisci l'UPN o il nome display dell'utente (es. mario.rossi@contoso.com)"
    $userInput = $userInput.Trim()

    if ([string]::IsNullOrWhiteSpace($userInput)) {
        Write-Fail "L'indirizzo non puo' essere vuoto. Riprova."
        $mailbox = $null
        continue
    }

    Write-Step "Ricerca cassetta postale per '$userInput'..."

    try {
        $mailbox = Get-Mailbox -Identity $userInput -ErrorAction Stop
    } catch {
        Write-Fail "Utente non trovato: $_"
        $mailbox = $null
    }

    if ($mailbox) {
        Write-OK "Trovato: $($mailbox.DisplayName) <$($mailbox.PrimarySmtpAddress)>"
    }

} while (-not $mailbox)

# Controlla se l'archivio e' gia' attivo
$archiveAlreadyEnabled = ($mailbox.ArchiveStatus -eq "Active")

if ($archiveAlreadyEnabled) {
    Write-Host ""
    Write-Host "  [i] L'archivio online e' GIA' attivo per questo utente." -ForegroundColor Magenta
    Write-Host "      Verra' aggiornata solo la retention policy e avviata la sincronizzazione." -ForegroundColor Magenta
}

# ---------------------------------------------------------------------------
# STEP 4 — Scelta soglia di archiviazione
# ---------------------------------------------------------------------------

Write-Header "Configurazione Retention (Soglia di Archiviazione)"

Write-Host "  Seleziona dopo quanti anni le email vengono spostate nell'archivio:" -ForegroundColor White
Write-Host ""
Write-Host "    [1]  1 anno  (365 giorni)" -ForegroundColor White
Write-Host "    [2]  2 anni  (730 giorni)  -- Default MRM Policy" -ForegroundColor White
Write-Host "    [3]  3 anni  (1095 giorni)" -ForegroundColor White
Write-Host "    [4]  4 anni   (1460 giorni)" -ForegroundColor White
Write-Host "    [5]  5 anni   (1825 giorni)" -ForegroundColor White
Write-Host "    [6]  6 anni   (2190 giorni)" -ForegroundColor White
Write-Host "    [7]  Personalizzato (inserisci tu gli anni)" -ForegroundColor White
Write-Host ""

do {
    $choice = Read-Host "Scelta (1-7)"
    $retentionDays = 0
    switch ($choice.Trim()) {
        "1" { $retentionYears = 1 ; $retentionDays = 365  ; break }
        "2" { $retentionYears = 2 ; $retentionDays = 730  ; break }
        "3" { $retentionYears = 3 ; $retentionDays = 1095 ; break }
        "4" { $retentionYears = 4 ; $retentionDays = 1460 ; break }
        "5" { $retentionYears = 5 ; $retentionDays = 1825 ; break }
        "6" { $retentionYears = 6 ; $retentionDays = 2190 ; break }
        "7" {
            do {
                $customInput = Read-Host "  Inserisci il numero di anni (es. 7, 10, 15...)"
                if ($customInput -match '^\d+$' -and [int]$customInput -ge 1 -and [int]$customInput -le 99) {
                    $retentionYears = [int]$customInput
                    $retentionDays  = $retentionYears * 365
                } else {
                    Write-Fail "Valore non valido. Inserisci un numero intero tra 1 e 99."
                    $retentionDays = 0
                }
            } while ($retentionDays -eq 0)
            break
        }
        default {
            Write-Fail "Scelta non valida. Digita un numero da 1 a 7."
            $retentionDays = 0
        }
    }
} while ($retentionDays -eq 0)

Write-OK "Soglia selezionata: $retentionYears anno/i ($retentionDays giorni)"

# Nomi tag/policy che verra' creata (o riutilizzata se esiste gia')
$tagName    = "Archivia dopo $retentionYears anno"
if ($retentionYears -gt 1) { $tagName = "Archivia dopo $retentionYears anni" }
$policyName = "Policy-Archivio-$retentionYears-Anno"
if ($retentionYears -gt 1) { $policyName = "Policy-Archivio-$retentionYears-Anni" }

# ---------------------------------------------------------------------------
# STEP 5 — Riepilogo e conferma
# ---------------------------------------------------------------------------

Write-Header "Riepilogo Operazioni"

Write-Host "  Utente              : $($mailbox.DisplayName)" -ForegroundColor White
Write-Host "  UPN                 : $($mailbox.PrimarySmtpAddress)" -ForegroundColor White
Write-Host "  Archivio online     : $(if ($archiveAlreadyEnabled) { 'Gia'' attivo (nessuna modifica)' } else { 'Da ATTIVARE' })" -ForegroundColor White
Write-Host "  Retention tag       : '$tagName'" -ForegroundColor White
Write-Host "  Retention policy    : '$policyName'" -ForegroundColor White
Write-Host "  Soglia archiviazione: email piu' vecchie di $retentionYears anno/i" -ForegroundColor White
Write-Host "  Azione              : MoveToArchive (sposta in archivio, non elimina)" -ForegroundColor White
Write-Host "  Avvio immediato     : Start-ManagedFolderAssistant (forzato)" -ForegroundColor White
Write-Host ""
Write-Host "  NOTA: Se il retention tag o la policy esistono gia' vengono riutilizzati." -ForegroundColor DarkYellow
Write-Host ""

$confirm = Read-Host "Procedere? (S per confermare / N per annullare)"

if ($confirm.Trim().ToUpper() -ne "S") {
    Write-Host ""
    Write-Host "  Operazione annullata dall'utente." -ForegroundColor Magenta
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
    exit 0
}

# ---------------------------------------------------------------------------
# STEP 6 — Esecuzione
# ---------------------------------------------------------------------------

Write-Header "Esecuzione"

# 6a. Abilita archivio se non attivo
if (-not $archiveAlreadyEnabled) {
    Write-Step "Abilitazione archivio online..."
    try {
        Enable-Mailbox -Identity $mailbox.PrimarySmtpAddress -Archive -ErrorAction Stop | Out-Null
		Set-Mailbox -Identity $mailbox.PrimarySmtpAddress -ArchiveName "Archivio Online - $($mailbox.PrimarySmtpAddress)" -ErrorAction Stop | Out-Null
        Write-OK "Archivio online abilitato."
    } catch {
        Write-Fail "Impossibile abilitare l'archivio: $_"
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
        exit 1
    }
} else {
    Write-Info "Archivio online gia' attivo, step saltato."
}

# 6a-bis. Abilita Auto-Expanding sulla mailbox
Write-Step "Abilitazione Auto-Expanding Archive sulla mailbox..."
try {
    Enable-Mailbox -Identity $mailbox.PrimarySmtpAddress -AutoExpandingArchive -ErrorAction Stop | Out-Null
    Write-OK "Auto-Expanding Archive abilitato per la mailbox."
} catch {
    Write-Fail "Impossibile abilitare Auto-Expanding Archive sulla mailbox: $_"
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
    exit 1
}

# 6b. Crea il retention tag (se non esiste)
Write-Step "Verifica/creazione retention tag '$tagName'..."
try {
    $existingTag = Get-RetentionPolicyTag -Identity $tagName -ErrorAction SilentlyContinue
    if ($existingTag) {
        Write-Info "Tag gia' esistente, riutilizzato."
    } else {
        New-RetentionPolicyTag `
            -Name            $tagName `
            -Type            All `
            -RetentionEnabled $true `
            -AgeLimitForRetention $retentionDays `
            -RetentionAction MoveToArchive `
            -ErrorAction Stop | Out-Null
        Write-OK "Retention tag creato."
    }
} catch {
    Write-Fail "Errore nella gestione del retention tag: $_"
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
    exit 1
}

# 6c. Crea la retention policy (se non esiste) e aggancia il tag
Write-Step "Verifica/creazione retention policy '$policyName'..."
try {
    $existingPolicy = Get-RetentionPolicy -Identity $policyName -ErrorAction SilentlyContinue
    if ($existingPolicy) {
        Write-Info "Policy gia' esistente, riutilizzata."
        # Assicura che il tag sia presente nella policy
        $tagLinks = $existingPolicy.RetentionPolicyTagLinks
        if ($tagLinks -notcontains $tagName) {
            Set-RetentionPolicy -Identity $policyName -RetentionPolicyTagLinks ($tagLinks + $tagName) -ErrorAction Stop | Out-Null
            Write-Info "Tag aggiunto alla policy esistente."
        }
    } else {
        New-RetentionPolicy `
            -Name                   $policyName `
            -RetentionPolicyTagLinks $tagName `
            -ErrorAction Stop | Out-Null
        Write-OK "Retention policy creata."
    }
} catch {
    Write-Fail "Errore nella gestione della retention policy: $_"
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
    exit 1
}

# 6d. Assegna la policy alla cassetta postale
Write-Step "Assegnazione della policy '$policyName' alla cassetta postale..."
try {
    Set-Mailbox -Identity $mailbox.PrimarySmtpAddress -RetentionPolicy $policyName -ErrorAction Stop
    Write-OK "Policy assegnata."
} catch {
    Write-Fail "Errore nell'assegnazione della policy: $_"
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
    exit 1
}

# 6e. Forza l'avvio immediato del Managed Folder Assistant
Write-Step "Avvio immediato del Managed Folder Assistant (archiviazione forzata)..."
try {
    Start-ManagedFolderAssistant -Identity $mailbox.PrimarySmtpAddress -ErrorAction Stop
    Write-OK "Managed Folder Assistant avviato."
    Write-Info "L'archiviazione partira' entro pochi minuti."
    Write-Info "Per monitorare il progresso, controlla la cartella 'Archivio In-Place' in OWA."
} catch {
    Write-Fail "Errore nell'avvio del Managed Folder Assistant: $_"
    Write-Info "La policy e' comunque stata applicata. L'archiviazione avverra' automaticamente entro 7 giorni."
}


# ---------------------------------------------------------------------------
# STEP 7 — Verifica finale e riepilogo
# ---------------------------------------------------------------------------

Write-Header "Verifica Finale"

Write-Step "Recupero stato aggiornato della cassetta postale..."
try {
    $updatedMailbox = Get-Mailbox -Identity $mailbox.PrimarySmtpAddress -ErrorAction Stop

    Write-Host ""
    Write-Host "  --- Stato Finale ---" -ForegroundColor Cyan
    Write-Host "  Utente          : $($updatedMailbox.DisplayName)" -ForegroundColor White
    Write-Host "  UPN             : $($updatedMailbox.PrimarySmtpAddress)" -ForegroundColor White
    Write-Host "  Stato archivio  : $($updatedMailbox.ArchiveStatus)" -ForegroundColor White
    Write-Host "  Nome archivio   : $($updatedMailbox.ArchiveName)" -ForegroundColor White
    Write-Host "  Retention policy: $($updatedMailbox.RetentionPolicy)" -ForegroundColor White
    Write-Host ""
} catch {
    Write-Info "Impossibile recuperare il riepilogo finale: $_"
}

Write-Host "  Operazione completata con successo!" -ForegroundColor Green
Write-Host ""

# ---------------------------------------------------------------------------
# Disconnessione
# ---------------------------------------------------------------------------

Write-Step "Disconnessione da Exchange Online..."
Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
Write-OK "Sessione chiusa."
Write-Host ""
