# --- CONFIGURACIÓN ---
# $rutaRaiz = "C:\Ruta\A\Los\PSTs"
# $archivoSalida = "C:\Ruta\De\Salida\emails.csv"

# --- INICIO ---
Write-Host "Iniciando modo FUERZA BRUTA para emails..." -ForegroundColor Cyan

# Iniciar Outlook
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
}
catch {
    Write-Error "Error: No se pudo iniciar Outlook."
    exit
}

$resultados = @()
$idsProcesados = New-Object System.Collections.Generic.HashSet[string]
$contadorEmails = 0

# CÓDIGOS MAPI
$PR_INTERNET_MESSAGE_ID = "http://schemas.microsoft.com/mapi/proptag/0x1035001F"
$PR_SENDER_SMTP         = "http://schemas.microsoft.com/mapi/proptag/0x5D01001F"
$PR_SENT_REPRESENTING   = "http://schemas.microsoft.com/mapi/proptag/0x5D02001F" # "Enviado en nombre de"
$PR_SMTP_ADDRESS        = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"

# --- FUNCIÓN AUXILIAR: Extraer dominio del nombre del archivo ---
function Obtener-DominioDelPST($nombreArchivo) {
    # Busca patrones tipo @dominio.com en el nombre del archivo
    if ($nombreArchivo -match "@([\w\.-]+\.[a-zA-Z]{2,4})") {
        return $matches[1]
    }
    # Dominio por defecto si no se detecta nada (puedes cambiarlo)
    return "ejemplo.com" 
}

# --- FUNCIÓN: RECONSTRUIR EMAIL (La clave del forzado) ---
function Forzar-Email($rawAddress, $dominioFallback) {
    # Si ya es un email, devolverlo
    if ($rawAddress -match "@" -and $rawAddress -notmatch "^/") {
        return $rawAddress
    }
    
    # Si es una ruta Legacy Exchange (/o=.../cn=usuario)
    if ($rawAddress -match "cn=([^/]+)$") {
        # Extraemos lo que hay después del último cn=
        $usuario = $matches[1]
        # Limpiamos basura común que a veces queda (ej: usuario-001)
        $usuario = $usuario -replace "-\d+$", "" 
        
        # Construimos el email
        return "$usuario@$dominioFallback".ToLower()
    }
    
    return ""
}

# --- FUNCIÓN: OBTENER REMITENTE (FROM) FORZADO ---
function Obtener-SenderForce($item, $dominioPST) {
    $senderFinal = ""

    # 1. Intentar propiedades SMTP directas (MAPI)
    try { $senderFinal = $item.PropertyAccessor.GetProperty($PR_SENDER_SMTP) } catch { }
    if ([string]::IsNullOrEmpty($senderFinal)) {
        try { $senderFinal = $item.PropertyAccessor.GetProperty($PR_SENT_REPRESENTING) } catch { }
    }

    # 2. Si falla, intentar resolver ExchangeUser
    if ([string]::IsNullOrEmpty($senderFinal) -or $senderFinal -notmatch "@") {
        try {
            if ($item.SenderEmailType -eq "EX") {
                $exUser = $item.Sender.GetExchangeUser()
                if ($exUser) { $senderFinal = $exUser.PrimarySmtpAddress }
            }
        } catch { }
    }

    # 3. SI FALLA TODO: RECONSTRUCCIÓN (La parte nueva)
    if ([string]::IsNullOrEmpty($senderFinal) -or $senderFinal -notmatch "@") {
        try {
            $raw = $item.SenderEmailAddress
            # Intentamos reconstruir usando el string interno y el dominio del PST
            $reconstruido = Forzar-Email $raw $dominioPST
            
            if (-not [string]::IsNullOrEmpty($reconstruido)) {
                $senderFinal = $reconstruido
            } else {
                # 4. Último recurso desesperado: Buscar email en el Nombre
                # A veces el nombre es "Perez, Juan (juan.perez@dominio.com)"
                if ($item.SenderName -match "([\w\.-]+@[\w\.-]+\.[a-zA-Z]{2,4})") {
                    $senderFinal = $matches[1]
                }
            }
        } catch { }
    }

    # Limpieza final
    return $senderFinal.ToString().Trim()
}

# --- FUNCIÓN: OBTENER DESTINATARIOS (TO) ---
function Obtener-RecipientsLimpio($mailItem, $dominioPST) {
    $listaEmails = @()
    foreach ($recip in $mailItem.Recipients) {
        $email = $null
        try {
            if ($recip.Type -eq 1 -or $recip.Address -match "@") { 
                $email = $recip.Address 
            }
            if ([string]::IsNullOrEmpty($email) -or $email -notmatch "@") {
                try { $email = $recip.PropertyAccessor.GetProperty($PR_SMTP_ADDRESS) } catch {}
                if ([string]::IsNullOrEmpty($email)) {
                    try {
                        $user = $recip.AddressEntry.GetExchangeUser()
                        if ($user) { $email = $user.PrimarySmtpAddress }
                    } catch {}
                }
            }
            
            # Si sigue sin ser email, intentamos forzar reconstrucción también para destinatarios
            if ([string]::IsNullOrEmpty($email) -or $email -notmatch "@") {
                $email = Forzar-Email $recip.Address $dominioPST
            }

            if ($email -match "@" -and $email -notmatch "^/") {
                $listaEmails += $email.ToLower().Trim()
            }
        } catch {}
    }
    return ($listaEmails -join "; ")
}

# --- PROCESO RECURSIVO ---
function Procesar-Carpeta($carpeta, $nombreArchivoPST, $dominioActual) {
    if ($carpeta.Items.Count -gt 0) {
        foreach ($item in $carpeta.Items) {
            if ($item -is [Microsoft.Office.Interop.Outlook.MailItem]) {
                try {
                    # ID
                    $msgId = $null
                    try { $msgId = $item.PropertyAccessor.GetProperty($PR_INTERNET_MESSAGE_ID) } catch { }
                    if ([string]::IsNullOrWhiteSpace($msgId)) { continue }

                    # Desduplicar
                    if (-not $idsProcesados.Contains($msgId)) {
                        $null = $idsProcesados.Add($msgId)
                        $contadorEmails++

                        # DATOS
                        $fromData = Obtener-SenderForce $item $dominioActual
                        $toData   = Obtener-RecipientsLimpio $item $dominioActual
                        
                        # Si aun forzando sale vacío (ej: sistema), poner valor por defecto
                        if ([string]::IsNullOrEmpty($fromData)) { $fromData = "unknown@no-email-found" }

                        $fecha = $item.SentOn
                        $fechaUTC = if ($fecha) { $fecha.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss") } else { "" }

                        $obj = [PSCustomObject]@{
                            'MessageID' = $msgId
                            'From'      = $fromData
                            'To'        = $toData
                            'DateUTC'   = $fechaUTC
                            'SourcePST' = $nombreArchivoPST
                        }
                        
                        $script:resultados += $obj

                        if ($contadorEmails % 100 -eq 0) {
                            Write-Host "   -> $contadorEmails (Último From: $fromData)" -ForegroundColor Gray
                        }
                    }
                }
                catch { }
            }
        }
    }
    foreach ($sub in $carpeta.Folders) { Procesar-Carpeta $sub $nombreArchivoPST $dominioActual }
}

# --- BUCLE PRINCIPAL ---
$archivosPst = Get-ChildItem -Path $rutaRaiz -Filter *.pst -Recurse

foreach ($archivo in $archivosPst) {
    Write-Host "Procesando: $($archivo.Name)" -ForegroundColor Yellow
    
    # DEDUCIR DOMINIO DE ESTE PST
    $dominioEstePST = Obtener-DominioDelPST $archivo.Name
    Write-Host "   -> Dominio base detectado: $dominioEstePST" -ForegroundColor DarkGray

    try {
        $namespace.AddStore($archivo.FullName)
        $store = $namespace.Stores | Where-Object { $_.FilePath -eq $archivo.FullName } | Select-Object -First 1
        if ($store) {
            $raiz = $store.GetRootFolder()
            Procesar-Carpeta $raiz $archivo.Name $dominioEstePST
            $namespace.RemoveStore($raiz)
        }
    } catch { Write-Host "   Error: $($_.Exception.Message)" -ForegroundColor Red }
}

# --- EXPORTAR ---
Write-Host "Guardando CSV..." -ForegroundColor Green
$resultados | Export-Csv -Path $archivoSalida -NoTypeInformation -Delimiter ";" -Encoding UTF8
Write-Host "¡TERMINADO! Verifica: $archivoSalida" -ForegroundColor Cyan