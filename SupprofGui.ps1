# Masquer la fenêtre PowerShell
$psWindow = Get-Process -Id $PID
$psWindow.MainWindowHandle | ForEach-Object { 
    Add-Type -TypeDefinition @"
    using System;
    using System.Runtime.InteropServices;
    public class ShowWindow {
        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);
        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
    }
"@
    [ShowWindow]::ShowWindowAsync($psWindow.MainWindowHandle, 0) # 0 signifie "Hide"
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# === Session avec cookies ===
$session = New-Object Microsoft.PowerShell.Commands.WebRequestSession
Invoke-WebRequest -Uri "http://supprof.cfpmr.com/supprof/etudiants/index.php?lstPlateau=SER&Niveau=etudiant" `
  -WebSession $session -Headers @{ "User-Agent" = "Mozilla/5.0" } | Out-Null
$response = Invoke-WebRequest -Uri "http://supprof.cfpmr.com/supprof/etudiants/index.php?selItem=ajouter_demande" `
  -WebSession $session -Headers @{ "User-Agent" = "Mozilla/5.0" }

# === Cours filtrés ===
$coursPattern = '<option value="([^"]+)">\s*([^<]+)\s*</option>'
$coursMatches = [regex]::Matches($response.Content, $coursPattern)
$coursListe = @{}
foreach ($match in $coursMatches) {
    $id = $match.Groups[1].Value.Trim()
    $nom = $match.Groups[2].Value.Trim()
    if ($nom -notmatch "(?i)poste|enseignant|explication|validation|local|A-121") {
        $coursListe[$nom] = $id
    }
}

# === Fenêtre principale ===
$form = New-Object System.Windows.Forms.Form
$form.Text = "Demande Supprof"
$form.Size = New-Object System.Drawing.Size(520, 600)
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::FromArgb(250, 250, 255)
$form.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox = $false

# === Bannière moderne ===
$banner = New-Object System.Windows.Forms.Label
$banner.Text = "Bienvenue Xavier Ménard"
$banner.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
$banner.ForeColor = "White"
$banner.BackColor = [System.Drawing.Color]::FromArgb(45, 125, 245)
$banner.TextAlign = "MiddleCenter"
$banner.Size = New-Object System.Drawing.Size(500, 50)
$banner.Location = New-Object System.Drawing.Point(10, 10)
$form.Controls.Add($banner)

# === Fonction : Label stylé ===
function Add-Label($text, $x, $y) {
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $text
    $lbl.Location = New-Object System.Drawing.Point($x, $y)
    $lbl.Size = New-Object System.Drawing.Size(450, 22)
    $lbl.ForeColor = [System.Drawing.Color]::FromArgb(45, 45, 45)
    $lbl.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    return $lbl
}

# === Fonction : ComboBox stylé ===
function Add-ComboBox($x, $y, $width) {
    $cb = New-Object System.Windows.Forms.ComboBox
    $cb.Location = New-Object System.Drawing.Point($x, $y)
    $cb.Size = New-Object System.Drawing.Size($width, 30)
    $cb.DropDownStyle = 'DropDownList'
    $cb.FlatStyle = 'Flat'
    return $cb
}

# === Interface utilisateur moderne ===
$form.Controls.Add((Add-Label "Sélectionner le cours :" 20 80))
$comboCours = Add-ComboBox 20 105 460
$coursListe.Keys | ForEach-Object { $comboCours.Items.Add($_) }
$form.Controls.Add($comboCours)

$form.Controls.Add((Add-Label "Bloc (1-70) :" 20 150))
$comboBloc = Add-ComboBox 20 175 150
1..70 | ForEach-Object { $comboBloc.Items.Add($_.ToString()) }
$form.Controls.Add($comboBloc)

$form.Controls.Add((Add-Label "Type de demande :" 20 220))
$comboType = Add-ComboBox 20 245 200
$comboType.Items.AddRange(@("Validation", "Explication"))
$form.Controls.Add($comboType)

$form.Controls.Add((Add-Label "Numéro de local :" 20 290))
$comboLocal = Add-ComboBox 20 315 200
$comboLocal.Items.AddRange(@("A-121", "Poste d’enseignement"))
$comboLocal.SelectedIndex = 0
$form.Controls.Add($comboLocal)

$form.Controls.Add((Add-Label "Poste (1-75 + Local des serveurs) :" 20 360))
$comboPoste = Add-ComboBox 20 385 200
1..75 | ForEach-Object { $comboPoste.Items.Add("Poste $_") }
$comboPoste.Items.Add("Local des serveurs")
$form.Controls.Add($comboPoste)

# === Bouton stylé moderne ===
$btnSoumettre = New-Object System.Windows.Forms.Button
$btnSoumettre.Text = "Soumettre la demande"
$btnSoumettre.Location = New-Object System.Drawing.Point(20, 440)
$btnSoumettre.Size = New-Object System.Drawing.Size(220,40)
$btnSoumettre.BackColor = [System.Drawing.Color]::FromArgb(45, 125, 245)
$btnSoumettre.ForeColor = "White"
$btnSoumettre.FlatStyle = "Flat"
$btnSoumettre.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($btnSoumettre)

# === Effet survol bouton ===
$btnSoumettre.Add_MouseEnter({
    $btnSoumettre.BackColor = [System.Drawing.Color]::FromArgb(25, 105, 225)
})
$btnSoumettre.Add_MouseLeave({
    $btnSoumettre.BackColor = [System.Drawing.Color]::FromArgb(45, 125, 245)
})

# === Action bouton ===
$btnSoumettre.Add_Click({
    $coursNom = $comboCours.SelectedItem
    $blocNum = $comboBloc.SelectedItem
    $typeDemande = $comboType.SelectedItem
    $local = $comboLocal.SelectedItem
    $posteSelection = $comboPoste.SelectedItem

    if (-not $coursNom -or -not $blocNum -or -not $typeDemande -or -not $posteSelection) {
        [System.Windows.Forms.MessageBox]::Show("Veuillez remplir tous les champs.", "Erreur")
        return
    }

    $coursID = $coursListe[$coursNom]
    $blocValue = "$blocNum;https://fp.cssdd.gouv.qc.ca/course/view.php?id=152&section=$blocNum;2014"
    $posteValue = if ($posteSelection -eq "Local des serveurs") { "Local des serveurs" } else { ($posteSelection -replace "Poste ", "").Trim() }

    $body = @{
        "lstCours"       = $coursID
        "lstBloc"        = $blocValue
        "lstTypeDemande" = $typeDemande
        "lstLocal"       = $local
        "txtPoste"       = $posteValue
        "txtNoFiche"     = "3161735"
        "question"       = ""
        "cmdDemande"     = "Ajouter"
    }

    try {
        Invoke-WebRequest -Uri "http://supprof.cfpmr.com/supprof/etudiants/index.php?selItem=ajouter_demande_db" `
            -Method POST `
            -WebSession $session `
            -Body $body `
            -ContentType "application/x-www-form-urlencoded" `
            -Headers @{ "User-Agent" = "Mozilla/5.0" } | Out-Null

        [System.Windows.Forms.MessageBox]::Show("✅ Demande soumise avec succès pour :`n$coursNom | Bloc $blocNum | Poste $posteValue", "Succès")
    } catch {
        [System.Windows.Forms.MessageBox]::Show("❌ Erreur lors de l'envoi : $_", "Erreur")
    }
})

# === Affiche la fenêtre ===
$form.Topmost = $true
$form.Add_Shown({ $form.Activate() })
[void]$form.ShowDialog()
