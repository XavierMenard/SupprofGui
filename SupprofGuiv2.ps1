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

# === Liste des cours pour Réseau ===
$coursPattern = '<option value="([^"]+)">\s*([^<]+)\s*</option>'
$coursMatches = [regex]::Matches($response.Content, $coursPattern)

# Liste des cours pour Réseau
$coursListeRéseau = @{
    "SYSTEME EXPLOITATION VIEILLISANT" = 60
    "SYSTEME EXPLOITATION RECENT" = 50
    "RESEAUX GESTION DACCES" = 30
    "RÉSEAUX: PARTAGE DES RESSOURCES" = 35
    "RÉSEAUX LOCAUX" = 48
    "VIRTUALISATION" = 30
    "SERVEUR ET GESTION D'ACCES" = 36
}

# === Fonction pour mettre à jour les blocs selon le cours sélectionné ===
function UpdateBlocList($coursNom) {
    # Déterminer le nombre de blocs selon le cours sélectionné
    $nombreDeBlocs = $coursListeRéseau[$coursNom]

    # Vérifier que nombreDeBlocs est bien un entier
    if (-not ($nombreDeBlocs -as [int])) {
        $nombreDeBlocs = 35 # Si ce n'est pas un entier, on met une valeur par défaut
    }

    # Vider la liste des blocs
    $comboBloc.Items.Clear()

    # Ajouter les blocs dans le ComboBox
    1..$nombreDeBlocs | ForEach-Object { 
        $comboBloc.Items.Add($_.ToString())
    }

    # Mettre à jour le label pour afficher la plage de blocs
    $lblBloc.Text = "Bloc (1-$nombreDeBlocs) :"
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

# === Fonction pour ajouter un Label stylé ===
function Add-Label($text, $x, $y) {
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $text
    $lbl.Location = New-Object System.Drawing.Point($x, $y)
    $lbl.Size = New-Object System.Drawing.Size(450, 22)
    $lbl.ForeColor = [System.Drawing.Color]::FromArgb(45, 45, 45)
    $lbl.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    return $lbl
}

# === Fonction pour ajouter un ComboBox stylé ===
function Add-ComboBox($x, $y, $width) {
    $cb = New-Object System.Windows.Forms.ComboBox
    $cb.Location = New-Object System.Drawing.Point($x, $y)
    $cb.Size = New-Object System.Drawing.Size($width, 30)
    $cb.DropDownStyle = 'DropDownList'
    $cb.FlatStyle = 'Flat'
    return $cb
}

# === Onglets ===
$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Size = New-Object System.Drawing.Size(500, 500)
$tabControl.Location = New-Object System.Drawing.Point(10, 70)

# === Onglet Réseau ===
$tabRéseau = New-Object System.Windows.Forms.TabPage
$tabRéseau.Text = "Réseau"
$tabControl.TabPages.Add($tabRéseau)

# === Ajouter TabControl à la fenêtre ===
$form.Controls.Add($tabControl)

# === Interface Réseau ===
$tabRéseau.Controls.Add((Add-Label "Sélectionner le cours réseau :" 20 20))
$comboCoursRéseau = Add-ComboBox 20 45 460
$coursListeRéseau.Keys | ForEach-Object { $comboCoursRéseau.Items.Add($_) }
$tabRéseau.Controls.Add($comboCoursRéseau)

$lblBloc = Add-Label "Bloc (1-70) :" 20 80
$tabRéseau.Controls.Add($lblBloc)

$comboBloc = Add-ComboBox 20 105 150
$tabRéseau.Controls.Add($comboBloc)

$tabRéseau.Controls.Add((Add-Label "Type de demande :" 20 140))
$comboType = Add-ComboBox 20 165 200
$comboType.Items.AddRange(@("Validation", "Explication"))
$tabRéseau.Controls.Add($comboType)

$tabRéseau.Controls.Add((Add-Label "Numéro de local :" 20 210))
$comboLocal = Add-ComboBox 20 235 200
$comboLocal.Items.AddRange(@("A-121")) # Remettre A-121 ici pour Réseau
$comboLocal.SelectedIndex = 0
$tabRéseau.Controls.Add($comboLocal)

$tabRéseau.Controls.Add((Add-Label "Poste (1-75 + Local des serveurs) :" 20 290))
$comboPoste = Add-ComboBox 20 315 200
1..75 | ForEach-Object { $comboPoste.Items.Add("Poste $_") }
$comboPoste.Items.Add("Local des serveurs")
$tabRéseau.Controls.Add($comboPoste)

$btnSoumettre = New-Object System.Windows.Forms.Button
$btnSoumettre.Text = "Soumettre la demande Réseau"
$btnSoumettre.Location = New-Object System.Drawing.Point(20, 385)
$btnSoumettre.Size = New-Object System.Drawing.Size(220,40)
$btnSoumettre.BackColor = [System.Drawing.Color]::FromArgb(45, 125, 245)
$btnSoumettre.ForeColor = "White"
$btnSoumettre.FlatStyle = "Flat"
$tabRéseau.Controls.Add($btnSoumettre)

# === Mise à jour dynamique des blocs lorsque le cours change ===
$comboCoursRéseau.Add_SelectedIndexChanged({
    UpdateBlocList($comboCoursRéseau.SelectedItem)
})

# === Affichage ===
$form.ShowDialog()
