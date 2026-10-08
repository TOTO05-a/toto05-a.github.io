$ErrorActionPreference = 'Stop'
$path = 'C:\Users\ac138\Downloads\portfolio (1)\CV-Eduardo-Cruz-Bolanos-Plantilla-Actualizada.docx'
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
    $doc = $word.Documents.Open($path, $false, $false)
    $table = $doc.Tables.Item(1)

    function Set-ParaText {
        param([int]$Index, [string]$Text, [int]$Alignment = 0)
        $paragraph = $doc.Paragraphs.Item($Index)
        $range = $paragraph.Range
        $oldText = $range.Text
        $suffixLength = 0
        if ($oldText.EndsWith(([string][char]13) + [char]7)) { $suffixLength = 2 }
        elseif ($oldText.EndsWith([string][char]13)) { $suffixLength = 1 }
        $range.SetRange($range.Start, $range.End - $suffixLength)
        $range.Text = $Text
        $paragraph.Range.Font.Name = 'Calibri'
        $paragraph.Range.ParagraphFormat.Alignment = $Alignment
    }

    function Move-Icon {
        param([string]$AltText, [int]$TargetRow)
        $shape = $null
        for ($i = 1; $i -le $doc.InlineShapes.Count; $i++) {
            if ($doc.InlineShapes.Item($i).AlternativeText -eq $AltText) {
                $shape = $doc.InlineShapes.Item($i)
                break
            }
        }
        if ($null -eq $shape) { throw "No se encontró el icono: $AltText" }
        $shape.Range.Cut()
        $target = $table.Cell($TargetRow, 1).Range
        $target.Collapse(1)
        $target.Paste()
    }

    function Set-ContactCell {
        param([int]$Row, [string]$Text)
        $cell = $table.Cell($Row, 2)
        $cell.Range.Text = $Text + [char]13 + [char]7
        $cell.Range.Font.Name = 'Calibri'
        $cell.Range.Font.Size = 8.5
        $cell.Range.Font.Color = 16777215
        $cell.Range.Font.Bold = 0
    }

    function Find-TextRange {
        param([string]$Text)
        $range = $doc.Content.Duplicate
        $find = $range.Find
        $find.ClearFormatting()
        $find.Text = $Text
        $find.Forward = $true
        $find.Wrap = 0
        if ($find.Execute()) { return $range }
        throw "No se encontró texto para formato: $Text"
    }

    function Style-Text {
        param([string]$Text, [double]$Size, [bool]$Bold, [int]$Color, [int]$Alignment, [bool]$Italic = $false)
        $range = Find-TextRange $Text
        $range.Font.Name = 'Calibri'
        $range.Font.Size = $Size
        $range.Font.Bold = if ($Bold) { -1 } else { 0 }
        $range.Font.Italic = if ($Italic) { -1 } else { 0 }
        $range.Font.Color = $Color
        $range.ParagraphFormat.Alignment = $Alignment
    }

    $paragraphs = @{
        6 = 'Desarrollador de software junior con formación en frontend, backend, APIs y bases de datos. He aplicado estos conocimientos en proyectos CLI, automatización con n8n y robótica. Busco integrarme a un equipo donde aportar soluciones y continuar creciendo profesionalmente.'
        9 = 'Frontend: HTML5 · CSS3 · JavaScript'
        10 = 'Backend: Python · JavaScript · Node.js · Express.js · APIs REST · POO · SOLID · Patrones de diseño'
        11 = 'Bases de datos: MongoDB · MongoDB Atlas · MongoDB Compass · MySQL · CRUD · Transacciones'
        12 = 'Automatización e IA: n8n · Telegram · Flujos de trabajo · Fundamentos de inteligencia artificial'
        13 = 'Robótica: C++ · Python · Arduino · Sistemas de control · Integración de componentes'
        14 = 'Herramientas: Git · GitHub · GitHub Projects · Docker (en aprendizaje) · Postman · npm · VS Code · Inquirer · Chalk · dotenv · Testing'
        17 = 'Gestor de Portafolio de Proyectos Freelance' + [char]11 + 'JavaScript · Node.js · MongoDB · Git · GitHub' + [char]11 + 'Aplicación CLI para gestionar clientes, propuestas, proyectos, contratos, entregables y finanzas. Arquitectura modular, POO, SOLID, validaciones y transacciones MongoDB.'
        18 = 'Inventario Gamer' + [char]11 + 'Node.js · JavaScript · MongoDB · Mongoose · Docker' + [char]11 + 'Aplicación CLI para administrar inventario tecnológico, con integración a MongoDB, estructura modular y variables de entorno.'
        19 = 'Invernadero Automatizado | Proyecto académico de bachillerato' + [char]11 + 'Robótica · Automatización' + [char]11 + 'Participé en una solución automatizada para un invernadero, integrando componentes y aplicando conocimientos de robótica.'
        20 = 'LOGROS DE ROBÓTICA Y AUTOMATIZACIÓN'
        21 = 'RoboFut · Competencia nacional de robótica Balam 2025' + [char]11 + 'Armé y programé el robot desde cero y quedé entre los 10 primeros lugares en la competencia nacional de Guatemala.'
        22 = 'Automatización con n8n y Telegram' + [char]11 + 'Creé flujos de trabajo para automatizar procesos e integrar Telegram con distintos servicios.'
        24 = 'Campuslands Guatemala — Formación en Desarrollo de Software | 2025–2026' + [char]11 + 'Programación · Frontend · Backend · Bases de datos · APIs · Testing · Inteligencia artificial · n8n · Git/GitHub · Inglés B1.2'
        25 = 'Universidad Mariano Gálvez de Guatemala — Ingeniería en Sistemas | 2026–actualidad' + [char]11 + 'Bases de datos · Programación · HTML y CSS · Lógica · Matemáticas · Contabilidad · Excel · Presentaciones'
        26 = 'Liceo Comercial Entre Valles — Bachillerato en Ciencias y Letras con Orientación en Computación | 2025'
        28 = 'Trabajo en equipo · Resolución de problemas · Adaptabilidad · Comunicación · Organización · Aprendizaje continuo · Responsabilidad · Perseverancia'
        31 = 'Español: Nativo · Inglés: B1.2'
    }
    foreach ($index in ($paragraphs.Keys | Sort-Object { [int]$_ })) {
        Set-ParaText -Index ([int]$index) -Text $paragraphs[$index] -Alignment 3
    }

    $softHeading = Find-TextRange 'HABILIDADES BLANDAS'
    $languageHeading = Find-TextRange 'IDIOMAS'
    $softStart = $softHeading.Paragraphs.Item(1).Range.Start
    $softEnd = $languageHeading.Paragraphs.Item(1).Range.Start
    $softBlock = $doc.Range($softStart, $softEnd)
    $softBlock.Cut()
    $educationHeading = Find-TextRange 'EDUCACIÓN'
    $educationTarget = $educationHeading.Paragraphs.Item(1).Range
    $educationTarget.Collapse(1)
    $educationTarget.Paste()

    Move-Icon -AltText 'Marcador' -TargetRow 11
    Move-Icon -AltText 'Destinatario' -TargetRow 4
    Move-Icon -AltText 'Correo electrónico' -TargetRow 6
    Move-Icon -AltText 'Mundo' -TargetRow 8
    Move-Icon -AltText 'Marcador' -TargetRow 10
    Set-ContactCell -Row 4 -Text '+502 3180-1268'
    Set-ContactCell -Row 6 -Text 'anntoniocruz7@gmail.com'
    Set-ContactCell -Row 8 -Text ('LinkedIn: linkedin.com/in/eduardo-antonio-cruz-bolaños-834750396' + [char]13 + 'GitHub: github.com/TOTO05-a')
    Set-ContactCell -Row 10 -Text 'San José Pinula, Guatemala'

    $skillsHeading = Find-TextRange 'HABILIDADES TÉCNICAS'
    $skillsStart = $skillsHeading.Paragraphs.Item(1).Range.Start
    $experience = @(
        'EXPERIENCIA PRÁCTICA',
        'Practicante de soporte técnico TI | RV4 DESARROLLOS | 1 sep. – 8 oct. 2025',
        'Edificio ASCEND, 13 calle 05-31, zona 9, Ciudad de Guatemala',
        'Mantenimiento y optimización de laptops y celulares; actualización de equipos e instalación de software desde USB booteable; apoyo con cableado y conexión de impresoras y routers a la red.'
    ) -join [char]13
    $insertRange = $doc.Range($skillsStart, $skillsStart)
    $insertRange.InsertBefore($experience + [char]13)

    Style-Text -Text 'EXPERIENCIA PRÁCTICA' -Size 10.5 -Bold $true -Color 9058846 -Alignment 0
    Style-Text -Text 'Practicante de soporte técnico TI | RV4 DESARROLLOS | 1 sep. – 8 oct. 2025' -Size 9 -Bold $true -Color 3877150 -Alignment 0
    Style-Text -Text 'Edificio ASCEND, 13 calle 05-31, zona 9, Ciudad de Guatemala' -Size 8.5 -Bold $false -Color 3877150 -Alignment 0 -Italic $true
    $tasks = Find-TextRange 'Mantenimiento y optimización de laptops y celulares'
    $tasks.Font.Name = 'Calibri'; $tasks.Font.Size = 8.8; $tasks.Font.Color = 3877150; $tasks.ParagraphFormat.Alignment = 3

    Style-Text -Text 'EDUARDO ANTONIO CRUZ BOLAÑOS' -Size 20 -Bold $true -Color 9058846 -Alignment 1
    Style-Text -Text 'Practicante de soporte técnico TI | RV4 DESARROLLOS' -Size 9 -Bold $true -Color 3877150 -Alignment 0
    $footer = $doc.Sections.Item(1).Footers.Item(1).Range
    $footer.Text = 'EDUARDO ANTONIO CRUZ BOLAÑOS | TÉCNICO EN DESARROLLO DE SOFTWARE'
    $footer.Font.Name = 'Calibri'; $footer.Font.Size = 8; $footer.Font.Color = 9058846
    $footer.ParagraphFormat.Alignment = 1

    $doc.Save()
    [pscustomobject]@{Saved=$true; Pages=$doc.ComputeStatistics(2); Tables=$doc.Tables.Count; InlineImages=$doc.InlineShapes.Count; Font=$doc.Styles.Item(-1).Font.Name; Footer=$footer.Text.Trim()} | Format-List
    $doc.Close(0) | Out-Null
} finally {
    $word.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
}
