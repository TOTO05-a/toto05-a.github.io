$ErrorActionPreference = 'Stop'
$path = 'C:\Users\ac138\Downloads\portfolio (1)\CV-Eduardo-Cruz-Bolanos-Plantilla-Actualizada.docx'
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
    $doc = $word.Documents.Open($path, $false, $false)
    $table = $doc.Tables.Item(1)

    function Find-Range {
        param([string]$Text)
        $range = $doc.Content.Duplicate
        $find = $range.Find
        $find.ClearFormatting()
        $find.Text = $Text
        $find.Forward = $true
        $find.Wrap = 0
        if ($find.Execute()) { return $range }
        throw "No se encontró el texto: $Text"
    }

    function Replace-Text {
        param([string]$Old, [string]$New)
        $range = Find-Range $Old
        $range.Text = $New
    }

    function Set-CellText {
        param([int]$Row, [string]$Text)
        $cell = $table.Cell($Row, 2)
        $cell.Range.Text = $Text + [char]13 + [char]7
        $cell.Range.Font.Name = 'Calibri'
        $cell.Range.Font.Size = 8.5
        $cell.Range.Font.Color = 16777215
        $cell.Range.Font.Bold = 0
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

    function Add-Link {
        param([string]$Text, [string]$Address)
        try { $range = Find-Range $Text; $doc.Hyperlinks.Add($range, $Address) | Out-Null } catch { }
    }

    function Style-Text {
        param([string]$Text,[double]$Size,[bool]$Bold,[int]$Color,[int]$Alignment,[bool]$Italic=$false)
        $range = Find-Range $Text
        $range.Font.Name = 'Calibri'
        $range.Font.Size = $Size
        $range.Font.Bold = if ($Bold) { -1 } else { 0 }
        $range.Font.Italic = if ($Italic) { -1 } else { 0 }
        $range.Font.Color = $Color
        $range.ParagraphFormat.Alignment = $Alignment
    }

    Replace-Text 'Soy un Desarrollador de Software Junior, con interés en seguir creciendo en el área de tecnología y convertir lo que aprendo en soluciones que realmente funcionen.' 'Desarrollador de Software Junior con formación en desarrollo web, backend, APIs y bases de datos.'
    Replace-Text 'Me considero una persona responsable, perseverante, adaptable y con disposición para aprender, especialmente cuando me enfrento a algo que todavía no domino.' 'He aplicado estos conocimientos en proyectos CLI, automatización con n8n y robótica.'
    Replace-Text 'Durante mi formación he trabajado en proyectos de desarrollo de software, bases de datos, automatización y robótica, donde he aprendido a resolver problemas, trabajar en equipo y enfrentar retos técnicos.' ''
    Replace-Text 'Busco una oportunidad donde pueda aportar mis conocimientos, asumir nuevos retos y continuar desarrollándome como profesional.' 'Busco integrarme a un equipo donde aportar soluciones y continuar creciendo profesionalmente.'

    Replace-Text 'JavaScrip' 'JavaScript'
    Replace-Text 'Robótica: Robótica · Automatización · Sistemas de control · Integración de componentes' 'Robótica: C++ · Python · Arduino · Sistemas de control · Integración de componentes'
    Replace-Text 'Automatización: n8n · Telegram · Flujos de trabajo · Automatización de procesos' 'Automatización e IA: n8n · Telegram · Flujos de trabajo · Fundamentos de inteligencia artificial'
    Replace-Text 'Herramientas: Git · GitHub · GitHub Projects · Docker · Postman · npm · VS Code · Inquirer · Chalk · dotenv · Testing' 'Herramientas: Git · GitHub · GitHub Projects · Docker (en aprendizaje) · Postman · npm · VS Code · Inquirer · Chalk · dotenv · Testing'

    Replace-Text 'Aplicación CLI para gestionar clientes, propuestas, proyectos, contratos, entregables y finanzas. Implementación de arquitectura modular, POO, SOLID, patrones de diseño, validaciones y transacciones MongoDB.' 'Aplicación CLI para gestionar clientes, propuestas, proyectos, contratos, entregables y finanzas. Implementé arquitectura modular, POO, SOLID, validaciones y transacciones en MongoDB.'
    Replace-Text 'Aplicación CLI para administrar productos de un inventario tecnológico. Integración con MongoDB, estructura modular, variables de entorno y control de versiones con Git.' 'Aplicación CLI para administrar un inventario tecnológico, con MongoDB/Mongoose, estructura modular, variables de entorno y Git.'
    Replace-Text 'Participación en el desarrollo de una solución automatizada para un invernadero, aplicando conocimientos de robótica, integración de componentes y automatización de procesos.' 'Proyecto académico de bachillerato: participé en una solución automatizada para un invernadero, integrando componentes y aplicando robótica.'
    Replace-Text 'EXPERIENCIA TECNOLÓGICA' 'LOGROS DE ROBÓTICA Y AUTOMATIZACIÓN'
    Replace-Text 'Balam 2025 — Participación en actividad/proyecto tecnológico, fortaleciendo experiencia práctica en robótica, resolución de problemas y trabajo colaborativo.' 'RoboFut | Balam 2025 — Armé y programé el robot desde cero; quedé entre los 10 primeros lugares de la competencia nacional de robótica en Guatemala.'
    Replace-Text 'Automatización con n8n y Telegram — Creación de flujos de trabajo para automatizar procesos e integrar Telegram con diferentes servicios.' 'Automatización con n8n y Telegram — Diseñé flujos de trabajo para automatizar procesos e integrar Telegram con distintos servicios.'

    Replace-Text 'Programación · Frontend · Backend · Bases de datos · Testing · Git/GitHub · Automatización' 'Programación · Frontend · Backend · Bases de datos · APIs · Testing · IA · n8n · Git/GitHub · Inglés B1.2'
    Replace-Text 'Feb. 2026' '2026'
    Replace-Text 'Jun. 2026' 'actualidad'
    Replace-Text 'Primer semestre cursado.' 'Primer semestre: bases de datos, programación, HTML y CSS, lógica, matemáticas, contabilidad, Excel y presentaciones.'
    Replace-Text 'Trabajo en equipo · Resolución de problemas · Adaptabilidad · Comunicación · Organización · Aprendizaje continuo · Ganas de progresar · Resolución de problemas' 'Trabajo en equipo · Resolución de problemas · Adaptabilidad · Comunicación · Organización · Aprendizaje continuo · Responsabilidad · Perseverancia'

    $softHeading = Find-Range 'HABILIDADES BLANDAS'
    $languageHeading = Find-Range 'IDIOMAS'
    $softStart = $softHeading.Paragraphs.Item(1).Range.Start
    $softEnd = $languageHeading.Paragraphs.Item(1).Range.Start
    $softBlock = $doc.Range($softStart, $softEnd)
    $softBlock.Cut()
    $educationHeading = Find-Range 'EDUCACIÓN'
    $educationTarget = $educationHeading.Paragraphs.Item(1).Range
    $educationTarget.Collapse(1)
    $educationTarget.Paste()

    Move-Icon -AltText 'Marcador' -TargetRow 11
    Move-Icon -AltText 'Destinatario' -TargetRow 4
    Move-Icon -AltText 'Correo electrónico' -TargetRow 6
    Move-Icon -AltText 'Mundo' -TargetRow 8
    Move-Icon -AltText 'Marcador' -TargetRow 10
    Set-CellText -Row 4 -Text '+502 3180-1268'
    Set-CellText -Row 6 -Text 'anntoniocruz7@gmail.com'
    Set-CellText -Row 8 -Text ('LinkedIn: linkedin.com/in/eduardo-antonio-cruz-bolaños-834750396' + [char]13 + 'GitHub: github.com/TOTO05-a')
    Set-CellText -Row 10 -Text 'San José Pinula, Guatemala'
    Add-Link -Text 'anntoniocruz7@gmail.com' -Address 'mailto:anntoniocruz7@gmail.com'
    Add-Link -Text 'linkedin.com/in/eduardo-antonio-cruz-bolaños-834750396' -Address 'https://www.linkedin.com/in/eduardo-antonio-cruz-bolaños-834750396/'
    Add-Link -Text 'github.com/TOTO05-a' -Address 'https://github.com/TOTO05-a'

    $skillsHeading = Find-Range 'HABILIDADES TÉCNICAS'
    $skillsStart = $skillsHeading.Paragraphs.Item(1).Range.Start
    $experienceText = @(
        'EXPERIENCIA PRÁCTICA',
        'Practicante de soporte técnico TI | RV4 DESARROLLOS | 1 sep. – 8 oct. 2025',
        'Edificio ASCEND, 13 calle 05-31, zona 9, Ciudad de Guatemala',
        'Mantenimiento y optimización de laptops y celulares; actualización de equipos e instalación de software desde USB booteable; apoyo con cableado y conexión de impresoras y routers a la red.'
    ) -join [char]13
    $insertRange = $doc.Range($skillsStart, $skillsStart)
    $insertRange.InsertBefore($experienceText + [char]13)
    Style-Text -Text 'EXPERIENCIA PRÁCTICA' -Size 10.5 -Bold $true -Color 9058846 -Alignment 0
    Style-Text -Text 'Practicante de soporte técnico TI | RV4 DESARROLLOS | 1 sep. – 8 oct. 2025' -Size 9 -Bold $true -Color 3877150 -Alignment 0
    Style-Text -Text 'Edificio ASCEND, 13 calle 05-31, zona 9, Ciudad de Guatemala' -Size 8.5 -Bold $false -Color 3877150 -Alignment 0 -Italic $true
    $tasks = Find-Range 'Mantenimiento y optimización de laptops y celulares'
    $tasks.ParagraphFormat.Alignment = 3
    $footer = $doc.Sections.Item(1).Footers.Item(1).Range
    $footer.Text = 'EDUARDO ANTONIO CRUZ BOLAÑOS | TÉCNICO EN DESARROLLO DE SOFTWARE'
    $footer.Font.Name = 'Calibri'; $footer.Font.Size = 8; $footer.Font.Color = 9058846
    $footer.ParagraphFormat.Alignment = 1
    $doc.Save()
    [pscustomobject]@{Saved=$true;Pages=$doc.ComputeStatistics(2);Tables=$doc.Tables.Count;InlineImages=$doc.InlineShapes.Count;Font=$doc.Styles.Item(-1).Font.Name;Footer=$footer.Text.Trim()}|Format-List
    $doc.Close(0)|Out-Null
} finally {
    $word.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word)|Out-Null
}
