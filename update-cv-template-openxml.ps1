$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
$source = 'C:\Users\ac138\OneDrive\Desktop\curriculum-vitae.docx'
$output = 'C:\Users\ac138\Downloads\portfolio (1)\CV-Eduardo-Cruz-Bolanos-Plantilla-Final.docx'
if (!(Test-Path -LiteralPath $source)) { throw 'No se encontró la plantilla original de OneDrive.' }
if (Test-Path -LiteralPath $output) { throw 'El archivo final ya existe; no se sobrescribirá.' }
Copy-Item -LiteralPath $source -Destination $output

$stream = [IO.File]::Open($output, [IO.FileMode]::Open, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
try {
    $archive = [IO.Compression.ZipArchive]::new($stream, [IO.Compression.ZipArchiveMode]::Update, $false)
    try {
        $documentEntry = $archive.GetEntry('word/document.xml')
        $reader = [IO.StreamReader]::new($documentEntry.Open())
        try { [xml]$xml = $reader.ReadToEnd() } finally { $reader.Dispose() }
        $xml.PreserveWhitespace = $true
        $ns = [Xml.XmlNamespaceManager]::new($xml.NameTable)
        $ns.AddNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
        $table = $xml.SelectSingleNode('//w:tbl', $ns)
        $mainCell = $table.SelectSingleNode('./w:tr[1]/w:tc[3]', $ns)
        if ($null -eq $mainCell) { throw 'No se encontró la columna principal de la plantilla.' }

        function Get-ParagraphText($Paragraph) {
            return (@($Paragraph.SelectNodes('.//w:t', $script:ns) | ForEach-Object { $_.InnerText }) -join '')
        }
        function Find-Paragraph([string]$Contains) {
            foreach ($paragraph in $script:mainCell.SelectNodes('./w:p', $script:ns)) {
                if ((Get-ParagraphText $paragraph).Contains($Contains)) { return $paragraph }
            }
            throw "No se encontró el párrafo: $Contains"
        }
        function Set-TextFragments($Paragraph, [string[]]$Values) {
            $nodes = @($Paragraph.SelectNodes('.//w:t', $script:ns))
            if (!$nodes.Count) { throw 'El párrafo no contiene nodos de texto.' }
            if ($Values.Count -gt $nodes.Count) { throw 'La cantidad de fragmentos supera los nodos disponibles.' }
            for ($i = 0; $i -lt $nodes.Count; $i++) {
                $value = if ($i -lt $Values.Count) { $Values[$i] } else { '' }
                $nodes[$i].InnerText = $value
                if ($value -match '^\s|\s$') { $nodes[$i].SetAttribute('space', 'http://www.w3.org/XML/1998/namespace', 'preserve') }
            }
        }
        function Copy-ParagraphWithText($Template, [string]$Text) {
            $paragraph = $Template.CloneNode($true)
            $nodes = @($paragraph.SelectNodes('.//w:t', $script:ns))
            if (!$nodes.Count) { throw 'El párrafo modelo no contiene texto.' }
            $nodes[0].InnerText = $Text
            for ($i = 1; $i -lt $nodes.Count; $i++) { $nodes[$i].InnerText = '' }
            return $paragraph
        }
        function Get-Cell($Row, $Column) { return $script:table.SelectSingleNode("./w:tr[$Row]/w:tc[$Column]", $script:ns) }
        function Get-CellContent($Cell) { return @($Cell.ChildNodes | Where-Object { $_.LocalName -ne 'tcPr' }) }
        function Clear-CellContent($Cell) { foreach ($node in @(Get-CellContent $Cell)) { [void]$Cell.RemoveChild($node) } }
        function Put-CellContent($Cell, $Nodes) { foreach ($node in $Nodes) { [void]$Cell.AppendChild($node) } }

        $profile = Find-Paragraph 'Soy un Desarrollador de Software Junior'
        Set-TextFragments $profile @('Desarrollador de Software Junior con formación en frontend, backend, APIs y bases de datos. He aplicado estos conocimientos en proyectos CLI, automatización con n8n y robótica. Busco integrarme a un equipo donde aportar soluciones y continuar creciendo profesionalmente.')

        $skillsHeading = Find-Paragraph 'HABILIDADES TÉCNICAS'
        $experienceHeading = Find-Paragraph 'EXPERIENCIA TECNOLÓGICA'
        Set-TextFragments $experienceHeading @('EXPERIENCIA PRÁCTICA', '')
        $bodyTemplate = $profile
        $experienceLines = @(
            'Practicante de soporte técnico TI | RV4 DESARROLLOS | 1 sep. – 8 oct. 2025',
            'Edificio ASCEND, 13 calle 05-31, zona 9, Ciudad de Guatemala',
            'Mantenimiento y optimización de laptops y celulares; actualización de equipos e instalación de software mediante USB booteable; apoyo con cableado y conexión de impresoras y routers a la red.'
        )
        $skillsHeading.ParentNode.InsertBefore($experienceHeading, $skillsHeading) | Out-Null
        foreach ($line in $experienceLines) {
            $newParagraph = Copy-ParagraphWithText $bodyTemplate $line
            $skillsHeading.ParentNode.InsertBefore($newParagraph, $skillsHeading) | Out-Null
        }

        $paragraph = Find-Paragraph 'Frontend: HTML'
        Set-TextFragments $paragraph @('Frontend', ': ', 'HTML5 · CSS3 · ', 'JavaScript')
        $paragraph = Find-Paragraph 'Backend: Python'
        Set-TextFragments $paragraph @('Backend', ': ', 'Python · JavaScript · Node.js · Express.js · APIs REST · POO · SOLID · Patrones de diseño', '', '')
        $paragraph = Find-Paragraph 'Bases de datos:'
        Set-TextFragments $paragraph @('Bases de datos', ': ', 'MongoDB · MongoDB Atlas · MongoDB Compass · MySQL · CRUD · Transacciones')
        $paragraph = Find-Paragraph 'Automatización:'
        Set-TextFragments $paragraph @('Automatización e IA', ': ', 'n8n · Telegram · Flujos de trabajo · Automatización de procesos · Fundamentos de inteligencia artificial', '', '')
        $paragraph = Find-Paragraph 'Robótica:'
        Set-TextFragments $paragraph @('Robótica', ': ', 'C++ · Python · Arduino · Sistemas de control · Integración de componentes')
        $paragraph = Find-Paragraph 'Herramientas:'
        $toolValues = @('Herramientas', ': ', 'Git · GitHub · GitHub ', 'Projects', ' · Docker (en aprendizaje) · Postman · ', 'npm', ' · VS Code · ', 'Inquirer', ' · Chalk · ', 'dotenv', ' · ', 'Testing')
        Set-TextFragments $paragraph $toolValues

        $paragraph = Find-Paragraph 'Gestor de Portafolio de Proyectos Freelance'
        Set-TextFragments $paragraph @('Gestor de Proyectos Freelance', 'JavaScript · Node.js · MongoDB · Git · GitHub', 'Aplicación CLI para gestionar clientes, propuestas, proyectos, contratos, entregables y finanzas. Apliqué arquitectura modular, POO, SOLID, validaciones y transacciones en MongoDB.')
        $paragraph = Find-Paragraph 'Inventario Gamer'
        Set-TextFragments $paragraph @('Inventario Gamer', 'Node.js · JavaScript · MongoDB · Mongoose · Docker', '', 'Aplicación CLI para administrar un inventario tecnológico, con integración con MongoDB, estructura modular, variables de entorno y Git.', '')
        $paragraph = Find-Paragraph 'Invernadero Automatizado'
        Set-TextFragments $paragraph @('Invernadero Automatizado | Proyecto académico de bachillerato', 'Robótica · Automatización', 'Participé en una solución automatizada para un invernadero, integrando componentes y aplicando conocimientos de robótica.')
        $paragraph = Find-Paragraph 'Balam 2025'
        Set-TextFragments $paragraph @('RoboFut | Balam 2025 — Armé y programé el robot desde cero y quedé entre los 10 primeros lugares en la competencia nacional de robótica de Guatemala.')
        $paragraph = Find-Paragraph 'Automatización con n8n y'
        Set-TextFragments $paragraph @('Automatización con n8n y ', 'Telegram', ' — Diseñé flujos de trabajo para automatizar procesos e integrar ', 'Telegram', ' con distintos servicios.')

        $paragraph = Find-Paragraph 'Campuslands'
        Set-TextFragments $paragraph @('Campuslands', ' Guatemala — Formación en Desarrollo de Software | 2025–2026', 'Programación · ', 'Frontend', ' · ', 'Backend', ' · Bases de datos · APIs · ', 'Testing', ' · Git/GitHub · n8n · IA · Habilidades blandas · Inglés B1.2')
        $paragraph = Find-Paragraph 'Universidad Mariano Gálvez'
        Set-TextFragments $paragraph @('Universidad Mariano Gálvez de Guatemala — Ingeniería en Sistemas | 2026', '', '–', '', 'actualidad', '', 'Bases de datos · Programación · HTML y CSS · Lógica · Matemáticas · Contabilidad · Excel · Presentaciones')
        $paragraph = Find-Paragraph 'HABILIDADES BLANDAS'
        Set-TextFragments $paragraph @('HABILIDADES BLANDAS', '')
        $paragraph = Find-Paragraph 'Trabajo en equipo'
        Set-TextFragments $paragraph @('Trabajo en equipo · Resolución de problemas · Adaptabilidad · Comunicación · Organización · Aprendizaje continuo', '', '', 'Responsabilidad', ' · Perseverancia')

        $softHeading = Find-Paragraph 'HABILIDADES BLANDAS'
        $softContent = $softHeading.NextSibling
        while ($softContent -and $softContent.LocalName -ne 'p') { $softContent = $softContent.NextSibling }
        $educationHeading = Find-Paragraph 'EDUCACIÓN'
        $parent = $educationHeading.ParentNode
        [void]$parent.RemoveChild($softHeading)
        if ($softContent) { [void]$parent.RemoveChild($softContent) }
        $parent.InsertBefore($softHeading, $educationHeading) | Out-Null
        if ($softContent) { $parent.InsertBefore($softContent, $educationHeading) | Out-Null }

        $projectHeading = Find-Paragraph 'PROYECTOS'
        $project1 = Find-Paragraph 'Gestor de Proyectos Freelance'
        $project2 = Find-Paragraph 'Inventario Gamer'
        $project3 = Find-Paragraph 'Invernadero Automatizado'
        $balam = Find-Paragraph 'RoboFut | Balam 2025'
        $automation = Find-Paragraph 'Automatización con n8n y'
        $languagesHeading = Find-Paragraph 'IDIOMAS'
        $ordered = @($projectHeading, $project1, $project2, $project3, $balam, $automation)
        foreach ($node in $ordered) { [void]$node.ParentNode.RemoveChild($node) }
        foreach ($node in $ordered) { $parent.InsertBefore($node, $languagesHeading) | Out-Null }

        $table = $xml.SelectSingleNode('//w:tbl', $ns)
        $socialCell = Get-Cell 10 2
        $socialParas = @($socialCell.SelectNodes('./w:p', $ns))
        if ($socialParas.Count -ge 2 -and (Get-ParagraphText $socialParas[0]).Contains('GitHub')) {
            [void]$socialCell.RemoveChild($socialParas[1])
            $socialCell.InsertBefore($socialParas[1], $socialParas[0]) | Out-Null
        }
        $socialCell = Get-Cell 10 2
        $socialParas = @($socialCell.SelectNodes('./w:p', $ns))
        $githubParagraph = $null
        $linkedinParagraph = $null
        foreach ($socialParagraph in $socialParas) {
            $socialText = Get-ParagraphText $socialParagraph
            if ($socialText.Contains('GitHub:')) { $githubParagraph = $socialParagraph }
            if ($socialText.Contains('LinkedIn:')) { $linkedinParagraph = $socialParagraph }
        }
        if ($null -eq $githubParagraph -or $null -eq $linkedinParagraph) { throw 'No se encontraron ambos enlaces sociales en la plantilla.' }

        $contactRows = @(4,6,8,10,11)
        $iconTargets = @(4,6,8,11)
        $iconSourceRows = @(6,8,10,4)
        $iconContent = @{}
        $textContent = @{}
        for ($i = 0; $i -lt $iconTargets.Count; $i++) {
            $iconContent[$iconTargets[$i]] = @(Get-CellContent (Get-Cell $iconSourceRows[$i] 1))
        }
        $textContent[4] = @(Get-CellContent (Get-Cell 6 2))
        $textContent[6] = @(Get-CellContent (Get-Cell 8 2))
        $textContent[8] = @($linkedinParagraph)
        $textContent[10] = @($githubParagraph)
        $textContent[11] = @(Get-CellContent (Get-Cell 4 2))
        foreach ($row in $contactRows) { Clear-CellContent (Get-Cell $row 1); Clear-CellContent (Get-Cell $row 2) }
        foreach ($row in $iconTargets) { Put-CellContent (Get-Cell $row 1) $iconContent[$row] }
        foreach ($row in $contactRows) { Put-CellContent (Get-Cell $row 2) $textContent[$row] }

        $documentEntry.Delete()
        $newEntry = $archive.CreateEntry('word/document.xml', [IO.Compression.CompressionLevel]::Optimal)
        $out = $newEntry.Open()
        try {
            $settings = [Xml.XmlWriterSettings]::new()
            $settings.Encoding = [System.Text.UTF8Encoding]::new($false)
            $settings.Indent = $false
            $writer = [Xml.XmlWriter]::Create($out, $settings)
            try { $xml.Save($writer) } finally { $writer.Dispose() }
        } finally { $out.Dispose() }
    } finally { $archive.Dispose() }
} finally { $stream.Dispose() }
Write-Output "Updated OpenXML copy: $output"
