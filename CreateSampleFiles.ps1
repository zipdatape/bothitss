# Script para crear archivos de prueba: BASE HITSS.csv y Excel de Bajas

$BaseFolder   = "C:\Users\DMC\Documents\BotHitss\Base"
$UsuarFolder  = "C:\Users\DMC\Documents\BotHitss\Usuario"
$today        = Get-Date -Format "dd.MM.yy"
$xlsxPath     = Join-Path $UsuarFolder "$today.xlsx"
$csvPath      = Join-Path $BaseFolder "BASE HITSS.csv"

# ── 1. BASE HITSS.csv ─────────────────────────────────────────────────────────
# Sin cabecera | col[0]=Clave | col[1]=DNI | col[2]=Nombre | col[3]=Puesto | col[4]=Area | col[5]=Empresa
# Los DNIs 87654321, 55667788, 19283746 coincidirán con las bajas del Excel
$csvLines = @(
    "EMP001,12345678,GARCIA LOPEZ JUAN CARLOS,ANALISTA TI,OPERACIONES,HITSS",
    "EMP002,87654321,MARTINEZ RUIZ MARIA ELENA,COORDINADORA,RECURSOS HUMANOS,HITSS",
    "EMP003,11223344,RODRIGUEZ PEREZ CARLOS ALBERTO,DESARROLLADOR,SISTEMAS,HITSS",
    "EMP004,44332211,HERNANDEZ GOMEZ ANA PATRICIA,GERENTE,ADMINISTRACION,HITSS",
    "EMP005,55667788,LOPEZ SANCHEZ ROBERTO MIGUEL,SOPORTE TI,HELPDESK,HITSS",
    "EMP006,99887766,TORRES MENDEZ LUCIA FERNANDA,ANALISTA RRHH,RECURSOS HUMANOS,HITSS",
    "EMP007,22334455,MORALES VEGA DIEGO ARMANDO,PROGRAMADOR,DESARROLLO,HITSS",
    "EMP008,66778899,RAMIREZ ORTIZ PATRICIA ELENA,EJECUTIVA,VENTAS,HITSS",
    "EMP009,33221100,FLORES CASTRO JORGE LUIS,CONTADOR,FINANZAS,HITSS",
    "EMP010,77665544,CHAVEZ RUEDA MONICA ISABEL,SUPERVISORA,OPERACIONES,HITSS",
    "EMP011,10293847,REYES SALINAS ERNESTO FABIAN,TECNICO,INFRAESTRUCTURA,HITSS",
    "EMP012,56473829,VARGAS MEDINA CLAUDIA ROSA,ANALISTA,CALIDAD,HITSS",
    "EMP013,19283746,CASTILLO RIOS FERNANDO JOSE,JEFE,PROYECTOS,HITSS",
    "EMP014,64738291,JIMENEZ LUNA ROSA AURORA,ASISTENTE,ADMINISTRACION,HITSS",
    "EMP015,29183746,AGUILAR SOTO MARIO ANTONIO,COORDINADOR,LOGISTICA,HITSS"
)

[System.IO.File]::WriteAllLines($csvPath, $csvLines, [System.Text.Encoding]::GetEncoding("iso-8859-1"))
Write-Host "✓ CSV creado: $csvPath  ($($csvLines.Count) empleados)" -ForegroundColor Green

# ── 2. Excel de Bajas (.xlsx) ─────────────────────────────────────────────────
# Cabecera en fila 1 (se omite al procesar) | col B (índice 1) = DNI
$excel = New-Object -ComObject Excel.Application
$excel.Visible       = $false
$excel.DisplayAlerts = $false

$wb = $excel.Workbooks.Add()
$ws = $wb.Worksheets.Item(1)
$ws.Name = "Hoja1"

# ── Cabecera ──────────────────────────────────────────────────────────────────
$headers = @("CLAVE","DNI","NOMBRE","FECHA_BAJA","MOTIVO","AREA")
for ($i = 0; $i -lt $headers.Count; $i++) {
    $ws.Cells(1, $i + 1).Value2 = $headers[$i]
}
$ws.Range("A1:F1").Font.Bold = $true
$ws.Range("A1:F1").Interior.Color = 0x2D78C8   # Azul Hitss
$ws.Range("A1:F1").Font.Color     = 0xFFFFFF   # Blanco

# ── Filas de baja (DNIs que coinciden con el CSV) ─────────────────────────────
$fechaBaja = Get-Date -Format "dd/MM/yyyy"
$bajas = @(
    @("EMP002", "87654321", "MARTINEZ RUIZ MARIA ELENA",    $fechaBaja, "RENUNCIA VOLUNTARIA", "RECURSOS HUMANOS"),
    @("EMP005", "55667788", "LOPEZ SANCHEZ ROBERTO MIGUEL", $fechaBaja, "TERMINO DE CONTRATO",  "HELPDESK"),
    @("EMP013", "19283746", "CASTILLO RIOS FERNANDO JOSE",  $fechaBaja, "LIQUIDACION",           "PROYECTOS")
)

for ($r = 0; $r -lt $bajas.Count; $r++) {
    $row = $bajas[$r]
    for ($c = 0; $c -lt $row.Count; $c++) {
        $ws.Cells($r + 2, $c + 1).Value2 = $row[$c]
    }
    # Fila alterna: fondo gris muy claro
    if ($r % 2 -eq 1) {
        $ws.Range("A$($r+2):F$($r+2)").Interior.Color = 0xF0F0F0
    }
}

$ws.Columns("A:F").AutoFit() | Out-Null

# 51 = xlOpenXMLWorkbook (.xlsx)
$wb.SaveAs($xlsxPath, 51)
$wb.Close($false)
$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null

Write-Host "✓ Excel creado: $xlsxPath  ($($bajas.Count) bajas)" -ForegroundColor Green
Write-Host ""
Write-Host "Bajas incluidas en el Excel:" -ForegroundColor Cyan
foreach ($b in $bajas) {
    Write-Host "  DNI $($b[1]) - $($b[2]) - $($b[4])" -ForegroundColor Gray
}
Write-Host ""
Write-Host "Empleados que serán eliminados de la base al procesar: EMP002, EMP005, EMP013" -ForegroundColor Yellow
