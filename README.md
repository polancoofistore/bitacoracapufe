using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using ClosedXML.Excel;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;

class Program
{
    // ==== CONFIG ====
    static string ConnectionString = @"Server=.\SQLEXPRESS;Database=TuBD;User Id=sa;Password=TuPassword!;TrustServerCertificate=True;";
    static string LogoPath = "./ofistore_logo.png";
    static string ExcelOut = "Preliquidacion_Ofistore.xlsx";
    static string PdfOut   = "Preliquidacion_Ofistore.pdf";

    // Parámetros del SP
    static string FechaInicio = "2025-09-10 00:00:00";
    static string FechaFin    = "2025-09-11 00:00:00";
    static int?   PlazaId     = 137;
    static int?   CarrilId    = null;
    static long?  SesionId    = null;

    static void Main()
    {
        var (resumen, porClase, porPago) = EjecutarSP();

        ExportarExcel(resumen, porClase, porPago);
        ExportarPdf(resumen, porClase, porPago);

        Console.WriteLine("[OK] Archivos generados.");
    }

    static (DataTable, DataTable, DataTable) EjecutarSP()
    {
        using var cn = new SqlConnection(ConnectionString);
        cn.Open();

        // Construye el comando EXEC con parámetros opcionales
        // Nota: también puedes usar SqlCommand con CommandType.StoredProcedure y parámetros.
        var parts = "";
        if (!string.IsNullOrEmpty(FechaInicio) && !string.IsNullOrEmpty(FechaFin))
            parts += $"@FechaInicio='{FechaInicio}', @FechaFin='{FechaFin}'";
        if (PlazaId.HasValue)
            parts += (parts.Length>0?", ":"") + $"@PlazaId={PlazaId.Value}";
        if (CarrilId.HasValue)
            parts += (parts.Length>0?", ":"") + $"@CarrilId={CarrilId.Value}";
        if (SesionId.HasValue)
            parts += (parts.Length>0?", ":"") + $"@SesionId={SesionId.Value}";

        var sql = $"EXEC cobro.sp_GenerarPreliquidacion {parts}";

        using var cmd = new SqlCommand(sql, cn);
        using var reader = cmd.ExecuteReader();

        var dtResumen = new DataTable();
        dtResumen.Load(reader);

        var dtClase = new DataTable();
        if (reader.NextResult())
            dtClase.Load(reader);

        var dtPago = new DataTable();
        if (reader.NextResult())
            dtPago.Load(reader);

        return (dtResumen, dtClase, dtPago);
    }

    static void ExportarExcel(DataTable resumen, DataTable porClase, DataTable porPago)
    {
        using var wb = new XLWorkbook();

        // Hoja Resumen
        var ws1 = wb.Worksheets.Add("Resumen");
        ws1.Cell("A1").Value = "PRELIQUIDACIÓN – Resumen";
        ws1.Cell("A1").Style.Font.Bold = true;
        ws1.Cell("A1").Style.Font.FontSize = 16;
        ws1.Range("A1:H1").Merge();
        ws1.Cell("A2").Value = $"Generado: {DateTime.Now:yyyy-MM-dd HH:mm}";
        ws1.Range("A2:H2").Merge();

        ws1.Cell(4, 1).InsertTable(resumen);
        AutoSize(ws1);

        if (File.Exists(LogoPath))
        {
            var img = ws1.AddPicture(LogoPath).MoveTo(ws1.Cell("A1").Address);
            img.Scale(0.5);
        }

        // Hoja por Clase
        var ws2 = wb.Worksheets.Add("DesglosePorClase");
        ws2.Cell("A1").Value = "PRELIQUIDACIÓN – Desglose por Clase";
        ws2.Cell("A1").Style.Font.Bold = true;
        ws2.Cell("A1").Style.Font.FontSize = 16;
        ws2.Range("A1:H1").Merge();
        ws2.Cell("A2").Value = "Plaza/Carril/Turno según parámetros";
        ws2.Range("A2:H2").Merge();

        ws2.Cell(4, 1).InsertTable(porClase);
        AutoSize(ws2);
        if (File.Exists(LogoPath))
        {
            var img2 = ws2.AddPicture(LogoPath).MoveTo(ws2.Cell("A1").Address);
            img2.Scale(0.5);
        }

        // Hoja por Forma de Pago
        var ws3 = wb.Worksheets.Add("DesglosePorPago");
        ws3.Cell("A1").Value = "PRELIQUIDACIÓN – Desglose por Forma de Pago";
        ws3.Cell("A1").Style.Font.Bold = true;
        ws3.Cell("A1").Style.Font.FontSize = 16;
        ws3.Range("A1:H1").Merge();
        ws3.Cell("A2").Value = "Ofistore – CAPUFE";
        ws3.Range("A2:H2").Merge();

        ws3.Cell(4, 1).InsertTable(porPago);
        AutoSize(ws3);
        if (File.Exists(LogoPath))
        {
            var img3 = ws3.AddPicture(LogoPath).MoveTo(ws3.Cell("A1").Address);
            img3.Scale(0.5);
        }

        wb.SaveAs(ExcelOut);
        Console.WriteLine("[OK] Excel -> " + ExcelOut);
    }

    static void AutoSize(IXLWorksheet ws)
    {
        foreach (var col in ws.ColumnsUsed())
            col.AdjustToContents(5.0, 50.0);
    }

    static void ExportarPdf(DataTable resumen, DataTable porClase, DataTable porPago)
    {
        QuestPDF.Settings.License = LicenseType.Community;

        Document.Create(container =>
        {
            container.Page(page =>
            {
                page.Margin(30);
                page.Size(PageSizes.A4);
                page.Header().Row(row =>
                {
                    if (File.Exists(LogoPath))
                        row.ConstantItem(120).Image(LogoPath);
                    row.RelativeItem().Column(col =>
                    {
                        col.Item().Text("PRELIQUIDACIÓN – Reporte").Bold().FontSize(18);
                        col.Item().Text($"Generado: {DateTime.Now:yyyy-MM-dd HH:mm}");
                    });
                });

                page.Content().Column(col =>
                {
                    col.Item().Text("Resumen").Bold().FontSize(14).PaddingBottom(5);
                    col.Item().Table(table =>
                    {
                        // headers
                        for (int c = 0; c < resumen.Columns.Count; c++)
                            table.Cell().Element(CellHeader).Text(resumen.Columns[c].ColumnName);

                        // rows
                        for (int r = 0; r < resumen.Rows.Count; r++)
                            for (int c = 0; c < resumen.Columns.Count; c++)
                                table.Cell().Element(CellBody).Text(resumen.Rows[r][c]?.ToString() ?? "");
                    });

                    col.Item().PaddingTop(10).Text("Desglose por Clase").Bold().FontSize(14).PaddingBottom(5);
                    col.Item().Table(table =>
                    {
                        for (int c = 0; c < porClase.Columns.Count; c++)
                            table.Cell().Element(CellHeader).Text(porClase.Columns[c].ColumnName);

                        for (int r = 0; r < porClase.Rows.Count; r++)
                            for (int c = 0; c < porClase.Columns.Count; c++)
                                table.Cell().Element(CellBody).Text(porClase.Rows[r][c]?.ToString() ?? "");
                    });

                    col.Item().PaddingTop(10).Text("Desglose por Forma de Pago").Bold().FontSize(14).PaddingBottom(5);
                    col.Item().Table(table =>
                    {
                        for (int c = 0; c < porPago.Columns.Count; c++)
                            table.Cell().Element(CellHeader).Text(porPago.Columns[c].ColumnName);

                        for (int r = 0; r < porPago.Rows.Count; r++)
                            for (int c = 0; c < porPago.Columns.Count; c++)
                                table.Cell().Element(CellBody).Text(porPago.Rows[r][c]?.ToString() ?? "");
                    });
                });

                page.Footer().AlignCenter().Text("Ofistore – Preliquidación").FontSize(10).Light();
            });

            static IContainer CellHeader(IContainer container) =>
                container.Border(0.5f).Background(Colors.Grey.Lighten3).Padding(3).DefaultTextStyle(x => x.SemiBold());

            static IContainer CellBody(IContainer container) =>
                container.Border(0.5f).Padding(3);
        })
        .GeneratePdf(PdfOut);

        Console.WriteLine("[OK] PDF -> " + PdfOut);
    }
}import pyodbc
import pandas as pd
from datetime import datetime
from reportlab.lib import colors
from reportlab.lib.pagesizes import LETTER
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import cm
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import get_column_letter

# ==== CONFIG ====
SERVER = r".\SQLEXPRESS"          # o "192.168.1.10,1433"
DATABASE = "TuBD"
USER = "sa"
PASSWORD = "TuPassword!"
# Usa True si tienes Trusted_Connection:
TRUSTED = False

# Parámetros de ejemplo (puedes cambiar a SesionId si lo prefieres)
PARAMS = {
    "FechaInicio": "2025-09-10 00:00:00",
    "FechaFin"   : "2025-09-11 00:00:00",
    "PlazaId"    : 137,
    "CarrilId"   : None,
    "SesionId"   : None
}

LOGO_PATH = r"./ofistore_logo.png"   # coloca aquí tu logo
SALIDA_XLSX = "Preliquidacion_Ofistore.xlsx"
SALIDA_PDF  = "Preliquidacion_Ofistore.pdf"

def get_connection():
    if TRUSTED:
        conn_str = (
            f"DRIVER={{ODBC Driver 18 for SQL Server}};"
            f"SERVER={SERVER};DATABASE={DATABASE};"
            "Trusted_Connection=Yes;Encrypt=no;"
        )
    else:
        conn_str = (
            f"DRIVER={{ODBC Driver 18 for SQL Server}};"
            f"SERVER={SERVER};DATABASE={DATABASE};"
            f"UID={USER};PWD={PASSWORD};Encrypt=no;"
        )
    return pyodbc.connect(conn_str)

def fetch_resultsets(cursor):
    """
    Lee los 3 resultsets del SP:
      1) Resumen
      2) Por Clase
      3) Por Forma de Pago
    Devuelve DataFrames.
    """
    # Primer resultset
    rows = cursor.fetchall()
    cols = [c[0] for c in cursor.description]
    df_resumen = pd.DataFrame.from_records(rows, columns=cols)

    # Segundo resultset
    cursor.nextset()
    rows = cursor.fetchall()
    cols = [c[0] for c in cursor.description]
    df_por_clase = pd.DataFrame.from_records(rows, columns=cols)

    # Tercer resultset
    cursor.nextset()
    rows = cursor.fetchall()
    cols = [c[0] for c in cursor.description]
    df_por_forma = pd.DataFrame.from_records(rows, columns=cols)

    return df_resumen, df_por_clase, df_por_forma

def autosize_columns(ws):
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            try:
                val = str(cell.value) if cell.value is not None else ""
                max_len = max(max_len, len(val))
            except Exception:
                pass
        ws.column_dimensions[col_letter].width = min(max_len + 2, 50)

def add_header_to_sheet(ws, titulo, subtitulo):
    ws.merge_cells("A1:H1")
    ws["A1"] = titulo
    ws["A1"].font = ws["A1"].font.copy(bold=True, size=16)
    ws.merge_cells("A2:H2")
    ws["A2"] = subtitulo
    ws["A2"].font = ws["A2"].font.copy(italic=True, size=11)

def insert_logo(ws, cell="A1"):
    try:
        img = XLImage(LOGO_PATH)
        img.height = 60
        img.width = 160
        ws.add_image(img, cell)
    except Exception:
        pass  # si no hay logo, continúa

def export_excel(df_resumen, df_por_clase, df_por_forma):
    with pd.ExcelWriter(SALIDA_XLSX, engine="openpyxl") as writer:
        # Hoja Resumen
        df_resumen.to_excel(writer, index=False, sheet_name="Resumen")
        ws = writer.sheets["Resumen"]
        add_header_to_sheet(ws, "PRELIQUIDACIÓN – Resumen", f"Generado: {datetime.now():%Y-%m-%d %H:%M}")
        insert_logo(ws, "A1")
        autosize_columns(ws)

        # Hoja por Clase
        df_por_clase.to_excel(writer, index=False, sheet_name="DesglosePorClase")
        ws2 = writer.sheets["DesglosePorClase"]
        add_header_to_sheet(ws2, "PRELIQUIDACIÓN – Desglose por Clase", f"Plaza/Carril/Turno según parámetros")
        insert_logo(ws2, "A1")
        autosize_columns(ws2)

        # Hoja por Forma de Pago
        df_por_forma.to_excel(writer, index=False, sheet_name="DesglosePorPago")
        ws3 = writer.sheets["DesglosePorPago"]
        add_header_to_sheet(ws3, "PRELIQUIDACIÓN – Desglose por Forma de Pago", f"Ofistore – CAPUFE")
        insert_logo(ws3, "A1")
        autosize_columns(ws3)

    print(f"[OK] Excel -> {SALIDA_XLSX}")

def export_pdf(df_resumen, df_por_clase, df_por_forma):
    doc = SimpleDocTemplate(SALIDA_PDF, pagesize=LETTER, leftMargin=36, rightMargin=36, topMargin=36, bottomMargin=36)
    story = []
    styles = getSampleStyleSheet()

    # Encabezado con logo
    try:
        logo = Image(LOGO_PATH, width=4*cm, height=1.5*cm)
        story.append(logo)
    except Exception:
        pass
    story.append(Paragraph("<b>PRELIQUIDACIÓN – Resumen</b>", styles["Title"]))
    story.append(Paragraph(f"Generado: {datetime.now():%Y-%m-%d %H:%M}", styles["Normal"]))
    story.append(Spacer(1, 10))

    # Tabla Resumen
    if not df_resumen.empty:
        data = [df_resumen.columns.tolist()] + df_resumen.astype(str).values.tolist()
        tbl = Table(data, hAlign="LEFT")
        tbl.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,0), colors.lightgrey),
            ("GRID", (0,0), (-1,-1), 0.5, colors.grey),
            ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
            ("ALIGN", (0,0), (-1,-1), "LEFT"),
        ]))
        story.append(tbl)
    else:
        story.append(Paragraph("Sin datos en el periodo/parámetros.", styles["Italic"]))
    story.append(Spacer(1, 16))

    # Desglose por Clase
    story.append(Paragraph("<b>Desglose por Clase</b>", styles["Heading2"]))
    if not df_por_clase.empty:
        data = [df_por_clase.columns.tolist()] + df_por_clase.astype(str).values.tolist()
        tbl = Table(data, hAlign="LEFT")
        tbl.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,0), colors.lightgrey),
            ("GRID", (0,0), (-1,-1), 0.5, colors.grey),
            ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
            ("ALIGN", (0,0), (-1,-1), "LEFT"),
        ]))
        story.append(tbl)
    else:
        story.append(Paragraph("Sin datos.", styles["Italic"]))
    story.append(Spacer(1, 16))

    # Desglose por Forma de Pago
    story.append(Paragraph("<b>Desglose por Forma de Pago</b>", styles["Heading2"]))
    if not df_por_forma.empty:
        data = [df_por_forma.columns.tolist()] + df_por_forma.astype(str).values.tolist()
        tbl = Table(data, hAlign="LEFT")
        tbl.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,0), colors.lightgrey),
            ("GRID", (0,0), (-1,-1), 0.5, colors.grey),
            ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
            ("ALIGN", (0,0), (-1,-1), "LEFT"),
        ]))
        story.append(tbl)
    else:
        story.append(Paragraph("Sin datos.", styles["Italic"]))

    doc.build(story)
    print(f"[OK] PDF -> {SALIDA_PDF}")

def main():
    # Construcción de llamada al SP
    params_sql = []
    if PARAMS.get("FechaInicio") and PARAMS.get("FechaFin"):
        params_sql.append(f"@FechaInicio='{PARAMS['FechaInicio']}', @FechaFin='{PARAMS['FechaFin']}'")
    if PARAMS.get("PlazaId") is not None:
        params_sql.append(f"@PlazaId={PARAMS['PlazaId']}")
    if PARAMS.get("CarrilId") is not None:
        params_sql.append(f"@CarrilId={PARAMS['CarrilId']}")
    if PARAMS.get("SesionId") is not None:
        params_sql.append(f"@SesionId={PARAMS['SesionId']}")

    call = f"EXEC cobro.sp_GenerarPreliquidacion {', '.join(params_sql)}"
    print("[INFO] Ejecutando:", call)

    with get_connection() as cn:
        cur = cn.cursor()
        cur.execute(call)
        df_resumen, df_por_clase, df_por_forma = fetch_resultsets(cur)

    export_excel(df_resumen, df_por_clase, df_por_forma)
    export_pdf(df_resumen, df_por_clase, df_por_forma)

if __name__ == "__main__":
    main()
-- Por rango de tiempo (día completo)
EXEC cobro.sp_GenerarPreliquidacion
    @FechaInicio = '2025-09-10 00:00:00',
    @FechaFin    = '2025-09-11 00:00:00',
    @PlazaId     = 137,           -- opcional
    @CarrilId    = NULL;          -- opcional

-- Por sesión (ignora fechas si vienen NULL)
EXEC cobro.sp_GenerarPreliquidacion
    @SesionId = 2025091001,
    @PlazaId  = 137,
    @CarrilId = 11;
-- Esquema
IF NOT EXISTS (SELECT 1 FROM sys.schemas WHERE name = 'cobro')
    EXEC('CREATE SCHEMA cobro');

-- Catálogos básicos
IF OBJECT_ID('cobro.CatPlazas') IS NULL
CREATE TABLE cobro.CatPlazas (
    PlazaId        INT           NOT NULL PRIMARY KEY,
    Nombre         NVARCHAR(100) NOT NULL
);

IF OBJECT_ID('cobro.CatCarriles') IS NULL
CREATE TABLE cobro.CatCarriles (
    CarrilId       INT           NOT NULL PRIMARY KEY,
    PlazaId        INT           NOT NULL,
    ClaveCarril    NVARCHAR(20)  NOT NULL,
    CONSTRAINT FK_Carril_Plaza FOREIGN KEY (PlazaId) REFERENCES cobro.CatPlazas(PlazaId)
);

IF OBJECT_ID('cobro.CatClases') IS NULL
CREATE TABLE cobro.CatClases (
    ClaseVehicular TINYINT       NOT NULL PRIMARY KEY,
    Descripcion    NVARCHAR(50)  NOT NULL
);

-- Transacciones base (ajusta columnas a tu realidad CEM.net/CEP-SICE)
IF OBJECT_ID('cobro.Transacciones') IS NULL
CREATE TABLE cobro.Transacciones (
    TransId        BIGINT        IDENTITY(1,1) PRIMARY KEY,
    FechaHora      DATETIME2(0)  NOT NULL,
    PlazaId        INT           NOT NULL,
    CarrilId       INT           NOT NULL,
    SesionId       BIGINT        NULL,              -- sesión/turno
    OperadorId     INT           NULL,
    ClaseVehicular TINYINT       NOT NULL,
    FormaPago      CHAR(1)       NOT NULL,          -- 'E' efectivo, 'T' telepeaje
    Tarifa         DECIMAL(10,2) NOT NULL,          -- tarifa base sin IVA
    Descuento      DECIMAL(10,2) NOT NULL DEFAULT 0,
    IVA            DECIMAL(10,2) NOT NULL DEFAULT 0,
    Importe        DECIMAL(10,2) NOT NULL,          -- total cobrado (Tarifa - Descuento + IVA)
    Estado         VARCHAR(20)   NOT NULL DEFAULT 'VALIDO', -- o 'CANCELADO'
    TagId          NVARCHAR(64)  NULL,              -- si aplica telepeaje
    RefVideo       NVARCHAR(200) NULL               -- opcional para conciliación
);
GO

-- Índices recomendados
CREATE INDEX IX_Transacciones_Rango
ON cobro.Transacciones (PlazaId, CarrilId, FechaHora)
INCLUDE (SesionId, FormaPago, ClaseVehicular, Tarifa, Descuento, IVA, Importe, Estado);

CREATE INDEX IX_Transacciones_Sesion
ON cobro.Transacciones (SesionId, PlazaId, CarrilId)
INCLUDE (FechaHora, FormaPago, ClaseVehicular, Tarifa, Descuento, IVA, Importe, Estado);
GO

