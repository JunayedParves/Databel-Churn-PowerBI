"""Generate the Databel Churn PBIP project (TMDL semantic model + PBIR report).

Re-running overwrites `Databel Churn.SemanticModel/definition` and
`Databel Churn.Report/definition`; object names are derived from stable keys,
so regenerating produces identical files.

    python _build/generate_pbip.py
"""
import hashlib
import json
import os
import shutil
import uuid
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]           # .../Databel Churn
NAME = "Databel Churn"
SM = ROOT / f"{NAME}.SemanticModel"
RP = ROOT / f"{NAME}.Report"
# Power Query needs an absolute path; default to this checkout's data/ folder.
DATA_FOLDER_DEFAULT = str(ROOT / "data")
THEME_SRC = Path(__file__).resolve().parent / "Databel_Executive_Theme.json"
BASE_THEME_SRC = Path(__file__).resolve().parent / "Fluent2-CY26SU08.json"
# Desktop caches themes by name: suffix a content hash so edits always reload.
THEME_FILE = f"Databel_Executive-{hashlib.sha1(THEME_SRC.read_bytes()).hexdigest()[:8]}.json"
BASE_THEME = "Fluent2-CY26SU08"

VC_SCHEMA = ("https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/"
             + os.environ.get("PBIR_VC_VERSION", "2.12.0") + "/schema.json")
PAGE_SCHEMA = "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/page/2.1.0/schema.json"

# Design tokens
NAVY, ORANGE, BLUE, GREY = "#0B1F3A", "#C8471D", "#1F4E79", "#9AA3AF"
PANEL, TEXT2, RULE, WHITE, PEACH, MIN_CELL = "#F5F6F8", "#5B6472", "#D9DDE3", "#FFFFFF", "#F2A07B", "#EEF0F3"
SEGOE = "'''Segoe UI'', wf_segoe-ui_normal, helvetica, arial, sans-serif'"
SEGOE_SB = "'''Segoe UI Semibold'', wf_segoe-ui_semibold, helvetica, arial, sans-serif'"

T, M = "Databel", "_Measures"


def hexid(key, n=20):
    return hashlib.sha1(key.encode()).hexdigest()[:n]


def guid(key):
    return str(uuid.UUID(hashlib.md5(key.encode()).hexdigest()))


# --------------------------------------------------------------------------
# Semantic model (TMDL)
# --------------------------------------------------------------------------
INT_COLS = {"Age", "Account Length (in months)", "Customer Service Calls",
            "Monthly Charge", "Number of Customers in Group"}
SOURCE_COLS = [
    "Customer ID", "Churn Label", "Account Length (in months)", "Local Calls", "Local Mins",
    "Intl Calls", "Intl Mins", "Intl Active", "Intl Plan", "Extra International Charges",
    "Customer Service Calls", "Avg Monthly GB Download", "Unlimited Data Plan",
    "Extra Data Charges", "State", "Phone Number", "Gender", "Age", "Under 30", "Senior",
    "Group", "Number of Customers in Group", "Device Protection & Online Backup",
    "Contract Type", "Payment Method", "Monthly Charge", "Total Charges",
    "Churn Category", "Churn Reason",
]

CALC_COLS = [
    # name, DAX, sortBy, hidden, numeric
    ("Age Group", 'SWITCH(TRUE(), Databel[Age] < 30, "Under 30", Databel[Age] < 40, "30-39", Databel[Age] < 50, "40-49", Databel[Age] < 65, "50-64", "65+")', "Age Sort", False, False),
    ("Age Sort", "SWITCH(TRUE(), Databel[Age] < 30, 1, Databel[Age] < 40, 2, Databel[Age] < 50, 3, Databel[Age] < 65, 4, 5)", None, True, True),
    ("Tenure Group", 'VAR m = Databel[Account Length (in months)] RETURN SWITCH(TRUE(), m <= 6, "0-6", m <= 12, "7-12", m <= 24, "13-24", m <= 48, "25-48", "49+")', "Tenure Sort", False, False),
    ("Tenure Sort", "VAR m = Databel[Account Length (in months)] RETURN SWITCH(TRUE(), m <= 6, 1, m <= 12, 2, m <= 24, 3, m <= 48, 4, 5)", None, True, True),
    ("Service Calls Group", 'IF(Databel[Customer Service Calls] >= 5, "5+", FORMAT(Databel[Customer Service Calls], "0"))', "Service Calls Sort", False, False),
    ("Service Calls Sort", "MIN(Databel[Customer Service Calls], 5)", None, True, True),
    ("Senior Label", 'IF(Databel[Senior] = "Yes", "Senior", "Not senior")', None, False, False),
]

PCT = "0.0%;-0.0%;0.0%"
USD = r"\$#,0;(\$#,0);\$#,0"
MEASURES = [
    ("Total Customers", "COUNTROWS(Databel)", "#,0"),
    ("Churned Customers", 'CALCULATE([Total Customers], Databel[Churn Label] = "Yes")', "#,0"),
    ("Churn Rate", "DIVIDE([Churned Customers], [Total Customers])", PCT),
    ("Total Monthly Revenue", "SUM(Databel[Monthly Charge])", USD),
    ("Monthly Revenue Lost", 'CALCULATE([Total Monthly Revenue], Databel[Churn Label] = "Yes")', USD),
    ("Revenue Lost %", "DIVIDE([Monthly Revenue Lost], [Total Monthly Revenue])", PCT),
    ("Share of Churners", "DIVIDE([Churned Customers], CALCULATE([Churned Customers], ALLSELECTED(Databel[Churn Category])))", PCT),
    ("Group Churn Rate", 'CALCULATE([Churn Rate], Databel[Group] = "Yes")', PCT),
    ("Non-Group Churn Rate", 'CALCULATE([Churn Rate], Databel[Group] = "No")', PCT),
    ("M2M Churn Rate", 'CALCULATE([Churn Rate], Databel[Contract Type] = "Month-to-Month")', PCT),
    ("Two Year Churn Rate", 'CALCULATE([Churn Rate], Databel[Contract Type] = "Two Year")', PCT),
    ("Revenue Lost Subtitle", 'FORMAT([Revenue Lost %], "0%") & " of $" & FORMAT([Total Monthly Revenue]/1000, "0.0") & "K total"', None),
    ("Group Subtitle", 'IF(ISBLANK([Non-Group Churn Rate]), "No non-group customers in selection", "vs " & FORMAT([Non-Group Churn Rate], "0.0%") & " non-group")', None),
    ("Contract Headline", '"Month-to-month customers churn " & FORMAT(DIVIDE([M2M Churn Rate],[Two Year Churn Rate]), "0") & "× more than two-year customers, and " & FORMAT(CALCULATE([Churn Rate], Databel[Tenure Group] = "0-6"), "0%") & " of new customers leave within six months"', None),
    ("Churn Colour", 'SWITCH(TRUE(), [Churn Rate] >= 0.35, "#C8471D", [Churn Rate] < 0.10, "#1F4E79", "#9AA3AF")', None),
    ("Payment Colour", 'SWITCH(TRUE(), [Churn Rate] >= 0.35, "#C8471D", [Churn Rate] < 0.15, "#1F4E79", "#9AA3AF")', None),
]
MEASURE_NAMES = {m[0] for m in MEASURES}
COLUMN_NAMES = set(SOURCE_COLS) | {c[0] for c in CALC_COLS}


def q(name):
    """TMDL object name quoting."""
    return name if name.replace("_", "").isalnum() else "'" + name.replace("'", "''") + "'"


def m_query():
    types = ", ".join(
        f'{{"{c}", {"Int64.Type" if c in INT_COLS else "type text"}}}' for c in SOURCE_COLS)
    lines = [
        "let",
        '    Folder = if Text.EndsWith(DataFolder, "\\") then DataFolder else DataFolder & "\\",',
        f'    Source = Csv.Document(File.Contents(Folder & "Databel - Data.csv"), [Delimiter = ",", Columns = {len(SOURCE_COLS)}, Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),',
        '    #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),',
        f'    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers", {{{types}}}),',
        '    #"Capitalized Each Word" = Table.TransformColumns(#"Changed Type", {{"Intl Plan", Text.Proper, type text}}),',
        '    #"Blank Churn Fields" = Table.ReplaceValue(#"Capitalized Each Word", "", null, Replacer.ReplaceValue, {"Churn Category", "Churn Reason"})',
        "in",
        '    #"Blank Churn Fields"',
    ]
    return lines


def write(path: Path, text: str):
    path.parent.mkdir(parents=True, exist_ok=True)
    # CRLF like Power BI Desktop; newline="" stops Python translating "\n" a second time
    path.write_text(text.replace("\r\n", "\n").replace("\n", "\r\n"), encoding="utf-8", newline="")


def write_json(path: Path, obj):
    write(path, json.dumps(obj, indent=2, ensure_ascii=False) + "\n")


def build_model():
    d = SM / "definition"
    if d.exists():
        shutil.rmtree(d)
    write(d / "database.tmdl", "database\n\tcompatibilityLevel: 1606\n\n")
    write(d / "model.tmdl", "\n".join([
        "model Model",
        "\tculture: en-US",
        "\tdefaultPowerBIDataSourceVersion: powerBI_V3",
        "\tsourceQueryCulture: en-US",
        "\tdataAccessOptions",
        "\t\tlegacyRedirects",
        "\t\treturnErrorValuesAsNull",
        "",
        "annotation __PBI_TimeIntelligenceEnabled = 0",
        "",
        'annotation PBI_QueryOrder = ["DataFolder","Databel"]',
        "",
        'annotation PBI_ProTooling = ["DevMode"]',
        "",
        f"ref table {T}",
        f"ref table {M}",
        "",
    ]))
    write(d / "expressions.tmdl", "\n".join([
        f'expression DataFolder = "{DATA_FOLDER_DEFAULT}" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true]',
        f"\tlineageTag: {guid('expr.DataFolder')}",
        "",
        "\tannotation PBI_ResultType = Text",
        "",
    ]))

    # Databel table
    out = [f"table {T}", f"\tlineageTag: {guid('table.Databel')}", ""]
    for c in SOURCE_COLS:
        is_int = c in INT_COLS
        out += [f"\tcolumn {q(c)}",
                f"\t\tdataType: {'int64' if is_int else 'string'}"]
        if is_int:
            out.append("\t\tformatString: 0")
        out += [f"\t\tlineageTag: {guid('col.' + c)}",
                f"\t\tsummarizeBy: {'sum' if is_int and c not in ('Age',) else 'none'}",
                f"\t\tsourceColumn: {c}",
                "",
                "\t\tannotation SummarizationSetBy = Automatic",
                ""]
    for name, dax, sort_by, hidden, numeric in CALC_COLS:
        out.append(f"\tcolumn {q(name)} = {dax}")
        if numeric:
            out.append("\t\tformatString: 0")
        out.append(f"\t\tlineageTag: {guid('col.' + name)}")
        out.append("\t\tsummarizeBy: none")
        if hidden:
            out.append("\t\tisHidden")
        if sort_by:
            out.append(f"\t\tsortByColumn: {q(sort_by)}")
        out += ["", "\t\tannotation SummarizationSetBy = Automatic", ""]
    out += [f"\tpartition {T} = m", "\t\tmode: import", "\t\tsource ="]
    out += ["\t\t\t\t" + ln for ln in m_query()]
    out += ["", "\tannotation PBI_ResultType = Table", ""]
    write(d / "tables" / f"{T}.tmdl", "\n".join(out))

    # _Measures table
    out = [f"table {M}", f"\tlineageTag: {guid('table._Measures')}", ""]
    for name, dax, fmt in MEASURES:
        out.append(f"\tmeasure {q(name)} = {dax}")
        if fmt:
            out.append(f"\t\tformatString: {fmt}")
        out += [f"\t\tlineageTag: {guid('measure.' + name)}", ""]
    out += ["\tcolumn Column", "\t\tisHidden", "\t\tformatString: 0",
            f"\t\tlineageTag: {guid('col._Measures.Column')}",
            "\t\tsummarizeBy: sum", "\t\tisNameInferred", "\t\tsourceColumn: [Column]", "",
            "\t\tannotation SummarizationSetBy = Automatic", "",
            f"\tpartition {M} = calculated", "\t\tmode: import",
            '\t\tsource = Row("Column", BLANK())', "",
            "\tannotation PBI_Id = " + hexid("measures-table", 32), ""]
    write(d / "tables" / f"{M}.tmdl", "\n".join(out))

    write_json(SM / "definition.pbism", {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/semanticModel/definitionProperties/1.0.0/schema.json",
        "version": "4.2",
        "settings": {},
    })
    write_json(SM / ".platform", {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/gitIntegration/platformProperties/2.0.0/schema.json",
        "metadata": {"type": "SemanticModel", "displayName": NAME},
        "config": {"version": "2.0", "logicalId": guid("logical.sm")},
    })


# --------------------------------------------------------------------------
# PBIR helpers
# --------------------------------------------------------------------------
def lit(v):
    return {"expr": {"Literal": {"Value": v}}}


def s(v):
    return lit("'" + v.replace("'", "''") + "'")


def b(v):
    return lit("true" if v else "false")


def dnum(v):
    return lit(f"{v}D")


def lnum(v):
    return lit(f"{v}L")


def color(h):
    return {"solid": {"color": lit(f"'{h}'")}}


def col(p, ent=T):
    assert p in COLUMN_NAMES, p
    return {"Column": {"Expression": {"SourceRef": {"Entity": ent}}, "Property": p}}


def meas(p):
    assert p in MEASURE_NAMES, p
    return {"Measure": {"Expression": {"SourceRef": {"Entity": M}}, "Property": p}}


def qref(field):
    k = "Column" if "Column" in field else "Measure"
    return f"{field[k]['Expression']['SourceRef']['Entity']}.{field[k]['Property']}"


def proj(field, active=None):
    k = "Column" if "Column" in field else "Measure"
    p = {"field": field, "queryRef": qref(field), "nativeQueryRef": field[k]["Property"]}
    if active is not None:
        p["active"] = active
    return p


def props(**kw):
    return {"properties": kw}


def sel(entry, selector):
    entry = dict(entry)
    entry["selector"] = selector
    return entry


def filter_name(key):
    return "Filter" + hexid("filter." + key, 24)


class Page:
    def __init__(self, key, display, number):
        self.key, self.display, self.number = key, display, number
        self.name = hexid("page." + key)
        self.visuals = []
        self.z = 0

    def add(self, key, x, y, w, h, visual, filters=None, extra=None):
        self.z += 1000
        v = {
            "$schema": VC_SCHEMA,
            "name": hexid(f"{self.key}.{key}"),
            "position": {"x": x, "y": y, "z": self.z, "height": h, "width": w, "tabOrder": self.z},
            "visual": visual,
        }
        if filters:
            v["filterConfig"] = {"filters": filters}
        if extra:
            v.update(extra)
        self.visuals.append(v)
        return v


def chrome(title=None, bg=None, pad=None, alt=None):
    """visualContainerObjects shared by every visual."""
    o = {
        "title": [props(show=b(False))],
        "subTitle": [props(show=b(False))],
        "divider": [props(show=b(False))],
        "background": [props(show=b(bg is not None), **({"color": color(bg), "transparency": dnum(0)} if bg else {}))],
        "border": [props(show=b(False))],
        "dropShadow": [props(show=b(False))],
        "visualHeader": [props(show=b(False))],
    }
    if title:
        o["title"] = [props(show=b(True), text=s(title), fontSize=dnum(11), bold=b(True),
                            fontColor=color(NAVY), alignment=s("left"), titleWrap=b(False),
                            fontFamily=lit(SEGOE_SB))]
    p = pad if pad is not None else (0, 0, 0, 0)
    o["padding"] = [props(top=dnum(p[0]), right=dnum(p[1]), bottom=dnum(p[2]), left=dnum(p[3]))]
    if alt:
        o["general"] = [props(altText=s(alt))]
    return o


def run(text, size, clr, weight=None, font="Segoe UI"):
    st = {"fontFamily": font, "fontSize": f"{size}pt", "color": clr}
    if weight:
        st["fontWeight"] = weight
    return {"value": text, "textStyle": st}


def para(runs, align="left"):
    return {"textRuns": runs, "horizontalTextAlignment": align}


def textbox(paragraphs, bg=None, pad=None):
    return {"visualType": "textbox",
            "objects": {"general": [props(paragraphs=paragraphs)]},
            "visualContainerObjects": chrome(bg=bg, pad=pad)}


def shape_rect(fill):
    return {"visualType": "shape",
            "objects": {
                "shape": [sel(props(tileShape=s("rectangle")), {"id": "default"})],
                "fill": [sel(props(show=b(True), fillColor=color(fill), transparency=dnum(0)), {"id": "default"})],
                # outline would otherwise draw in theme dataColors[0] (orange)
                "outline": [props(show=b(False)),
                            sel(props(show=b(False), lineColor=color(fill), weight=dnum(0)), {"id": "default"})],
            },
            "visualContainerObjects": chrome()}


def sort_def(field, direction):
    return {"sort": [{"field": field, "direction": direction}], "isDefaultSort": False}


def chart(vtype, category, measure, title, alt, colour=None, tooltips=None, sort=None,
          static_colours=None, label_size=10):
    qs = {"Category": {"projections": [proj(col(category), active=True)]},
          "Y": {"projections": [proj(meas(measure))]}}
    if tooltips:
        qs["Tooltips"] = {"projections": [proj(meas(t)) for t in tooltips]}
    query = {"queryState": qs}
    if sort:
        field = meas(measure) if sort[0] == "measure" else col(category)
        query["sortDefinition"] = sort_def(field, sort[1])
    if colour:  # field-value conditional colour from a text measure
        data_point = [sel(props(fill={"solid": {"color": {"expr": meas(colour)}}}),
                          {"data": [{"dataViewWildcard": {"matchingOption": 1}}]})]
    else:
        data_point = [props(defaultColor=color(GREY))]
        for value, hexc in (static_colours or {}).items():
            data_point.append(sel(props(fill=color(hexc)), {"data": [{"scopeId": {"Comparison": {
                "ComparisonKind": 0, "Left": col(category),
                "Right": {"Literal": {"Value": "'" + value.replace("'", "''") + "'"}}}}}]}))
    return {
        "visualType": vtype,
        "query": query,
        "objects": {
            "dataPoint": data_point,
            "categoryAxis": [props(show=b(True), fontSize=dnum(9), labelColor=color(TEXT2),
                                   fontFamily=lit(SEGOE), showAxisTitle=b(False), gridlineShow=b(False))],
            "valueAxis": [props(show=b(False), showAxisTitle=b(False), gridlineShow=b(False))],
            "labels": [props(show=b(True), fontSize=dnum(label_size), bold=b(True), color=color(NAVY),
                             fontFamily=lit(SEGOE_SB), labelPosition=s("OutsideEnd"))],
            "legend": [props(show=b(False))],
        },
        "visualContainerObjects": chrome(title=title, alt=alt),
        "drillFilterOtherVisuals": True,
    }


def kpi_card(measure, caption, value_colour, subtitle=None, subtitle_measure=None, size=28, alt=None):
    ref = qref(meas(measure))
    md = {"metadata": ref}
    label = [sel(props(show=b(True), position=s("belowValue"), fontSize=dnum(9), fontColor=color(TEXT2),
                       fontFamily=lit(SEGOE), bold=b(False), horizontalAlignment=s("left")), md)]
    if subtitle_measure:
        label.insert(0, sel(props(text={"expr": meas(subtitle_measure)}),
                            {"data": [{"dataViewWildcard": {"matchingOption": 0}}], "metadata": ref}))
    else:
        label[0]["properties"]["text"] = s(subtitle)
    vco = chrome(bg=PANEL, pad=(7, 12, 0, 12), alt=alt or f"{caption}: {measure}")
    vco["title"] = [props(show=b(True), text=s(caption), fontSize=dnum(9), bold=b(False),
                          fontColor=color(TEXT2), alignment=s("left"), titleWrap=b(False),
                          fontFamily=lit(SEGOE_SB), background=color(PANEL))]
    vco["spacing"] = [props(customizeSpacing=b(True), spaceBelowTitleArea=dnum(0), spaceAbovePlotArea=dnum(0)),
                      sel(props(verticalSpacing=dnum(0)), {"id": "default"})]
    return {
        "visualType": "cardVisual",
        "query": {"queryState": {"Data": {"projections": [proj(meas(measure))]}}},
        "objects": {
            "value": [sel(props(fontSize=dnum(size), bold=b(True), fontColor=color(value_colour),
                                fontFamily=lit(SEGOE_SB), horizontalAlignment=s("left"),
                                labelDisplayUnits=lit("1D")), md)],
            "label": label,
            "layout": [props(style=s("Cards"), paddingUniform=lnum(0), topOuterMargin=lnum(0),
                             bottomOuterMargin=lnum(0), leftOuterMargin=lnum(0), rightOuterMargin=lnum(0))],
            "fillCustom": [sel(props(show=b(True), fillColor=color(PANEL)), md)],
            "outline": [sel(props(show=b(False)), md)],
            "divider": [sel(props(show=b(False)), md)],
            "accentBar": [sel(props(show=b(False)), md)],
        },
        "visualContainerObjects": vco,
        "drillFilterOtherVisuals": True,
    }


def slicer(field, header):
    return {
        "visualType": "slicer",
        "query": {"queryState": {"Values": {"projections": [proj(col(field), active=True)]}}},
        "objects": {
            "data": [props(mode=s("Dropdown"))],
            "header": [props(show=b(True), text=s(header), textSize=dnum(9), fontColor=color(TEXT2),
                             fontFamily=lit(SEGOE))],
            "items": [props(textSize=dnum(10), fontColor=color(NAVY), fontFamily=lit(SEGOE))],
            "selection": [props(selectAllCheckboxEnabled=b(True), singleSelect=b(False))],
        },
        "visualContainerObjects": chrome(bg=WHITE, alt=f"{header} filter"),
        "syncGroup": {"groupName": "sync_" + field.replace(" ", "_"), "fieldChanges": True, "filterChanges": True},
        "drillFilterOtherVisuals": True,
    }


def matrix(rows, columns, title, alt, value_size=18, row_padding=24):
    ref = qref(meas("Churn Rate"))
    wild = {"data": [{"dataViewWildcard": {"matchingOption": 1}}], "metadata": ref}
    gradient = {"solid": {"color": {"expr": {"FillRule": {
        "Input": {"SelectRef": {"ExpressionName": ref}},
        "FillRule": {"linearGradient2": {
            "min": {"color": {"Literal": {"Value": f"'{MIN_CELL}'"}}},
            "max": {"color": {"Literal": {"Value": f"'{ORANGE}'"}}},
            "nullColoringStrategy": {"strategy": {"Literal": {"Value": "'noColor'"}}}}}}}}}}
    font_rule = {"solid": {"color": {"expr": {"Conditional": {
        "Cases": [{"Condition": {"Comparison": {"ComparisonKind": 2, "Left": meas("Churn Rate"),
                                                 "Right": {"Literal": {"Value": "0.35D"}}}},
                   "Value": {"Literal": {"Value": f"'{WHITE}'"}}}],
        "DefaultValue": {"Literal": {"Value": f"'{NAVY}'"}}}}}}}
    vco = chrome(title=title, alt=alt)
    vco["stylePreset"] = [props(name=s("None"))]
    return {
        "visualType": "pivotTable",
        "query": {"queryState": {
            "Rows": {"projections": [proj(col(rows), active=True)]},
            "Columns": {"projections": [proj(col(columns), active=True)]},
            "Values": {"projections": [proj(meas("Churn Rate"))]},
        }},
        "objects": {
            "values": [props(fontSize=dnum(value_size), bold=b(True), fontFamily=lit(SEGOE_SB)),
                       sel(props(backColor=gradient, fontColor=font_rule), wild)],
            "columnHeaders": [props(fontSize=dnum(9), fontColor=color(TEXT2), fontFamily=lit(SEGOE),
                                    alignment=s("Center"), columnAdjustment=s("growToFit"),
                                    autoSizeColumnWidth=b(True), backColor=color(WHITE))],
            "rowHeaders": [props(fontSize=dnum(9), fontColor=color(TEXT2), fontFamily=lit(SEGOE),
                                 backColor=color(WHITE))],
            "grid": [props(gridVertical=b(True), gridVerticalColor=color(WHITE), gridVerticalWeight=dnum(4),
                           gridHorizontal=b(True), gridHorizontalColor=color(WHITE), gridHorizontalWeight=dnum(4),
                           rowPadding=dnum(row_padding), outlineColor=color(WHITE))],
            "subTotals": [props(rowSubtotals=b(False), columnSubtotals=b(False))],
        },
        "visualContainerObjects": vco,
        "drillFilterOtherVisuals": True,
    }


def not_blank_filter(key, column):
    return {"name": filter_name(key), "field": col(column), "type": "Categorical",
            "filter": {"Version": 2, "From": [{"Name": "d", "Entity": T, "Type": 0}],
                       "Where": [{"Condition": {"Not": {"Expression": {"In": {
                           "Expressions": [{"Column": {"Expression": {"SourceRef": {"Source": "d"}}, "Property": column}}],
                           "Values": [[{"Literal": {"Value": "null"}}]]}}}}}]},
            "howCreated": "User"}


def topn_filter(key, column, order_expr, n=5):
    """TopN visual filter. order_expr gets aliases: 'd' = Databel, 'm' = _Measures."""
    return {"name": filter_name(key), "field": col(column), "type": "TopN",
            "filter": {"Version": 2, "From": [
                {"Name": "subquery", "Expression": {"Subquery": {"Query": {
                    "Version": 2,
                    "From": [{"Name": "d", "Entity": T, "Type": 0}, {"Name": "m", "Entity": M, "Type": 0}],
                    "Select": [{"Column": {"Expression": {"SourceRef": {"Source": "d"}}, "Property": column}, "Name": "field"}],
                    "OrderBy": [{"Direction": 2, "Expression": order_expr}],
                    "Top": n}}}, "Type": 2},
                {"Name": "d", "Entity": T, "Type": 0}],
                "Where": [{"Condition": {"In": {
                    "Expressions": [{"Column": {"Expression": {"SourceRef": {"Source": "d"}}, "Property": column}}],
                    "Table": {"SourceRef": {"Source": "subquery"}}}}}]},
            "howCreated": "User"}


def so_what(page, key, x, y, w, h, text, title="So what"):
    page.add(key + "_bar", x, y, w, 3, shape_rect(ORANGE))
    page.add(key, x, y + 3, w, h - 3, textbox(
        [para([run(title, 11, NAVY, "bold", "Segoe UI Semibold")]),
         para([run(text, 10, NAVY)])], bg=PANEL, pad=(10, 14, 10, 14)))


# --------------------------------------------------------------------------
# Report pages
# --------------------------------------------------------------------------
PAGES = [
    Page("summary", "Executive Summary", 1),
    Page("contract", "Contract & Tenure", 2),
    Page("service", "Service & Plan Fit", 3),
    Page("who", "Who Churns", 4),
]


def page_chrome(p, headline=None, subtitle=None, headline_measure=None, footer=""):
    p.add("brand", 40, 12, 600, 32, textbox(
        [para([run("DATABEL · CUSTOMER CHURN REVIEW", 9, NAVY, "bold", "Segoe UI Semibold")])],
        pad=(6, 0, 0, 0)))
    nav = {
        "visualType": "pageNavigator",
        "objects": {
            "fill": [props(show=b(False)),
                     sel(props(show=b(False)), {"id": "default"}),
                     sel(props(show=b(False)), {"id": "selected"}),
                     sel(props(show=b(False)), {"id": "hover"})],
            "text": [props(show=b(True)),
                     sel(props(fontSize=dnum(9), fontColor=color(TEXT2), bold=b(False), fontFamily=lit(SEGOE),
                               horizontalAlignment=s("center"), leftMargin=lnum(0), rightMargin=lnum(0)),
                         {"id": "default"}),
                     sel(props(fontSize=dnum(9), fontColor=color(NAVY), bold=b(True), fontFamily=lit(SEGOE_SB)),
                         {"id": "selected"}),
                     sel(props(fontColor=color(NAVY)), {"id": "hover"})],
            "outline": [props(show=b(False)),
                        sel(props(show=b(False)), {"id": "default"}),
                        sel(props(show=b(False)), {"id": "selected"})],
            "accentBar": [sel(props(show=b(False)), {"id": "default"}),
                          sel(props(show=b(False)), {"id": "selected"})],
        },
        "visualContainerObjects": chrome(alt="Page navigation"),
    }
    # Wider than the brief's 440px so "Contract & Tenure" / "Service & Plan Fit" aren't truncated.
    p.add("nav", 680, 12, 560, 32, nav, extra={"howCreated": "InsertVisualButton"})
    p.add("rule_top", 40, 48, 1200, 2, shape_rect(NAVY))
    if headline_measure:
        # Dynamic textbox: the text run points at a `values` entry bound to the measure,
        # so the sentence wraps like static text and updates with slicers.
        head_run = {"value": {"propertyIdentifier": {"objectName": "values", "propertyName": "expr"},
                              "selector": {"id": "headline"}},
                    "textStyle": {"fontFamily": "Segoe UI Semibold", "fontSize": "20pt", "color": NAVY}}
        vis = textbox([para([head_run]), para([run(subtitle, 9, TEXT2)])], pad=(0, 0, 0, 0))
        vis["objects"]["values"] = [sel(props(expr={"expr": meas(headline_measure)}), {"id": "headline"})]
        p.add("headline", 40, 56, 1200, 96, vis)
    else:
        p.add("headline", 40, 56, 1200, 96, textbox(
            [para([run(headline, 20, NAVY, None, "Segoe UI Semibold")]),
             para([run(subtitle, 9, TEXT2)])], pad=(0, 0, 0, 0)))
    p.add("rule_foot", 40, 678, 1200, 1, shape_rect(RULE))
    p.add("footer", 40, 682, 1000, 26, textbox([para([run(footer, 8, TEXT2)])]))
    p.add("page_no", 1140, 682, 100, 26, textbox([para([run(str(p.number), 8, TEXT2)], "right")]))


def slicer_row(p, fields):
    labels = {"State": "State", "Age Group": "Age group", "Gender": "Gender", "Contract Type": "Contract type"}
    for i, f in enumerate(fields):
        p.add("slicer_" + f, 40 + i * 202, 154, 190, 52, slicer(f, labels[f]))


def build_pages():
    summary, contract, service, who = PAGES

    # ---- Page 1: Executive Summary -------------------------------------
    page_chrome(summary,
                headline="Databel lost 1 in 4 customers and a third of monthly revenue, driven mostly by competitor offers and poor service",
                subtitle="All customers · churn = customers with Churn Label = Yes",
                footer="Source: Databel customer dataset, n = 6,687 customers. Revenue at risk = sum of Monthly Charge for churned customers.")
    kpis = [
        ("CUSTOMERS", "Total Customers", NAVY, NAVY, "Total base", None),
        ("CHURNED", "Churned Customers", NAVY, ORANGE, "Customers lost", None),
        ("CHURN RATE", "Churn Rate", ORANGE, ORANGE, "Of all customers", None),
        ("MONTHLY REVENUE LOST", "Monthly Revenue Lost", ORANGE, ORANGE, None, "Revenue Lost Subtitle"),
        ("GROUP-PLAN CHURN", "Group Churn Rate", BLUE, BLUE, None, "Group Subtitle"),
    ]
    for i, (cap, m, fg, accent, sub, sub_m) in enumerate(kpis):
        x = 40 + i * 242
        summary.add(f"kpi{i}", x, 160, 230, 90, kpi_card(m, cap, fg, sub, sub_m))
        summary.add(f"kpi{i}_accent", x, 160, 230, 3, shape_rect(accent))

    summary.add("cat", 40, 270, 381, 400, chart(
        "clusteredBarChart", "Churn Category", "Share of Churners",
        "Churn category, % of churned customers",
        "Bar chart of churn category as a share of churned customers; Competitor is the largest",
        sort=("measure", "Descending"), static_colours={"Competitor": ORANGE}, label_size=11),
        filters=[not_blank_filter("summary.cat.blank", "Churn Category")])
    reason_count = {"Aggregation": {"Expression": {"Column": {"Expression": {"SourceRef": {"Source": "d"}},
                                                              "Property": "Churn Reason"}}, "Function": 5}}
    summary.add("reasons", 449, 270, 381, 400, chart(
        "clusteredBarChart", "Churn Reason", "Churned Customers",
        "Top 5 churn reasons, # customers",
        "Bar chart of the five most common churn reasons by number of churned customers",
        sort=("measure", "Descending"),
        static_colours={"Competitor made better offer": ORANGE, "Competitor had better devices": ORANGE},
        label_size=11),
        filters=[topn_filter("summary.reasons.top5", "Churn Reason", reason_count)])

    def rec(n, text):
        return para([run(f"{n}   ", 11, PEACH, "bold", "Segoe UI Semibold"), run(text, 11, WHITE)])
    summary.add("recommend", 858, 270, 381, 400, textbox([
        para([run("What we recommend", 11, WHITE, "bold", "Segoe UI Semibold")]),
        para([run(" ", 6, WHITE)]),
        rec(1, "Move month-to-month customers onto 1–2 year contracts; they churn at 46% vs 3% on two-year plans."),
        para([run(" ", 6, WHITE)]),
        rec(2, "Escalate any customer on their 3rd support call; 88–100% of them leave."),
        para([run(" ", 6, WHITE)]),
        rec(3, "Fix plan mismatch: sell the international plan to active callers and push group plans, which churn at 7%."),
    ], bg=NAVY, pad=(18, 20, 18, 20)))

    # ---- Page 2: Contract & Tenure -------------------------------------
    page_chrome(contract, subtitle="Churn rate, % of customers in each segment", headline_measure="Contract Headline",
                footer="Source: Databel customer dataset, n = 6,687. Tenure = Account Length (in months).")
    slicer_row(contract, ["State", "Age Group", "Gender"])
    contract.add("by_contract", 40, 210, 381, 440, chart(
        "columnChart", "Contract Type", "Churn Rate", "By contract type",
        "Column chart of churn rate by contract type", colour="Churn Colour",
        tooltips=["Total Customers"], sort=("category", "Ascending"), label_size=12))
    contract.add("by_tenure", 449, 210, 381, 440, chart(
        "columnChart", "Tenure Group", "Churn Rate", "By tenure, months with Databel",
        "Column chart of churn rate by tenure group", colour="Churn Colour",
        tooltips=["Total Customers"], sort=("category", "Ascending"), label_size=12))
    contract.add("by_payment", 858, 210, 381, 230, chart(
        "clusteredBarChart", "Payment Method", "Churn Rate", "By payment method",
        "Bar chart of churn rate by payment method", colour="Payment Colour",
        tooltips=["Total Customers"], sort=("measure", "Descending"), label_size=11))
    so_what(contract, "so_what", 858, 455, 381, 195,
            "Offer a discount or device upgrade for switching to a 12-month contract, and target the first six "
            "months with onboarding check-ins. Credit-card payers churn at less than half the rate of direct-debit "
            "and paper-check payers.")

    # ---- Page 3: Service & Plan Fit ------------------------------------
    page_chrome(service,
                headline="A third support call is the tipping point, and customers on the wrong plan leave far more often",
                subtitle="Churn rate, % of customers in each segment",
                footer="Source: Databel customer dataset, n = 6,687. 5+ calls grouped.")
    slicer_row(service, ["Contract Type", "State", "Age Group"])
    service.add("by_calls", 40, 210, 450, 410, chart(
        "columnChart", "Service Calls Group", "Churn Rate", "By number of customer service calls",
        "Column chart of churn rate by number of customer service calls", colour="Churn Colour",
        tooltips=["Total Customers"], sort=("category", "Ascending"), label_size=12))
    service.add("calls_note", 40, 624, 450, 26, textbox([para([run("3+ calls = 47% of all churners", 9, TEXT2)])]))
    service.add("intl", 530, 210, 380, 380, matrix(
        "Intl Plan", "Intl Active", "International plan fit",
        "Matrix of churn rate by international plan (rows) and international calling activity (columns)",
        value_size=20, row_padding=40))
    service.add("intl_note", 530, 594, 380, 56, textbox([
        para([run("Rows: has international plan · Columns: makes international calls.", 9, TEXT2)]),
        para([run("Paying for an unused plan (Yes / No) is the worst fit.", 9, TEXT2)])]))
    service.add("plans_title", 950, 210, 290, 26, textbox(
        [para([run("Plans and add-ons", 11, NAVY, "bold", "Segoe UI Semibold")])]))
    service.add("unlimited", 950, 236, 290, 125, chart(
        "clusteredBarChart", "Unlimited Data Plan", "Churn Rate", "Unlimited data plan",
        "Bar chart of churn rate by unlimited data plan", colour="Churn Colour",
        sort=("measure", "Descending"), label_size=10))
    service.add("protection", 950, 365, 290, 125, chart(
        "clusteredBarChart", "Device Protection & Online Backup", "Churn Rate", "Device protection & backup",
        "Bar chart of churn rate by device protection and online backup", colour="Churn Colour",
        sort=("measure", "Descending"), label_size=10))
    for v in service.visuals[-2:]:
        v["visual"]["visualContainerObjects"]["title"][0]["properties"]["fontSize"] = dnum(9)
        v["visual"]["visualContainerObjects"]["title"][0]["properties"]["bold"] = b(False)
        v["visual"]["visualContainerObjects"]["title"][0]["properties"]["fontColor"] = color(TEXT2)
    so_what(service, "so_what", 950, 500, 290, 150,
            "Route any customer with 2+ calls to a retention team, and review plan fit every quarter.")

    # ---- Page 4: Who Churns --------------------------------------------
    page_chrome(who,
                headline="Seniors on month-to-month contracts are the highest-risk group, while group plans almost eliminate churn",
                subtitle="Churn rate, % of customers in each segment",
                footer="Source: Databel customer dataset, n = 6,687. Senior = age 65+. CA has only 68 customers.")
    slicer_row(who, ["Contract Type", "State"])
    who.add("by_age", 40, 210, 381, 230, chart(
        "columnChart", "Age Group", "Churn Rate", "By age group",
        "Column chart of churn rate by age group", colour="Churn Colour",
        tooltips=["Total Customers"], sort=("category", "Ascending"), label_size=11))
    who.add("senior_contract", 40, 450, 381, 200, matrix(
        "Contract Type", "Senior Label", "Senior × contract",
        "Matrix of churn rate by contract type (rows) and senior status (columns)",
        value_size=12, row_padding=8))

    who.add("group_title", 449, 210, 381, 26, textbox(
        [para([run("Group vs individual plans", 11, NAVY, "bold", "Segoe UI Semibold")])]))
    for i, (m, cap, clr) in enumerate([("Group Churn Rate", "GROUP PLAN", BLUE),
                                        ("Non-Group Churn Rate", "INDIVIDUAL", ORANGE)]):
        x = 449 + i * 195
        who.add(f"grp{i}", x, 240, 186, 100, kpi_card(m, cap, clr, "Churn rate", size=28))
        who.add(f"grp{i}_accent", x, 240, 186, 3, shape_rect(clr))
    who.add("by_gender", 449, 365, 381, 285, chart(
        "clusteredBarChart", "Gender", "Churn Rate", "By gender",
        "Bar chart of churn rate by gender; differences are small", sort=("measure", "Descending"),
        label_size=11))

    # No map: Shape Map is blocked in current Desktop builds (FilledMapVisualNotEnabled),
    # so the state view is a Top-10 table with in-cell data bars.

    rate_ref = qref(meas("Churn Rate"))
    data_bars = sel(props(dataBars={
        "positiveColor": color(GREY), "negativeColor": color(GREY), "axisColor": color(RULE),
        "reverseDirection": b(False), "hideText": b(False)}), {"metadata": rate_ref})
    table = {
        "visualType": "tableEx",
        "query": {"queryState": {"Values": {"projections": [
            proj(col("State")), {**proj(meas("Total Customers")), "displayName": "Customers"},
            {**proj(meas("Churn Rate")), "displayName": "Churn"}]}},
            "sortDefinition": sort_def(meas("Churn Rate"), "Descending")},
        "objects": {
            "values": [props(fontSize=dnum(10), fontColor=color(NAVY), fontFamily=lit(SEGOE),
                             backColorPrimary=color(WHITE), backColorSecondary=color(WHITE)),
                       sel(props(fontColor={"solid": {"color": {"expr": {"Conditional": {
                           "Cases": [{"Condition": {"Comparison": {"ComparisonKind": 2, "Left": meas("Churn Rate"),
                                                                    "Right": {"Literal": {"Value": "0.35D"}}}},
                                      "Value": {"Literal": {"Value": f"'{ORANGE}'"}}}],
                           "DefaultValue": {"Literal": {"Value": f"'{NAVY}'"}}}}}}}),
                           {"data": [{"dataViewWildcard": {"matchingOption": 1}}], "metadata": rate_ref})],
            "columnHeaders": [props(fontSize=dnum(9), fontColor=color(TEXT2), fontFamily=lit(SEGOE),
                                    backColor=color(WHITE), columnAdjustment=s("growToFit"),
                                    autoSizeColumnWidth=b(True))],
            "grid": [props(gridVertical=b(False), gridHorizontal=b(True), gridHorizontalColor=color(RULE),
                           rowPadding=dnum(3), outlineColor=color(RULE))],
            "total": [props(totals=b(False))],
            "columnFormatting": [data_bars],
        },
        "visualContainerObjects": {**chrome(title="Highest-churn states, top 10",
                                            alt="Table of the ten states with the highest churn rate, with data bars"),
                                   "stylePreset": [props(name=s("None"))]},
        "drillFilterOtherVisuals": True,
    }
    rate_order = {"Measure": {"Expression": {"SourceRef": {"Source": "m"}}, "Property": "Churn Rate"}}
    who.add("top_states", 858, 210, 381, 336, table,
            filters=[topn_filter("who.states.top10", "State", rate_order, n=10)])
    so_what(who, "so_what", 858, 554, 381, 96,
            "Build a senior retention offer and promote family and group bundles. Investigate California, "
            "where 63% of a small base left.")


def build_report():
    d = RP / "definition"
    if d.exists():
        shutil.rmtree(d)
    build_pages()
    for p in PAGES:
        pd_ = d / "pages" / p.name
        write_json(pd_ / "page.json", {
            "$schema": PAGE_SCHEMA, "name": p.name, "displayName": p.display,
            "displayOption": "FitToPage", "height": 720, "width": 1280,
            "objects": {"background": [props(color=color(WHITE), transparency=dnum(0))]},
        })
        for v in p.visuals:
            write_json(pd_ / "visuals" / v["name"] / "visual.json", v)
    write_json(d / "pages" / "pages.json", {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/pagesMetadata/1.1.0/schema.json",
        "pageOrder": [p.name for p in PAGES],
        "activePageName": PAGES[0].name,
    })
    write_json(d / "version.json", {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/versionMetadata/1.0.0/schema.json",
        "version": "2.0.0",
    })
    at_import = {"visual": "2.12.0", "report": "3.4.0", "page": "2.3.1"}
    write_json(d / "report.json", {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/report/3.3.0/schema.json",
        "themeCollection": {
            "baseTheme": {"name": BASE_THEME, "reportVersionAtImport": at_import, "type": "SharedResources"},
            "customTheme": {"name": THEME_FILE, "reportVersionAtImport": at_import, "type": "RegisteredResources"},
        },
        "objects": {
            "section": [props(verticalAlignment=s("Top"))],
            "outspacePane": [props(expanded=b(False))],
        },
        "resourcePackages": [
            {"name": "SharedResources", "type": "SharedResources",
             "items": [{"name": BASE_THEME, "path": f"BaseThemes/{BASE_THEME}.json", "type": "BaseTheme"}]},
            {"name": "RegisteredResources", "type": "RegisteredResources",
             "items": [{"name": THEME_FILE, "path": THEME_FILE, "type": "CustomTheme"}]},
        ],
        "settings": {
            "useStylableVisualContainerHeader": True,
            "exportDataMode": "AllowSummarized",
            "defaultDrillFilterOtherVisuals": True,
            "allowChangeFilterTypes": True,
            "useEnhancedTooltips": True,
            "useDefaultAggregateDisplayName": True,
        },
    })
    res = RP / "StaticResources"
    if res.exists():
        shutil.rmtree(res)  # drop stale theme versions
    (res / "SharedResources" / "BaseThemes").mkdir(parents=True, exist_ok=True)
    (res / "RegisteredResources").mkdir(parents=True, exist_ok=True)
    shutil.copyfile(BASE_THEME_SRC, res / "SharedResources" / "BaseThemes" / f"{BASE_THEME}.json")
    theme = json.loads(THEME_SRC.read_text(encoding="utf-8"))
    theme["name"] = THEME_FILE
    write_json(res / "RegisteredResources" / THEME_FILE, theme)

    write_json(RP / "definition.pbir", {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definitionProperties/2.0.0/schema.json",
        "version": "4.0",
        "datasetReference": {"byPath": {"path": f"../{NAME}.SemanticModel"}},
    })
    write_json(RP / ".platform", {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/gitIntegration/platformProperties/2.0.0/schema.json",
        "metadata": {"type": "Report", "displayName": NAME},
        "config": {"version": "2.0", "logicalId": guid("logical.report")},
    })


def build_project():
    write_json(ROOT / f"{NAME}.pbip", {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/pbip/pbipProperties/1.0.0/schema.json",
        "version": "1.0",
        "artifacts": [{"report": {"path": f"{NAME}.Report"}}],
        "settings": {"enableAutoRecovery": True},
    })
    write(ROOT / ".gitignore", "**/.pbi/localSettings.json\n**/.pbi/cache.abf\n__pycache__/\n.claude/\n")


if __name__ == "__main__":
    build_model()
    build_report()
    build_project()
    n = sum(len(p.visuals) for p in PAGES)
    print(f"Wrote {NAME}: {len(PAGES)} pages, {n} visuals, {len(MEASURES)} measures")
