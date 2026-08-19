def determinar_tipo_ariba(row):
    """
    Clasifica la solicitud integrando los flujos SAP y Ariba según reglas unificadas:
    - Serie 1 (100) y Serie 5 (500):
        * Encargado Cesar -> ⚪ SAP MRP
        * Otro encargado -> ⚪ SAP ERP
    - Serie 6 (600):
        * Sin material o registrada en Trazabilidad -> 🔵 ARIBA NO CATALOGADA
        * Con código de material / catálogo -> 🟢 ARIBA DIRECTA / CATALOGADA
    """
    # 1. Respetar clasificación previa explícita si existe
    for col in ["Tipo Ariba", "Tipo_Ariba", "Origen Ariba", "Origen", "Tipo Flujo"]:
        if col in row and pd.notna(row[col]) and str(row[col]).strip() != "":
            val = str(row[col]).upper()
            if "NO CATALOGAD" in val or "NOCATALOGAD" in val:
                return "🔵 ARIBA NO CATALOGADA"
            elif "CATALOGAD" in val or "DIRECTA" in val:
                return "🟢 ARIBA DIRECTA / CATALOGADA"
            elif "MRP" in val:
                return "⚪ SAP MRP"
            elif "ERP" in val:
                return "⚪ SAP ERP"

    sol = str(row.get("Solicitud de pedido", "")).strip()
    material = str(row.get("Material", "")).strip()
    tiene_material = bool(material and material.lower() not in ["nan", "none", "n/a", "-", "0"])

    # Obtener nombre del encargado/responsable desde la fila
    encargado = str(
        row.get("Responsable MRP", row.get("Encargado", row.get("Responsable", "")))
    ).upper()
    es_cesar = "CESAR" in encargado or "CÉSAR" in encargado

    # Serie 1 (100) y Serie 5 (500)
    if sol.startswith("1") or sol.startswith("5"):
        return "⚪ SAP MRP" if es_cesar else "⚪ SAP ERP"

    # Serie 6 (600)
    if sol.startswith("6"):
        if not tiene_material:
            return "🔵 ARIBA NO CATALOGADA"
        else:
            return "🟢 ARIBA DIRECTA / CATALOGADA"

    return "⚪ OTROS"
