def clean_num(val, default=0.0):
            try:
                # Si ya es número desde Excel, lo retorna directamente
                if isinstance(val, (int, float)): return float(val)
                
                s = str(val).strip()
                # Elimina todo excepto números, puntos, comas y signos menos
                s = re.sub(r'[^\d.,-]', '', s)
                
                # Caso 1: Tiene puntos y comas (ej: 1.234.567,89) -> quita puntos, cambia coma por punto
                if '.' in s and ',' in s:
                    s = s.replace('.', '').replace(',', '.')
                # Caso 2: Solo tiene puntos y el último bloque tiene 3 dígitos (ej: 283.000) -> asume miles
                elif '.' in s and len(s.split('.')[-1]) == 3:
                    s = s.replace('.', '')
                # Caso 3: Solo tiene coma (ej: 283,50) -> cambia coma por punto (decimal)
                elif ',' in s:
                    s = s.replace(',', '.')
                    
                return float(s)
            except Exception:
                return default
