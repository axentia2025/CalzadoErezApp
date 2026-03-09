"""
Intérprete de instrucciones en lenguaje natural usando Claude API.
Convierte texto del cliente en reglas estructuradas para el motor de distribución.
"""
import json
import os
from anthropic import Anthropic

SYSTEM_PROMPT = """Eres un asistente que interpreta instrucciones de distribución de calzado
para el sistema de Calzado Erez.

El sistema distribuye inventario desde un almacén central (Tienda 13) a múltiples tiendas.
Cada producto tiene: MARCA, MODELO (DESC1), COLOR (ATTR), SUBLINEA (categoría como
SANDALIA, BOTA, FORMAL, CASUAL, DEPORTIVO, TENIS, HUARACHE, etc.), TALLAS (22-30), y PRECIO.

Las tiendas se identifican por número (ej: Tienda 5, Tienda 33).

Tu trabajo es convertir instrucciones en español a reglas JSON estructuradas.

Tipos de reglas disponibles:

1. exclude_store_category — No enviar cierta categoría a una tienda
   Campos: store (int), category (string MAYUSCULAS)

2. exclude_store_size — No enviar cierta talla a una tienda
   Campos: store (int), sizes (array de int)

3. exclude_store_product — No enviar cierto producto/marca a una tienda
   Campos: store (int), brand (string|null), category (string|null)

4. exclude_store_all — No enviar nada a una tienda
   Campos: store (int)

5. prioritize_stores — Priorizar ciertas tiendas (reciben más)
   Campos: stores (array de int), category (string|null), boost_factor (float, default 2.0)

6. deprioritize_stores — Reducir envío a ciertas tiendas
   Campos: stores (array de int), category (string|null), penalty_factor (float, default 0.3)

7. max_per_size — Máximo pares por talla por tienda
   Campos: max_qty (int), stores (array de int o null para todas), category (string|null)

8. max_total_store — Máximo pares totales a una tienda
   Campos: max_qty (int), store (int), category (string|null)

9. process_first — Procesar cierta categoría primero
   Campos: category (string)

10. custom_note — Nota que no se puede estructurar como regla
    Campos: note (string)

REGLAS:
- Responde SOLAMENTE con un JSON array válido, sin markdown, sin explicaciones.
- Cada regla DEBE incluir "type" y "description_es" (descripción en español).
- Las categorías (SUBLINEA) siempre en MAYÚSCULAS.
- Si la instrucción es ambigua o no se puede mapear a una regla, usa "custom_note".
- Si una instrucción genera múltiples reglas, devuelve múltiples objetos en el array.
- Los números de tienda son enteros (ej: 33, no "33").

Ejemplo de respuesta:
[{"type": "exclude_store_category", "store": 33, "category": "SANDALIA", "description_es": "No enviar sandalias a tienda 33"}]
"""


def interpret_instruction(instruction):
    """Interpreta una instrucción en lenguaje natural y retorna reglas estructuradas.

    Args:
        instruction: Texto en español con la instrucción del cliente.

    Returns:
        Lista de dicts, cada uno representando una regla.

    Raises:
        ValueError: Si no se puede interpretar la instrucción.
        Exception: Si hay error de API.
    """
    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        raise ValueError(
            "No se encontró ANTHROPIC_API_KEY. "
            "Configura la variable de entorno para usar esta función."
        )

    client = Anthropic(api_key=api_key)

    response = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=1024,
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": instruction}],
    )

    text = response.content[0].text.strip()

    # Limpiar posible markdown wrapping
    if text.startswith("```"):
        lines = text.split("\n")
        lines = [l for l in lines if not l.startswith("```")]
        text = "\n".join(lines).strip()

    try:
        parsed = json.loads(text)
    except json.JSONDecodeError:
        # Fallback: crear nota personalizada
        return [{
            "type": "custom_note",
            "note": instruction,
            "description_es": f"Nota: {instruction}",
        }]

    if isinstance(parsed, dict):
        parsed = [parsed]

    if not isinstance(parsed, list):
        return [{
            "type": "custom_note",
            "note": instruction,
            "description_es": f"Nota: {instruction}",
        }]

    # Validar que cada regla tiene los campos mínimos
    valid_types = {
        "exclude_store_category", "exclude_store_size", "exclude_store_product",
        "exclude_store_all", "prioritize_stores", "deprioritize_stores",
        "max_per_size", "max_total_store", "process_first", "custom_note",
    }

    validated = []
    for rule in parsed:
        if not isinstance(rule, dict):
            continue
        if rule.get("type") not in valid_types:
            rule["type"] = "custom_note"
            rule["note"] = instruction
        if "description_es" not in rule:
            rule["description_es"] = instruction
        validated.append(rule)

    return validated if validated else [{
        "type": "custom_note",
        "note": instruction,
        "description_es": f"Nota: {instruction}",
    }]
