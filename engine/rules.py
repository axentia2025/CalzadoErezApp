"""
Motor de reglas personalizadas para distribución.
Almacena, carga y aplica reglas creadas desde instrucciones del cliente.
"""
import json
import os
import uuid
from collections import defaultdict
from datetime import datetime

DEFAULT_RULES_PATH = os.path.join(os.path.dirname(os.path.dirname(__file__)), "data", "rules.json")


class RuleStore:
    """Almacena y gestiona reglas de distribución en un archivo JSON."""

    def __init__(self, filepath=None):
        self.filepath = filepath or DEFAULT_RULES_PATH
        os.makedirs(os.path.dirname(self.filepath), exist_ok=True)
        self.rules = self._load()

    def _load(self):
        if os.path.exists(self.filepath):
            try:
                with open(self.filepath, "r", encoding="utf-8") as f:
                    return json.load(f)
            except (json.JSONDecodeError, IOError):
                return []
        return []

    def _save(self):
        with open(self.filepath, "w", encoding="utf-8") as f:
            json.dump(self.rules, f, ensure_ascii=False, indent=2)

    def add_rule(self, rule):
        """Agrega una regla y retorna su ID."""
        rule["id"] = str(uuid.uuid4())[:8]
        rule["created_at"] = datetime.now().isoformat()
        rule.setdefault("active", True)
        self.rules.append(rule)
        self._save()
        return rule["id"]

    def delete_rule(self, rule_id):
        """Elimina una regla por su ID."""
        self.rules = [r for r in self.rules if r.get("id") != rule_id]
        self._save()

    def get_active_rules(self):
        """Retorna solo las reglas activas."""
        return [r for r in self.rules if r.get("active", True)]

    def get_all_rules(self):
        """Retorna todas las reglas."""
        return list(self.rules)


# ---------------------------------------------------------------------------
# Funciones de aplicación de reglas (una por fase del algoritmo)
# ---------------------------------------------------------------------------

def _match_category(rule_category, product_sublinea):
    """Compara categoría de regla con sublínea de producto (case-insensitive)."""
    if not rule_category:
        return True
    return rule_category.upper() in product_sublinea.upper()


def _match_brand(rule_brand, product_brand):
    """Compara marca de regla con marca de producto (case-insensitive)."""
    if not rule_brand:
        return True
    return rule_brand.upper() in product_brand.upper()


def apply_pre_filter_rules(rules, almacen_por_producto, tiendas_info, tiendas_ordenadas):
    """Fase 1: Excluir tiendas, productos o tallas antes de la distribución.

    Procesa tipos: exclude_store_all, exclude_store_category,
                   exclude_store_size, exclude_store_product.

    Modifica tiendas_info in-place y retorna tiendas_ordenadas filtrada.
    """
    filtered_tiendas = list(tiendas_ordenadas)

    for rule in rules:
        rtype = rule.get("type", "")

        if rtype == "exclude_store_all":
            store = rule.get("store")
            if store and store in filtered_tiendas:
                filtered_tiendas.remove(store)

        elif rtype == "exclude_store_category":
            store = rule.get("store")
            category = rule.get("category", "")
            if store and category:
                # Eliminar productos de esa categoría para esa tienda
                if store in tiendas_info:
                    prods_to_remove = []
                    for prod_key in tiendas_info[store]:
                        prod_info = almacen_por_producto.get(prod_key)
                        if prod_info and _match_category(category, prod_info.get("sublinea", "")):
                            prods_to_remove.append(prod_key)
                    for pk in prods_to_remove:
                        del tiendas_info[store][pk]

        elif rtype == "exclude_store_size":
            store = rule.get("store")
            sizes = rule.get("sizes", [])
            if store and sizes and store in tiendas_info:
                for prod_key in tiendas_info[store]:
                    for size in sizes:
                        if size in tiendas_info[store][prod_key]:
                            del tiendas_info[store][prod_key][size]

        elif rtype == "exclude_store_product":
            store = rule.get("store")
            brand = rule.get("brand")
            category = rule.get("category")
            if store and store in tiendas_info:
                prods_to_remove = []
                for prod_key in tiendas_info[store]:
                    prod_marca = prod_key[0]
                    prod_info = almacen_por_producto.get(prod_key)
                    prod_sublinea = prod_info.get("sublinea", "") if prod_info else ""
                    if _match_brand(brand, prod_marca) and _match_category(category, prod_sublinea):
                        prods_to_remove.append(prod_key)
                for pk in prods_to_remove:
                    del tiendas_info[store][pk]

    return filtered_tiendas


def apply_scoring_rules(rules, necesidades, almacen_por_producto):
    """Fase 2: Modificar scores para priorizar o depriorizar tiendas.

    Procesa tipos: prioritize_stores, deprioritize_stores.
    Modifica necesidades in-place.
    """
    for rule in rules:
        rtype = rule.get("type", "")

        if rtype == "prioritize_stores":
            stores = set(rule.get("stores", []))
            category = rule.get("category")
            boost = rule.get("boost_factor", 2.0)

            for prod_key, tiendas_list in necesidades.items():
                prod_info = almacen_por_producto.get(prod_key)
                if category and prod_info:
                    if not _match_category(category, prod_info.get("sublinea", "")):
                        continue
                for nec in tiendas_list:
                    if nec["tienda"] in stores:
                        nec["score"] *= boost
                # Re-sort by score
                tiendas_list.sort(key=lambda x: -x["score"])

        elif rtype == "deprioritize_stores":
            stores = set(rule.get("stores", []))
            category = rule.get("category")
            penalty = rule.get("penalty_factor", 0.3)

            for prod_key, tiendas_list in necesidades.items():
                prod_info = almacen_por_producto.get(prod_key)
                if category and prod_info:
                    if not _match_category(category, prod_info.get("sublinea", "")):
                        continue
                for nec in tiendas_list:
                    if nec["tienda"] in stores:
                        nec["score"] *= penalty
                tiendas_list.sort(key=lambda x: -x["score"])


def apply_ordering_rules(rules, necesidades, almacen_por_producto):
    """Fase 3: Reordenar procesamiento de productos.

    Procesa tipo: process_first.
    Retorna lista ordenada de prod_keys.
    """
    first_categories = []
    for rule in rules:
        if rule.get("type") == "process_first":
            cat = rule.get("category", "")
            if cat:
                first_categories.append(cat.upper())

    if not first_categories:
        return None

    priority_keys = []
    normal_keys = []

    for prod_key in necesidades:
        prod_info = almacen_por_producto.get(prod_key)
        sublinea = prod_info.get("sublinea", "").upper() if prod_info else ""
        is_priority = any(cat in sublinea for cat in first_categories)
        if is_priority:
            priority_keys.append(prod_key)
        else:
            normal_keys.append(prod_key)

    return priority_keys + normal_keys


def apply_distribution_rules(rules, tienda, talla, qty, prod_key,
                              almacen_por_producto, store_running_totals):
    """Fase 4: Aplicar límites de cantidad durante la distribución.

    Procesa tipos: max_per_size, max_total_store.
    Retorna qty ajustada (puede ser 0).
    """
    for rule in rules:
        rtype = rule.get("type", "")

        if rtype == "max_per_size":
            max_qty = rule.get("max_qty", 999)
            rule_stores = rule.get("stores")  # None = todas
            category = rule.get("category")

            # Verificar si aplica a esta tienda
            if rule_stores and tienda not in rule_stores:
                continue
            # Verificar si aplica a esta categoría
            if category:
                prod_info = almacen_por_producto.get(prod_key)
                if prod_info and not _match_category(category, prod_info.get("sublinea", "")):
                    continue
            qty = min(qty, max_qty)

        elif rtype == "max_total_store":
            store = rule.get("store")
            max_total = rule.get("max_qty", 9999)
            category = rule.get("category")

            if store != tienda:
                continue
            if category:
                prod_info = almacen_por_producto.get(prod_key)
                if prod_info and not _match_category(category, prod_info.get("sublinea", "")):
                    continue

            current = store_running_totals.get(tienda, 0)
            remaining_capacity = max(0, max_total - current)
            qty = min(qty, remaining_capacity)

    return qty


def apply_post_filter_rules(rules, distribuciones):
    """Fase 5: Filtrar asignaciones que violen reglas (safety net).

    Actualmente se usa como respaldo — la mayoría de reglas se aplican
    en fases anteriores. Retorna distribuciones filtrada.
    """
    # Las reglas de exclusión ya se aplican en pre-filter, pero por seguridad
    # verificamos aquí también
    exclude_pairs = set()  # (store, category) pairs to exclude
    for rule in rules:
        if rule.get("type") == "exclude_store_category":
            store = rule.get("store")
            category = rule.get("category", "").upper()
            if store and category:
                exclude_pairs.add((store, category))

    if not exclude_pairs:
        return distribuciones

    filtered = []
    for d in distribuciones:
        skip = False
        for store, category in exclude_pairs:
            if d["TIENDA"] == store and category in d.get("SUBLINEA", "").upper():
                skip = True
                break
        if not skip:
            filtered.append(d)

    return filtered
