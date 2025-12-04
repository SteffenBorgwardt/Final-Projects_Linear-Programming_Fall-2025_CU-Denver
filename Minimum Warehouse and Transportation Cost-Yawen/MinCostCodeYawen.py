# -*- coding: utf-8 -*-
"""
Outbound logistics MILP with plant capacity in NUMBER OF ORDERS
and parametric sensitivity analysis.

Model:
  I aggregate orders by Product ID. For each product k, we choose
  a (plant w, origin port p, carrier band c) that is feasible and has
  minimal total cost.

Decision variable:
  x[k,w,p,c] = 1 if product k uses plant w, port p, and carrier band c.

Objective:
  Minimize total logistics cost = plant handling cost + freight cost.

Capacity:
  "Daily Capacity" in Sheet1 is interpreted as MAX NUMBER OF ORDERS
  that a plant can process. I aggregate by product, so the capacity
  constraint is:

      sum_k (#orders of product k assigned to w) <= cap_factor * Capacity[w]

Default:
  cap_factor = 1.0 (100% capacity) in the baseline scenario.

Sensitivity analysis:
  - Capacity ±20%: cap_factor = 1.2, 0.8
  - Freight ±10%:  freight_factor = 1.1, 0.9
"""

from gurobipy import Model, GRB, quicksum
import xlrd
from collections import defaultdict

FILE_NAME = "Supply chain logisitcs problem.xls"

# ------------------------------------------------------------
# 1. Read data from Excel
# ------------------------------------------------------------

def read_data(filename=FILE_NAME):
    wb = xlrd.open_workbook(filename)

    # ---------- Sheet 0: Orders ----------
    # Columns: Product ID | Order ID | Unit quantity | Weight | Destination Port
    sh_orders = wb.sheet_by_index(0)
    hdr_o = {str(sh_orders.cell_value(0, j)).strip(): j
             for j in range(sh_orders.ncols)}

    col_prod = hdr_o["Product ID"]
    col_oid  = hdr_o["Order ID"]
    col_qty  = hdr_o["Unit quantity"]
    col_wgt  = hdr_o["Weight"]
    col_dest = hdr_o["Destination Port"]

    orders = []
    for r in range(1, sh_orders.nrows):
        prod = int(sh_orders.cell_value(r, col_prod))
        oid  = int(sh_orders.cell_value(r, col_oid))
        qty  = float(sh_orders.cell_value(r, col_qty))
        wgt  = float(sh_orders.cell_value(r, col_wgt))
        dest = str(sh_orders.cell_value(r, col_dest)).strip()
        orders.append((prod, oid, qty, wgt, dest))

    # Aggregate by product
    prod_units        = defaultdict(float)  # total unit quantity per product
    prod_weight       = defaultdict(float)  # total weight per product
    prod_dest         = {}                  # destination port per product
    prod_order_count  = defaultdict(int)    # number of orders per product

    for prod, oid, qty, wgt, dest in orders:
        prod_units[prod]       += qty
        prod_weight[prod]      += wgt
        prod_dest[prod]         = dest   # assume same dest per product
        prod_order_count[prod] += 1

    products = sorted(prod_units.keys())

    # ---------- Sheet 1: Plants ----------
    # Columns: Plant ID | Daily Capacity | Cost/unit
    sh_plants = wb.sheet_by_index(1)
    hdr_p = {str(sh_plants.cell_value(0, j)).strip(): j
             for j in range(sh_plants.ncols)}

    col_pid   = hdr_p["Plant ID"]
    col_cap   = hdr_p["Daily Capacity"]
    col_cost  = hdr_p["Cost/unit"]

    plants          = []
    plant_capacity  = {}   # interpreted as "max number of orders per day"
    plant_unit_cost = {}   # cost per unit quantity

    for r in range(1, sh_plants.nrows):
        w    = str(sh_plants.cell_value(r, col_pid)).strip()
        cap  = float(sh_plants.cell_value(r, col_cap))
        cost = float(sh_plants.cell_value(r, col_cost))
        plants.append(w)
        plant_capacity[w]  = cap
        plant_unit_cost[w] = cost

    # ---------- Sheet 2: Product -> Plant mapping ----------
    # Columns: Product ID | Plant Code
    sh_pp = wb.sheet_by_index(2)
    hdr_pp = {str(sh_pp.cell_value(0, j)).strip(): j
              for j in range(sh_pp.ncols)}

    col_pp_prod  = hdr_pp["Product ID"]
    col_pp_plant = hdr_pp["Plant Code"]

    prod_plants = defaultdict(set)
    for r in range(1, sh_pp.nrows):
        prod = int(sh_pp.cell_value(r, col_pp_prod))
        w    = str(sh_pp.cell_value(r, col_pp_plant)).strip()
        prod_plants[prod].add(w)

    # ---------- Sheet 4: Plant -> Port mapping ----------
    # Columns: Plant Code | Port
    sh_plantport = wb.sheet_by_index(4)
    hdr_port = {str(sh_plantport.cell_value(0, j)).strip(): j
                for j in range(sh_plantport.ncols)}

    col_plcode = hdr_port["Plant Code"]
    col_port   = hdr_port["Port"]

    plant_ports = defaultdict(set)
    for r in range(1, sh_plantport.nrows):
        w = str(sh_plantport.cell_value(r, col_plcode)).strip()
        p = str(sh_plantport.cell_value(r, col_port)).strip()
        plant_ports[w].add(p)

    # ---------- Sheet 3: Carrier bands ----------
    # Columns:
    #   Carrier | orig_port_cd | minm_wgh_qty | max_wgh_qty | svc_cd |
    #   minimum cost | rate | mode_dsc | tpt_day_cnt | Carrier type | dest_port_cd
    sh_car = wb.sheet_by_index(3)
    hdr_c = {str(sh_car.cell_value(0, j)).strip(): j
             for j in range(sh_car.ncols)}

    col_carrier = hdr_c["Carrier"]
    col_orig    = hdr_c["orig_port_cd"]
    col_minw    = hdr_c["minm_wgh_qty"]
    col_maxw    = hdr_c["max_wgh_qty"]
    col_svc     = hdr_c["svc_cd"]
    col_mincost = hdr_c["minimum cost"]
    col_rate    = hdr_c["rate"]
    col_ctype   = hdr_c["Carrier type"]
    col_destc   = hdr_c["dest_port_cd"]

    carrier_bands = []
    # band_info[c] = (carrier_name, carrier_type, orig_port,
    #                 minW, maxW, minCost, rate, dest_port, svc_cd)
    band_info = {}

    for r in range(1, sh_car.nrows):
        c_idx = r  # one band per row

        carrier_name = str(sh_car.cell_value(r, col_carrier)).strip()
        orig         = str(sh_car.cell_value(r, col_orig)).strip()
        minW         = float(sh_car.cell_value(r, col_minw))
        maxW         = float(sh_car.cell_value(r, col_maxw))
        svc          = str(sh_car.cell_value(r, col_svc)).strip()
        minCost      = float(sh_car.cell_value(r, col_mincost))
        rate         = float(sh_car.cell_value(r, col_rate))
        ctype        = str(sh_car.cell_value(r, col_ctype)).strip()
        dest         = str(sh_car.cell_value(r, col_destc)).strip()

        carrier_bands.append(c_idx)
        band_info[c_idx] = (
            carrier_name, ctype, orig,
            minW, maxW, minCost, rate, dest, svc
        )

    return (orders, products, prod_units, prod_weight, prod_dest,
            plants, plant_capacity, plant_unit_cost,
            prod_plants, plant_ports,
            carrier_bands, band_info,
            prod_order_count)


# ------------------------------------------------------------
# 2. Build feasible (product, plant, port, carrier) routes
# ------------------------------------------------------------

def build_candidates(products, prod_units, prod_weight, prod_dest,
                     prod_plants, plant_ports,
                     carrier_bands, band_info,
                     plant_unit_cost):
    """
    Enumerate all feasible product-plant-port-carrier combinations and
    compute their base cost components.

    Returns
    -------
    candidates : list of (k, w, p, c)
    fixed_cost : dict[(k,w,p,c)] = plant handling + minimum freight charge
    var_cost   : dict[(k,w,p,c)] = rate * total_weight_k
    """

    candidates = []
    fixed_cost = {}
    var_cost   = {}

    for k in products:
        units_k  = prod_units[k]
        weight_k = prod_weight[k]
        dest_k   = prod_dest[k]

        for w in prod_plants.get(k, []):
            for p in plant_ports.get(w, []):
                for c in carrier_bands:
                    (carrier_name, ctype, orig,
                     minW, maxW, minCost, rate, dest_c, svc) = band_info[c]

                    # Feasibility: origin port match, destination port match,
                    # and product weight <= band's max weight.
                    if orig != p:
                        continue
                    if dest_c != dest_k:
                        continue
                    if weight_k > maxW:
                        continue

                    idx = (k, w, p, c)

                    # Plant handling cost for all units of this product
                    plant_c = plant_unit_cost[w] * units_k

                    # Fixed part: plant cost + minimum freight charge
                    fixed_cost[idx] = plant_c + minCost

                    # Variable part: rate * total weight (used for freight sensitivity)
                    var_cost[idx] = rate * weight_k

                    candidates.append(idx)

    return candidates, fixed_cost, var_cost


# ------------------------------------------------------------
# 3. MILP solver (capacity in number of orders)
# ------------------------------------------------------------

def solve_model(cap_factor=1.0, freight_factor=1.0, verbose=False):
    """
    Solve the outbound logistics model with given capacity and freight factors.

    Parameters
    ----------
    cap_factor : float
        Capacity multiplier:
          1.0  -> original capacity (100%)
          1.2  -> +20% capacity
          0.8  -> -20% capacity
    freight_factor : float
        Multiplier for the variable freight cost part (rate * weight),
        e.g. 1.10 for +10% rates, 0.90 for -10%.
    verbose : bool
        If True, print Gurobi's solver log.

    Returns
    -------
    total_cost : float or None
        Optimal objective value (None if infeasible).
    chosen_routes : dict
        chosen_routes[k] = (w, p, c) for each product k.
    """
    (orders, products, prod_units, prod_weight, prod_dest,
     plants, plant_capacity, plant_unit_cost,
     prod_plants, plant_ports,
     carrier_bands, band_info,
     prod_order_count) = read_data()

    # Build feasible candidates
    candidates, fixed_cost, var_cost = build_candidates(
        products, prod_units, prod_weight, prod_dest,
        prod_plants, plant_ports,
        carrier_bands, band_info,
        plant_unit_cost
    )

    # Quick check: each product must have at least one candidate
    bad_products = [k for k in products
                    if not any(idx[0] == k for idx in candidates)]
    if bad_products:
        print("WARNING: Some products have no feasible route:", bad_products)
        return None, {}

    m = Model("OutboundLogistics_withCapacity")
    m.Params.OutputFlag = 1 if verbose else 0

    # Decision variables
    x = m.addVars(candidates, vtype=GRB.BINARY, name="x")

    # Objective: fixed_cost + freight_factor * var_cost
    m.setObjective(
        quicksum((fixed_cost[idx] + freight_factor * var_cost[idx]) * x[idx]
                 for idx in candidates),
        GRB.MINIMIZE
    )

    # Assignment: each product k chooses exactly one route
    for k in products:
        idx_k = [idx for idx in candidates if idx[0] == k]
        m.addConstr(quicksum(x[idx] for idx in idx_k) == 1,
                    name=f"assign_{k}")

    # Capacity constraints in NUMBER OF ORDERS 
    for w in plants:
        idx_w = [idx for idx in candidates if idx[1] == w]
        if idx_w and plant_capacity[w] > 0:
            cap_eff = cap_factor * plant_capacity[w]
            m.addConstr(
                quicksum(prod_order_count[idx[0]] * x[idx] for idx in idx_w)
                <= cap_eff,
                name=f"capacity_{w}"
            )

    m.optimize()

    if m.Status not in (GRB.OPTIMAL, GRB.SUBOPTIMAL):
        print(f"Model did not solve to optimality. Status = {m.Status}")
        if m.Status == GRB.INFEASIBLE:
            print("  -> Model infeasible under these capacity settings.")
        return None, {}

    total_cost = m.ObjVal

    # Extract chosen route per product
    chosen_routes = {}
    for idx in candidates:
        if x[idx].X > 0.5:
            k, w, p, c = idx
            chosen_routes[k] = (w, p, c)

    return total_cost, chosen_routes


# ------------------------------------------------------------
# 4. Sensitivity analysis
# ------------------------------------------------------------

def run_scenario(label, cap_factor=1.0, freight_factor=1.0, verbose=False):
    """
    Run one scenario and print a short summary.
    """
    print(f"\n--- {label} ---")
    cost, routes = solve_model(
        cap_factor=cap_factor,
        freight_factor=freight_factor,
        verbose=verbose
    )
    if cost is None:
        print(f"{label}: infeasible.")
    else:
        print(f"{label}: total cost = {cost:,.2f}")
        print(f"{label}: number of products assigned = {len(routes)}")
    return cost, routes


if __name__ == "__main__":
    # Baseline: 100% capacity, freight at nominal level
    base_cost, base_routes = run_scenario(
        "Baseline (Capacity 100%, Freight 100%)",
        cap_factor=1.0,
        freight_factor=1.0
    )

    if base_cost is not None:
        # Capacity sensitivity: ±20% capacity, same freight
        cap120_cost, _ = run_scenario(
            "Capacity +20% (120%), Freight 100%",
            cap_factor=1.2,
            freight_factor=1.0
        )
        cap080_cost, _ = run_scenario(
            "Capacity -20% (80%), Freight 100%",
            cap_factor=0.8,
            freight_factor=1.0
        )

        # Freight sensitivity: ±10% freight, capacity fixed at 100%
        fr110_cost, _ = run_scenario(
            "Capacity 100%, Freight +10% (110%)",
            cap_factor=1.0,
            freight_factor=1.1
        )
        fr090_cost, _ = run_scenario(
            "Capacity 100%, Freight -10% (90%)",
            cap_factor=1.0,
            freight_factor=0.9
        )

        # Report deltas vs baseline (only for feasible scenarios)
        def report_delta(label, c):
            if c is not None:
                delta_abs = c - base_cost
                delta_rel = (c / base_cost - 1.0) * 100.0
                print(f"{label}: Δ cost vs baseline = {delta_abs:,.2f} ({delta_rel:+.2f}% )")

        print("\n=== Cost change relative to baseline (Capacity 100%, Freight 100%) ===")
        report_delta("Capacity +20%", cap120_cost)
        report_delta("Capacity -20%", cap080_cost)
        report_delta("Freight +10%",  fr110_cost)
        report_delta("Freight -10%",  fr090_cost)
    else:
        print("Baseline scenario infeasible; sensitivity results not computed.")

