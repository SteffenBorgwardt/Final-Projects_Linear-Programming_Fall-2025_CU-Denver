##############################################################
# Description:
#   This optimization model determines the maximum possible
#   value of the secondary objective to 'proj.mod'
##############################################################

# ===============================
# Sets
# ===============================
set DEPOTS;     # Depots where products are supplied
set STATIONS;   # Stations that receive shipments
set PRODUCTS;   # Types of products to be delivered
set TRUCKS;     # Available trucks for distribution

# ===============================
# Parameters
# ===============================
param distance {DEPOTS, STATIONS} >= 0;                 # Distance between depot and station, and back again (km)
param supply {DEPOTS, PRODUCTS} >= 0;                   # Available product supply at each depot
param capacity_station {STATIONS, PRODUCTS} >= 0;       # Current storage capacity used per station per product
param full_capacity_station {STATIONS, PRODUCTS} >= 0;  # Maximum storage capacity per station per product
param capacity_truck {TRUCKS, PRODUCTS} >= 0;           # Truck capacity per product type

param avg_speed {TRUCKS} >= 0, default 45;                       # Average truck speed (km/h)
param shift_duration {TRUCKS} >= 0, default 12;                  # Maximum driving hours per driver per shift
param load_unload_time {DEPOTS,STATIONS,TRUCKS} >= 0, default 2; # Fixed loading/unloading time at a given depot/station (hours)

# ===============================
# Decision Variables
# ===============================
var assign_truck {DEPOTS, STATIONS, TRUCKS} >= 0, integer; # number of trips made from DEPOT to STATION by TRUCK

# ===============================
# Objective Function
# ===============================

maximize Shipping_Efficiency:
    (sum {d in DEPOTS, s in STATIONS, t in TRUCKS} (distance[d,s] * assign_truck[d,s,t]));

# ===============================
# Constraints
# ===============================

# (1) Depot supply cannot be exceeded for any product
s.t. Depot_Supply_Limit {d in DEPOTS, p in PRODUCTS}:
    sum {s in STATIONS, t in TRUCKS} capacity_truck[t,p] 
    * assign_truck[d,s,t] <= supply[d,p];

# (2) Station storage capacity cannot be exceeded for any product
s.t. Station_Capacity_Limit {s in STATIONS, p in PRODUCTS}:
    sum {d in DEPOTS} sum {t in TRUCKS} capacity_truck[t,p] 
    * assign_truck[d,s,t] <= full_capacity_station[s,p] - capacity_station[s,p];

# (3) Each truck's total working time (travel + load/unload) must not exceed the driver shift limit
s.t. Truck_Shift_Limit {t in TRUCKS}:
    sum {d in DEPOTS, s in STATIONS}
        (load_unload_time[d,s,t] + (distance[d,s] / avg_speed[t])) * assign_truck[d,s,t]
        <= shift_duration[t];