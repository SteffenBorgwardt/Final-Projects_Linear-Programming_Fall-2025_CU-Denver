# ---- Sets & data ----
set V ordered;                 # vertices, e.g., 1..n
set KSETS;                     # index of all k-sets
set S within KSETS;            # the "independent" family
param inc {KSETS, V} binary;   # 1 if vertex v ∈ U, else 0

# Undirected edge index
set E := {i in V, j in V: i < j};

# ---- Variables ----
# Edges (use binary for an actual graph; relax to [0,1] for LP)
var x {E} binary;

# Weights / threshold margin (shifted test)
var w {V} >= 0, <= 1;
var t >= 0, <= 5;
var delta >= 0;



# ---- Graph side: exact independent k-sets = S ----
# All edges inside every F in S must be zero
s.t. indep_F {U in S}:
    sum { (i,j) in E: inc[U,i] = 1 and inc[U,j] = 1 } x[i,j] = 0;

# Every other k-set must contain at least one edge
s.t. nonindep_U {U in KSETS diff S}:
    sum { (i,j) in E: inc[U,i] = 1 and inc[U,j] = 1 } x[i,j] >= 1;

# ---- Shifted / k-threshold side on the same S ----
s.t. pos {U in S}:
    sum {v in V} inc[U,v] * w[v] <= t - delta;

s.t. neg {U in KSETS diff S}:
    sum {v in V} inc[U,v] * w[v] >= t + delta;

# ---- Objective ----
maximize margin: delta;











