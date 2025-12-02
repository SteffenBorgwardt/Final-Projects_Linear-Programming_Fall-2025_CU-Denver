#Setting our sets
set Car;
set Variables;

#setting our parameters where y is our response and x our predictors
param y {i in Car};
param x {i in Car, j in Variables};

# Our budget for L1 explained further down.
param t;

# This is just another way to write |b| like we have learened in class.
var bplus{j in Variables}  >= 0;
var bminus{j in Variables} >= 0;

minimize SSE:
    sum {i in Car}
        ( y[i] - sum {j in Variables} (bplus[j] - bminus[j]) * x[i,j] )^2;


# What our t is just simply our budget for the L1 norm, this means that we are 
# not doing a lambda penalty but instead contraining it by not alloiwng it to go past a limit.

subject to L1_budget:
    sum {j in Variables: j != 'intercept'} (bplus[j] + bminus[j]) <= t;
