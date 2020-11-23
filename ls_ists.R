library(lpSolveAPI)

stsslp <- function( 
     # Shortage of two-stage stochastic linear programming:
y,   # DPUT to be satisfied in period. 
p,   # probability of scenario. 
o,   # Fixed order cost in period.
h,   # Holding cost at the end of period.
s,   # Shortage cost at the end of period.
C,   # Budget purchase in period.
u    # Unitary cost of purchase in period.
) { 


### PRIMERA ETAPA
lps.1st <- make.lp(0, 2)                
set.objfn(lps.1st, c(o,u))
set.type(lps.1st, 1, "binary")
set.type(lps.1st, 2, "real")
xt <- c(-C,u)
add.constraint(lps.1st, xt, "<=", 0)
xt <- c(0,1)
add.constraint(lps.1st, xt, ">=", 1)

solve(lps.1st)
FO1st = get.objective(lps.1st)
 vars = get.variables(lps.1st)
    z = vars[1]
    q = vars[2]
Q = q

primal_dual = function(Q){

I = B = Pi1 = Pi2 = numeric()
for(i in 1:length(y)){
### PRIMAL
lps.primal <- make.lp(0, 2)  
set.objfn(lps.primal, c(h,s))
lps.primal
xt <- c(-1,1)
add.constraint(lps.primal, xt, "=", y[i] - Q)

solve(lps.primal)
FOoprimal = get.objective(lps.primal)
     vars = get.variables(lps.primal)
    iprimal = vars[1]
    bprimal = vars[2]
I = c(I, iprimal)
B = c(B, bprimal)

### DUAL
lps.dual <- make.lp(0, 2) 
set.objfn(lps.dual, -c(0,y[i]-Q))
set.bounds(lps.dual, lower = rep(-Inf, 2), upper = rep(Inf, 2))
xt <- c(1,0)
add.constraint(lps.dual, xt, "<=", -h)
xt <- c(0,1)
add.constraint(lps.dual, xt, "<=", s)

solve(lps.dual)
FOdual = get.objective(lps.dual)
     vars = get.variables(lps.dual)
    pi1 = vars[1]
    pi2 = vars[2]
 Pi1 = c(Pi1, pi1)
 Pi2 = c(Pi2, pi2)
}

e = sum(Pi2*p*y)
E = sum(Pi2*p)
W = e - E*q
return(list("e" = e, "E" = E, "W" = W))}



##### i-th stage



fac = primal_dual(q)
e = fac$e
E = fac$E
W = fac$W

W
E
e


lps.ist <- make.lp(0, 3)                
set.objfn(lps.ist, c(o,u,1))
set.type(lps.ist, 1, "binary")
xt <- c(-C,u,0)
add.constraint(lps.ist, xt, "<=", 0)
xt <- c(0,E,1)
add.constraint(lps.ist, xt, ">=", e)
set.bounds(lps.ist, 
lower = c(0,1,-Inf),upper = c(1,Inf,0))

solve(lps.ist)
FOist = get.objective(lps.ist)
     vars = get.variables(lps.ist)
        z = vars[1]
        Q = vars[2]
    theta = vars[3]
z
Q
theta

I = B = Pi1 = Pi2 = FOP = numeric()
for(i in 1:length(y)){
### PRIMAL
lps.primal <- make.lp(0, 2)  
set.objfn(lps.primal, c(h,s))
lps.primal
xt <- c(-1,1)
add.constraint(lps.primal, xt, "=", y[i] - Q)
solve(lps.primal)
FOoprimal = get.objective(lps.primal)
     vars = get.variables(lps.primal)
    iprimal = vars[1]
    bprimal = vars[2]
FOP = c(FOP, FOoprimal)
I = c(I, iprimal)
B = c(B, bprimal)
}

for( i in 1:length(y)){
lps.dual <- make.lp(0, 2) 
set.objfn(lps.dual, -c(0,y[i]-Q))
set.bounds(lps.dual, lower = rep(-Inf, 2), upper = rep(Inf, 2))
xt <- c(1,0)
add.constraint(lps.dual, xt, "<=", -h)
xt <- c(0,1)
add.constraint(lps.dual, xt, "<=", s)

solve(lps.dual)
FOdual = get.objective(lps.dual)
     vars = get.variables(lps.dual)
    pi1 = vars[1]
    pi2 = vars[2]
 Pi1 = c(Pi1, pi1)
 Pi2 = c(Pi2, pi2)
}
Pi1[1:length(y)] = Pi1[length(y)]
Pi2[1:length(y)] = Pi2[length(y)]

e = sum(Pi2*p*y)
E = sum(Pi2*p)
W = e - E*Q

return(list("L_shaped_results" = c("W" = round(W,4), "Theta" = round(theta,4)), 
      "Inv_mod_results" = c("z" = z, "Q" = Q, "I" = mean(I), "B" = mean(B)), 
      "FOprimal" = c("FO_primal" = FOP), "FOTC" = c(FOist), "I" = I, "B" = B ))
}



