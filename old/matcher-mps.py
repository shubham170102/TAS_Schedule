from pulp import *

var, model = LpProblem.fromMPS("TAS.mps", LpMaximize)

#Solve
model.writeLP("TAS.lp")
model.solve()
print("Status:", LpStatus[model.status])
print("Objective value: ", value(model.objective))
