from sympy import Eq, Symbol, solve

y = Symbol('y')
eqn = Eq(y*(8.0 - y**3), 8.0)

print (solve(eqn))