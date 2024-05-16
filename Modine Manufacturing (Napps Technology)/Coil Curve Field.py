import numpy as np
import matplotlib.pyplot as plt

fig = plt.figure()

"""Co Current Temps"""

eq40 = lambda x: (-0.0425*(x**4)) + 0.7959*(x**3) - 5.4438*(x**2) + 11.542*x + 23.455
eq35 = lambda x: (-0.164*(x**4)) + 2.6235*(x**3) - 14.806*(x**2) + 30.016*x + 7.4071
eq30 = lambda x: (-0.2311*(x**4)) + 3.2692*(x**3) - 16.717*(x**2) + 31.769*x + 2.1571
eq25 = lambda x: (-0.1729*(x**4)) + 1.9847*(x**3) - 9.3729*(x**2) + 16.855*x + 6.25
eq20 = lambda x: (-2*(x**3)) + 9.4*(x**2) - 15.2*x + 18.5
eq15 = lambda x: (-1.9167*(x**3)) + 10.15*(x**2) - 18.233*x + 15.6
zout = [eq15,eq20,eq25,eq30,eq35,eq40]

x = [50000,75000,100000,125000,150000,175000,200000,225000]
y = [15,20,25,30,35,40]
plt.xlabel('Ambient [°F]')
plt.ylabel('Condenser Leaving Water [°F]')
# plt.title('Heat Pump Chiller Operation Limits')
plt.title('Heat Pump Chiller Predicted Performance and Operation Limits')
plt.xlim(-0.2,40)
plt.annotate('% of Nominal Capacity',(7.5,125))

x, y = np.meshgrid(x, y)
z = []
for eq in zout:
   zrow = []
   for val in x:
      zrow.append(eq(x))
   z.append(zrow)
print(z)




# CS1 = plt.contour(x, y, z,10, colors='black')
# plt.plot(x1,y1,"r-",label = 'Operation Limit')

# plt.clabel(CS1, inline=True, fontsize=8)
# plt.legend()

# plt.show()

