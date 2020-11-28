import random as r
from matplotlib import pyplot as plt
class Coupon():
    def __init__(self,id,value):
        self.id = id
        self.value = value
c_values= []
c_average = []
for n in range(100,4100,100):
    coupons = []
    while len(coupons) != n:
        coupon = Coupon(r.randrange(0,n),r.randrange(1,4001))
        temp = next((x for x in coupons if x.id == coupon.id), None)
        if temp != None:
            if temp.value > coupon.value:
                next(x for x in coupons if x.id == coupon.id).value = coupon.value
        else:
            coupons.append(coupon)
    c_values.append(sum(c.value for c in coupons))
    c_average.append(sum(c.value for c in coupons)/len(coupons))
n = [x for x in range(100,4100,100)]
plt.plot(n,c_values,label='V')
plt.plot(n,c_average,label='V/n')
plt.legend()
plt.xlabel('n')
plt.ylabel('Values')
plt.show()