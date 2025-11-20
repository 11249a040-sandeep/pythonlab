import math

class Shape:
    def area(self):
        pass


class Rectangle(Shape):
    def __init__(self, length, breadth):
        self.length = length
        self.breadth = breadth

    def area(self):
        return self.length * self.breadth


class Triangle(Shape):
    def __init__(self, base, height):
        self.base = base
        self.height = height

    def area(self):
        return 0.5 * self.base * self.height


class Circle(Shape):
    def __init__(self, radius):
        self.radius = radius

    def area(self):
        return math.pi * (self.radius ** 2)


r = Rectangle(10, 5)
t = Triangle(6, 4)
c = Circle(3)

print("Area of Rectangle:", r.area())
print("Area of Triangle:", t.area())
print("Area of Circle:", round(c.area(), 2))
