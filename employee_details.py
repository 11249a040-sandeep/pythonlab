class Employee:
    def __init__(self, emp_id, name, dept, salary):
        self.id = emp_id
        self.name = name
        self.dept = dept
        self.salary = salary

    def update_salary(self, new_salary):
        self.salary = new_salary


emp = Employee(1, "neeraj", "AI&DS", 10000)
emp.update_salary(20000)
print(emp.id, emp.name, emp.dept, emp.salary)
