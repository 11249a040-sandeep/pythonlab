marks = []
for i in range(3):
    marks1 = int(input("Enter the marks: "))
    marks.append(marks1)
print(f"The marks list is: {marks}")
avg1 = (marks[0] + marks[1]) / 2
avg2 = (marks[0] + marks[2]) / 2
avg3 = (marks[1] + marks[2]) / 2
avg_marks = [avg1, avg2, avg3]
print(f"Average of marks list is: {avg_marks}")
avg_marks.remove(min(avg_marks))
print(f"Max of 2 avg marks are: {avg_marks}")
