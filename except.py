a = 10
b = 0
c = 0
try:
    if b == 0:
        raise ArithmeticError
    else:
        print(a/b)


except ArithmeticError:
    print("b cannot be 0")

finally:
    print("DONEEE")
