import xloil as xlo


@xlo.converter()
def my_number(value):
    return value-1


@xlo.func()
def my_adder(value1: my_number, value2: my_number):
    return value1 + value2
