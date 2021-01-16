class MethodValueDemo:
    def __init__(self):
        self.method_variable_string = 'Default'
        self.method_variable_int = 100
        self.method_variable_float = 1.01

    def get_method_value(self):
        print(f"string value is: {self.method_variable_string}")
        print(f"int value is: {self.method_variable_int}")
        print(f"float value is: {self.method_variable_float}")

    def set_method_value(self, str_value, int_value, float_value):
        self.method_variable_string = str_value
        self.method_variable_int = int_value
        self.method_variable_float = float_value

    def caculation(self):
        result = self.method_variable_int * self.method_variable_float
        return self.method_variable_string, result




MD = MethodValueDemo()
MD.get_method_value()
res = MD.caculation()
print(f"{res}")
MD.set_method_value(str_value="change", int_value=200, float_value=2.00)
MD.get_method_value()
res = MD.caculation()
print(f"{MD.caculation()}")


