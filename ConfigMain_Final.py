import logic


hello = logic.display_instance(input("What is the name of the file:: ").replace(".xlsm","") + ".xlsm")
hello.configGen()


